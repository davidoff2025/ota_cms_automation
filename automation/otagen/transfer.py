"""FTP download + S3 upload — framework-independent, server-portable.

Ported from v3's worker logic with two deliberate changes for the web/AWS target:
  * No Tkinter / log-queue: progress and logs go through a simple callback.
  * No `aws` CLI subprocess: S3 upload uses boto3 (works on EC2/ECS instance
    roles, env creds, or ~/.aws — nothing extra to install in the container).

The FTP URL parsing mirrors v3 exactly, including the decoded-vs-encoded path
fallback that real Vidaa/Hisense FTP servers require.
"""
from __future__ import annotations

import os
import shutil
from dataclasses import dataclass, field
from ftplib import FTP
from typing import Callable
from urllib.parse import urlparse, unquote

from . import config

LogFn = Callable[[str], None]
ProgressFn = Callable[[float, str], None]


def _noop_log(_: str) -> None:
    pass


def _noop_progress(_pct: float, _msg: str) -> None:
    pass


@dataclass
class TransferResult:
    downloaded: list[str] = field(default_factory=list)
    uploaded: list[str] = field(default_factory=list)
    failed: list[str] = field(default_factory=list)
    s3_prefix: str = ""


# --------------------------------------------------------------------- FTP
def parse_ftp_url(ftp_url: str):
    """Parse an FTP ZIP URL into connection parts (decoded + encoded path).

    Returns (host, remote_dir, filename, remote_path, url_user,
             encoded_remote_dir, encoded_filename, encoded_remote_path).
    """
    raw_url = str(ftp_url).strip()
    parsed = urlparse(raw_url)
    host = parsed.hostname
    encoded_remote_path = parsed.path or ""
    remote_path = unquote(encoded_remote_path)
    url_user = unquote(parsed.username or "")
    if not host or not remote_path:
        raise ValueError(f"无效 FTP 地址：{ftp_url}")
    filename = os.path.basename(remote_path)
    remote_dir = os.path.dirname(remote_path) or "/"
    encoded_filename = os.path.basename(encoded_remote_path)
    encoded_remote_dir = os.path.dirname(encoded_remote_path) or "/"
    if not filename.lower().endswith(".zip"):
        raise ValueError(f"FTP 地址不是 .zip 文件：{ftp_url}")
    return (host, remote_dir, filename, remote_path, url_user,
            encoded_remote_dir, encoded_filename, encoded_remote_path)


def download_ftp_zip(
    ftp_url: str,
    dest_dir: str,
    *,
    file_index: int = 1,
    total_files: int = 1,
    log: LogFn = _noop_log,
    progress: ProgressFn = _noop_progress,
) -> str:
    """Download one FTP .zip into dest_dir, returning the local path."""
    (host, remote_dir, filename, remote_path, url_user,
     encoded_remote_dir, _encoded_filename, _encoded_remote_path) = parse_ftp_url(ftp_url)

    os.makedirs(dest_dir, exist_ok=True)
    local_path = os.path.join(dest_dir, filename)
    login_user = url_user or config.FTP_USER

    log(f"下载任务 {file_index}/{total_files}：{filename}")
    log(f"FTP 服务器：{host}  目录：{remote_dir}")
    log(f"本地暂存：{local_path}  登录用户：{login_user}")

    ftp = None
    try:
        ftp = FTP(host, timeout=120)
        ftp.login(login_user, config.FTP_PASS)
        log("FTP 登录成功。")

        cwd_success = False
        for candidate_dir, label in [(remote_dir, "解码后目录"), (encoded_remote_dir, "URL原始编码目录")]:
            if not candidate_dir:
                continue
            try:
                ftp.cwd(candidate_dir)
                cwd_success = True
                log(f"切换目录成功（{label}）：{candidate_dir}")
                break
            except Exception as e:
                log(f"切换目录失败（{label}）：{e}")
        if not cwd_success:
            log("目录切换失败，改用完整远程路径直接 RETR。")

        retr_name = filename if cwd_success else remote_path
        try:
            file_size = ftp.size(retr_name) or 0
        except Exception:
            file_size = 0

        downloaded = 0
        with open(local_path, "wb") as f:
            def callback(data):
                nonlocal downloaded
                f.write(data)
                downloaded += len(data)
                if file_size > 0:
                    file_pct = downloaded / file_size * 100
                    total_pct = ((file_index - 1) + file_pct / 100) / total_files * 100
                    progress(total_pct, f"下载 {filename} {file_pct:.0f}%")
            ftp.retrbinary(f"RETR {retr_name}", callback, blocksize=1024 * 128)

        log(f"下载完成：{local_path}（{downloaded} bytes）")
        progress(file_index / total_files * 100, f"已完成 {file_index}/{total_files}")
        return local_path
    finally:
        if ftp is not None:
            try:
                ftp.quit()
            except Exception:
                try:
                    ftp.close()
                except Exception:
                    pass


# ---------------------------------------------------------------------- S3
def build_s3_destination(device_code: str) -> tuple[str, str]:
    """Return (bucket, key_prefix) for today's upload.

    Matches v3: s3://{bucket}/ota/{device_code}/{YYYYMMDD}/
    """
    from datetime import datetime
    today = datetime.now().strftime("%Y%m%d")
    device_safe = (device_code or "unknown").strip()
    prefix = f"ota/{device_safe}/{today}/"
    return config.S3_UPLOAD_BUCKET, prefix


def upload_dir_to_s3(
    local_dir: str,
    device_code: str,
    *,
    log: LogFn = _noop_log,
    progress: ProgressFn = _noop_progress,
) -> TransferResult:
    """Upload every .zip in local_dir to S3 via boto3 (no aws CLI)."""
    try:
        import boto3  # imported lazily so the web app starts without AWS creds
    except ImportError:
        raise RuntimeError("未安装 boto3，无法上传 S3。请 pip install boto3。")

    result = TransferResult()
    bucket, prefix = build_s3_destination(device_code)
    result.s3_prefix = f"s3://{bucket}/{prefix}"

    if not os.path.isdir(local_dir):
        raise RuntimeError(f"本地目录不存在：{local_dir}")
    zip_files = sorted(
        os.path.join(local_dir, f) for f in os.listdir(local_dir) if f.lower().endswith(".zip")
    )
    if not zip_files:
        raise RuntimeError(f"目录中没有 .zip 文件：{local_dir}")

    s3 = boto3.client("s3")
    log(f"S3 目标：{result.s3_prefix}（共 {len(zip_files)} 个文件）")
    total = len(zip_files)
    for idx, local_path in enumerate(zip_files, start=1):
        filename = os.path.basename(local_path)
        key = prefix + filename
        progress((idx - 1) / total * 100, f"上传 {filename}（{idx}/{total}）")
        try:
            s3.upload_file(local_path, bucket, key)
            log(f"  ✓ 上传成功：{filename}")
            result.uploaded.append(filename)
        except Exception as e:
            log(f"  ✗ 上传失败：{filename}：{e}")
            result.failed.append(filename)

    progress(100, f"S3 上传完成：成功 {len(result.uploaded)} / 失败 {len(result.failed)}")
    return result


def cleanup_zip_dir(local_dir: str, log: LogFn = _noop_log) -> None:
    """Remove a 'temp'-style download dir after a successful upload (safe-guarded)."""
    local_dir = os.path.abspath(local_dir)
    if not os.path.isdir(local_dir):
        return
    try:
        shutil.rmtree(local_dir)
        log(f"已清理本地下载目录：{local_dir}")
    except Exception as e:
        log(f"清理本地目录失败：{local_dir}：{e}")


# ----------------------------------------------------------- orchestration
def download_and_upload(
    zip_urls: list[str],
    device_code: str,
    *,
    work_dir: str,
    auto_upload: bool = True,
    cleanup_after_upload: bool = True,
    log: LogFn = _noop_log,
    progress: ProgressFn = _noop_progress,
) -> TransferResult:
    """Full pipeline: download all FTP zips, then (optionally) upload to S3.

    Any download failure stops before upload (matches v3's fail-stop behaviour).
    """
    result = TransferResult()
    temp_dir = os.path.join(work_dir, "temp")
    os.makedirs(temp_dir, exist_ok=True)

    total = len(zip_urls)
    if total == 0:
        raise RuntimeError("未发现任何 FTP ZIP 地址。")

    log(f"开始下载 {total} 个 ZIP 文件...")
    for i, url in enumerate(zip_urls, start=1):
        try:
            path = download_ftp_zip(
                url, temp_dir, file_index=i, total_files=total, log=log, progress=progress
            )
            result.downloaded.append(path)
        except Exception as e:
            log(f"下载失败，终止后续动作：{url}：{e}")
            result.failed.append(url)
            return result

    if not auto_upload:
        return result

    up = upload_dir_to_s3(temp_dir, device_code, log=log, progress=progress)
    result.uploaded = up.uploaded
    result.failed.extend(up.failed)
    result.s3_prefix = up.s3_prefix

    if cleanup_after_upload and up.uploaded and not up.failed:
        cleanup_zip_dir(temp_dir, log=log)
    return result
