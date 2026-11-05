"""OTA 模板『prep CLI』：命令行粘贴/传入升级信息文本，三步串成一键流程——

  步骤1 生成：用自带的 OTAEngine 解析文本、生成 OTA 导入 Excel（升级规则 + 升级提示语两表）。
  步骤2 校验：对生成的工作簿做严格校验（表名/表头/每行升级文件地址与 OTA 类型等规则）。
  步骤3 上传：从源数据抽取 FTP .zip 地址，下载后上传到 S3（download_and_upload）。

用法：
    cd ota_cms_automation
    # 从 stdin 粘贴文本（推荐）：
    cat 文本.txt | ./.venv/bin/python -m automation.prep --device-code MT9655AMAndrGTV
    # 或从文件读：
    ./.venv/bin/python -m automation.prep --file 文本.txt --device-code MT9655AMAndrGTV
    # 只生成+校验，不碰 FTP/S3（打印 S3 计划与 ftp zip 清单）：
    ... | ./.venv/bin/python -m automation.prep --no-upload --device-code MT9655AMAndrGTV

文本来源：给了 --file 读该文件，否则读 stdin。CLI 的 --device-code/--device-name/
--feature-code/--mac-group 非空时覆盖从文本解析出的设备信息。

真实上传前置条件（仅在不带 --no-upload 时需要）：
  - FTP：config/.env 里的 OTA_FTP_PASS（及可选 OTA_FTP_USER）非占位值。
  - AWS：区域/密钥不在 .env 里配置，走 boto3 默认链（~/.aws），region=us-west。
『真实 S3 上传 / FTP 下载』是不可逆的外部 I/O；要安全演练请始终带 --no-upload。

返回码：
    0  成功（生成+校验通过；上传模式下全部上传成功，或 --no-upload 正常结束）
    1  业务失败（生成报错 / 校验不通过 / 无可上传地址 / 上传有失败项）
    2  参数或安全失败（未读到文本 / --file 不存在）
"""
from __future__ import annotations

import argparse
import os
import re
import sys
import tempfile

from openpyxl import load_workbook

from . import config as cms_config  # noqa: F401  触发 ota_cms 的 config/.env 加载
from .otagen.engine import OTAEngine
from .otagen import transfer
from .otagen.legacy_engine import OUTPUT_HEADERS


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="python -m automation.prep",
        description="OTA 模板 prep CLI：生成 Excel → 严格校验 → 上传 S3（三步一键）。",
    )
    p.add_argument("--file", dest="file", default=None,
                   help="从该路径读取升级信息文本；省略则从 stdin 读。")
    p.add_argument("--device-code", dest="device_code", default="",
                   help="设备类型代码（方案号）；非空时覆盖文本解析值。")
    p.add_argument("--device-name", dest="device_name", default="",
                   help="设备类型名称；非空时覆盖文本解析值。")
    p.add_argument("--feature-code", dest="feature_code", default="",
                   help="特征码（24 位数字，可留空）；非空时覆盖文本解析值。")
    p.add_argument("--mac-group", dest="mac_group", default="",
                   help="MAC 组；非空时覆盖文本解析值。")
    p.add_argument("--out-dir", dest="out_dir", default=None,
                   help="生成 Excel 的输出目录；省略则用 automation/otagen/data/outputs。")
    p.add_argument("--no-upload", dest="no_upload", action="store_true", default=False,
                   help="只生成+校验并打印 S3 计划与 ftp zip 清单，绝不触碰 FTP/S3。")
    return p


def _validate_output_row_like(row: dict) -> list[str]:
    """单行规则校验（运行时与 _selftest 共用的唯一行规则）。返回问题串列表（空=通过）。

    规则：
      - 升级文件地址：非空；re.match(r'^https?://', v, re.I)；不含 3+ 连续斜杠
        （re.search(r'/{3,}', v)，用正则而非字面 '///'）；以 .zip 结尾（忽略大小写）。
      - OTA类型：str(v).strip() 属于 ('1','2','11')。
    """
    problems: list[str] = []

    url = str(row.get("升级文件地址", "") or "").strip()
    if not url:
        problems.append("升级文件地址为空")
    else:
        if not re.match(r"^https?://", url, re.I):
            problems.append(f"升级文件地址非 http(s)://：{url!r}")
        if re.search(r"/{3,}", url):
            problems.append(f"升级文件地址含 3+ 连续斜杠：{url!r}")
        if not url.lower().endswith(".zip"):
            problems.append(f"升级文件地址未以 .zip 结尾：{url!r}")

    ota_type = str(row.get("OTA类型", "") or "").strip()
    if ota_type not in ("1", "2", "11"):
        problems.append(f"OTA类型非法（应为 1/2/11）：{ota_type!r}")

    return problems


def _validate_excel(path: str) -> list[str]:
    """严格校验生成的工作簿，收集所有问题并返回（空列表=通过）。"""
    problems: list[str] = []
    wb = load_workbook(path, read_only=True)
    try:
        if wb.sheetnames != ["升级规则", "升级提示语"]:
            problems.append(f"表名应为 ['升级规则','升级提示语']，实际：{wb.sheetnames}")
            return problems

        ws = wb["升级规则"]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            problems.append("升级规则表为空（无任何行）。")
            return problems

        header = [("" if c is None else str(c)) for c in rows[0]]
        if header != list(OUTPUT_HEADERS):
            problems.append(f"升级规则表头与 OUTPUT_HEADERS 不一致，实际：{header}")
            return problems

        data_rows = rows[1:]
        if len(data_rows) < 1:
            problems.append("升级规则表无数据行（至少需要 1 行）。")
            return problems

        for idx, raw in enumerate(data_rows, start=2):
            rowdict = {h: ("" if v is None else str(v)) for h, v in zip(OUTPUT_HEADERS, raw)}

            for prob in _validate_output_row_like(rowdict):
                problems.append(f"第{idx}行：{prob}")

            device_code = str(rowdict.get("设备类型代码", "") or "").strip()
            if not device_code:
                problems.append(f"第{idx}行：设备类型代码为空")
            elif re.search(r"\s", device_code):
                problems.append(f"第{idx}行：设备类型代码含空白：{device_code!r}")
            elif not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_\-().]*", device_code):
                problems.append(f"第{idx}行：设备类型代码含非法字符：{device_code!r}")

            feature_code = str(rowdict.get("特征码", "") or "").strip()
            if feature_code and not re.fullmatch(r"\d{24}", feature_code):
                problems.append(f"第{idx}行：特征码非 24 位数字：{feature_code!r}")

            if not str(rowdict.get("MAC组", "") or "").strip():
                problems.append(f"第{idx}行：MAC组为空")

            internal = str(rowdict.get("内部机型信息", "") or "").strip()
            if not internal:
                problems.append(f"第{idx}行：内部机型信息为空")

            ext = str(rowdict.get("设备扩展信息", "") or "").strip()
            if not ext:
                problems.append(f"第{idx}行：设备扩展信息为空")
            else:
                if ext == internal:
                    problems.append(f"第{idx}行：设备扩展信息与内部机型信息相同：{ext!r}")
                if ext.isdigit():
                    problems.append(f"第{idx}行：设备扩展信息不应全为数字：{ext!r}")

        return problems
    finally:
        wb.close()


def main() -> int:
    args = _build_parser().parse_args()

    # (1) 读取文本来源。
    if args.file:
        if not os.path.isfile(args.file):
            print(f"✗ 找不到 --file 指定的文件：{args.file}", file=sys.stderr)
            return 2
        with open(args.file, "r", encoding="utf-8") as f:
            text = f.read()
    else:
        text = sys.stdin.read()

    if text.strip() == "":
        print("✗ 未从 stdin/--file 读到任何文本", file=sys.stderr)
        return 2

    eng = OTAEngine()

    # 设备信息：先从文本解析，再用 CLI 非空值覆盖。
    info = eng.resolve_device_info_from_text(text)
    device_code = args.device_code.strip() or info.get("device_code", "")
    device_name = args.device_name.strip() or info.get("device_name", "")
    feature_code = args.feature_code.strip() or info.get("feature_code", "")
    mac_group = args.mac_group.strip() or info.get("mac_group", "")

    # 若设备码经 CLI 覆盖后才确定、而 MAC 组仍为空，则按设备预设补齐
    # （预设默认 '测试'+设备类型代码，与引擎 resolve 的建议值一致）。
    if not mac_group and device_code:
        mac_group = eng.device_preset(device_code).get("mac_group", "")

    sj = eng.parse_pasted_text(text)

    if args.out_dir:
        out_dir = args.out_dir
    else:
        from .otagen import config as otagen_config
        out_dir = str(otagen_config.DATA_DIR / "outputs")

    # (2) 步骤1 生成。
    try:
        res = eng.generate_excel(
            sj,
            device_code=device_code,
            device_name=device_name,
            feature_code=feature_code,
            mac_group=mac_group,
            output_dir=out_dir,
        )
    except ValueError as e:
        print("步骤1 生成失败：", file=sys.stderr)
        print(str(e), file=sys.stderr)
        return 1

    output_path = os.path.abspath(res["output_path"])
    final_device_code = device_code or ""

    # (3) 步骤2 严格校验。
    problems = _validate_excel(output_path)
    if problems:
        print("步骤2 严格校验未通过：", file=sys.stderr)
        for prob in problems:
            print("- " + prob, file=sys.stderr)
        return 1

    # (4) 步骤3 上传（或 --no-upload 计划）。
    urls = eng.extract_ftp_zip_urls(sj)
    bucket, prefix = transfer.build_s3_destination(final_device_code)

    if args.no_upload:
        print(output_path)
        print(f"行数：{res.get('row_count')}  提示语数：{res.get('message_count')}")
        print(f"[--no-upload] S3 计划：s3://{bucket}/{prefix}")
        print(f"[--no-upload] 待上传 FTP zip（{len(urls)} 个）：")
        for u in urls:
            print(f"  - {u}")
        return 0

    if not urls:
        print("无可上传的 FTP zip（源数据未含 ftp:// 地址）", file=sys.stderr)
        return 1

    work_dir = tempfile.mkdtemp(prefix="otagen_")
    result = transfer.download_and_upload(
        urls,
        final_device_code,
        work_dir=work_dir,
        auto_upload=True,
        cleanup_after_upload=True,
        log=print,
    )
    print(output_path)
    print(f"S3 前缀：{result.s3_prefix}")
    print(f"已下载：{len(result.downloaded)}  已上传：{len(result.uploaded)}  失败：{len(result.failed)}")
    if result.failed:
        print(f"失败项：{result.failed}", file=sys.stderr)
    return 1 if result.failed else 0


# —— 不触碰 FTP/S3/~/.aws 的纯离线自测（python -m automation.prep --selftest）——
def _selftest() -> int:
    from datetime import datetime

    ok = True
    temp_dirs: list[str] = []

    def check(name, cond):
        nonlocal ok
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        if not cond:
            ok = False

    try:
        fixture = (
            "机型\t扩展信息\t源版本\t目标版本\tOTA类型\t升级包地址\n"
            "MT9655AMAndrGTV\tABCD1234\t0100AP0920\t0100AP1030\t1\t"
            "ftp://ftpuser@host/a/USBOTA_x.zip\n"
            "语种\t描述\n"
            "简体中文\t测试升级\n"
        )
        eng = OTAEngine()
        sj = eng.parse_pasted_text(fixture)

        out_dir = tempfile.mkdtemp(prefix="otagen_selftest_")
        temp_dirs.append(out_dir)
        res = eng.generate_excel(
            sj,
            device_code="MT9655AMAndrGTV",
            device_name="测试机",
            feature_code="",
            mac_group="测试MT9655AMAndrGTV",
            output_dir=out_dir,
        )
        gen = res["output_path"]

        wb = load_workbook(gen, read_only=True)
        try:
            check("表名==['升级规则','升级提示语']", wb.sheetnames == ["升级规则", "升级提示语"])
            ws = wb["升级规则"]
            rows = list(ws.iter_rows(values_only=True))
            header = [("" if c is None else str(c)) for c in rows[0]]
            check("表头==OUTPUT_HEADERS", header == list(OUTPUT_HEADERS))
            first = {h: ("" if v is None else str(v)) for h, v in zip(OUTPUT_HEADERS, rows[1])}
            check("OTA类型 ∈ {1,2,11}", str(first.get("OTA类型", "")).strip() in ("1", "2", "11"))
            addr = str(first.get("升级文件地址", "") or "")
            check("升级文件地址 https:// 前缀", addr.startswith("https://"))
            check("升级文件地址 .zip 结尾", addr.lower().endswith(".zip"))
            check("升级文件地址 无 3+ 连续斜杠", not re.search(r"/{3,}", addr))
        finally:
            wb.close()

        check("_validate_excel 生成件无问题", _validate_excel(gen) == [])

        # 坏行：_validate_output_row_like 必须抓出各类问题（用正则，不用字面 '///'）。
        check("坏行 OTA类型=9 被抓",
              any("OTA类型" in p for p in _validate_output_row_like(
                  {"升级文件地址": "https://a/x.zip", "OTA类型": "9"})))
        check("坏行 地址非 http 被抓",
              any("升级文件地址" in p for p in _validate_output_row_like(
                  {"升级文件地址": "ftp://a/x.zip", "OTA类型": "1"})))
        check("坏行 地址含 3+ 斜杠被抓",
              any("连续斜杠" in p for p in _validate_output_row_like(
                  {"升级文件地址": "https://a///x.zip", "OTA类型": "1"})))
        # 特征码非 24 位：由 _validate_excel 层校验，这里单测 fullmatch 语义。
        check("特征码非 24 位判定", re.fullmatch(r"\d{24}", "12345") is None)

        # extract_ftp_zip_urls 去重。
        urls = eng.extract_ftp_zip_urls(sj)
        check("extract_ftp_zip_urls 去重为单条",
              urls == ["ftp://ftpuser@host/a/USBOTA_x.zip"])

        # build_s3_destination 目标。
        today = datetime.now().strftime("%Y%m%d")
        dest = transfer.build_s3_destination("CODE")
        check("build_s3_destination==('fam-media-andr','ota/CODE/<today>/')",
              dest == ("fam-media-andr", f"ota/CODE/{today}/"))

        # _build_parser 解析。
        ns = _build_parser().parse_args(
            ["--no-upload", "--file", "x.txt", "--device-code", "C"])
        check("解析 --no-upload/--file/--device-code", ns.no_upload and ns.file == "x.txt" and ns.device_code == "C")
        ns2 = _build_parser().parse_args([])
        check("--no-upload 默认 False", ns2.no_upload is False)

        # 自带子包可导入。
        import importlib
        for mod in ("automation.otagen.engine", "automation.otagen.transfer", "automation.otagen.config"):
            try:
                importlib.import_module(mod)
                check(f"import {mod} 成功", True)
            except Exception as e:
                check(f"import {mod} 成功（{e}）", False)

    finally:
        import shutil
        for d in temp_dirs:
            shutil.rmtree(d, ignore_errors=True)

    print("\n自测结果：", "全部通过 ✅" if ok else "存在失败 ❌")
    return 0 if ok else 1


if __name__ == "__main__":
    if "--selftest" in sys.argv[1:]:
        sys.exit(_selftest())
    sys.exit(main())
