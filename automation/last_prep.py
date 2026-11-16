"""记录 / 读取「prep 最近一次生成（或复用）的 Excel」。

prep 通过严格校验后调用 save()；deployota 在分发 upload / online / dass 时，
若命令行没有给 Excel，就用 load() 取这个路径补上，让整条流水线接着同一份 Excel 往下走。
状态文件放在项目根目录 .last_prep.json（已 gitignore，仅本机有效）。
"""
from __future__ import annotations

import json
import os
import time
from pathlib import Path

STATE_FILE = Path(__file__).resolve().parent.parent / ".last_prep.json"


def save(excel_path: str) -> None:
    """记录最近一次 prep 的 Excel（绝对路径）。写失败不影响 prep 主流程。"""
    try:
        data = {
            "excel": os.path.abspath(excel_path),
            "time": time.strftime("%Y-%m-%d %H:%M:%S"),
        }
        STATE_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    except OSError:
        pass


def load() -> dict | None:
    """返回 {"excel": 路径, "time": 时间}；没有记录或记录损坏返回 None。"""
    try:
        data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        if isinstance(data, dict) and data.get("excel"):
            return data
    except (OSError, ValueError):
        pass
    return None


def forget(excel_path: str) -> bool:
    """若『最近 prep』记的正是 excel_path，就删掉这条记录（prep 的 MAC 组失败时用）。返回是否删了。"""
    rec = load()
    if not rec or os.path.abspath(rec["excel"]) != os.path.abspath(excel_path):
        return False
    try:
        STATE_FILE.unlink()
        return True
    except OSError:
        return False


# ----------------------------------------------------------------------
# Excel 附带信息：<Excel 名>.meta.json（目前只有『发起人 / Originator』）。
# 不写进 Excel 本身（标准 18 列 Excel 要导入 CMS，多一列可能导入失败）。
# ----------------------------------------------------------------------

def meta_path(excel_path: str) -> str:
    base, _ = os.path.splitext(os.path.abspath(excel_path))
    return base + ".meta.json"


def save_meta(excel_path: str, **fields) -> None:
    """合并写入附带信息（空值不覆盖已有值）。写失败不影响主流程。"""
    try:
        data = load_meta(excel_path)
        data.update({k: str(v).strip() for k, v in fields.items() if str(v or "").strip()})
        data["time"] = time.strftime("%Y-%m-%d %H:%M:%S")
        Path(meta_path(excel_path)).write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n",
                                               encoding="utf-8")
    except OSError:
        pass


def load_meta(excel_path: str) -> dict:
    try:
        data = json.loads(Path(meta_path(excel_path)).read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def originator_for(excel_path: str | None) -> str:
    """发起人：环境变量 OTA_ORIGINATOR（--originator 传入）优先，其次 Excel 附带信息。"""
    env = os.environ.get("OTA_ORIGINATOR", "").strip()
    if env:
        return env
    return str(load_meta(excel_path).get("originator", "")).strip() if excel_path else ""
