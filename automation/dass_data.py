"""Read the "updated Rule ID" Excel and map each row to DASS fill-in fields.

The input is the workbook ota_v4 produces AFTER the Rule ID update (it has the
three inserted columns 升级方式 / 升级日期 / Rule ID). Each data row becomes one
DASS record.

Confirmed Excel -> DASS mapping below. Columns not present in the Excel
(大区域/国家/所有人/是否展示/当前阶段/状态/各时间/灰度档位) are left blank for now
and will be filled per rules you confirm later.
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

# DASS fill-in table columns (left to right), excluding the auto 记录id column.
DASS_COLUMNS = [
    "大区域", "国家", "方案号", "升级类型", "OTA类型", "featurecode", "机型",
    "设备扩展信息", "品牌", "所有人", "是否展示", "ruleid", "源版本", "目标版本",
    "当前阶段", "当前阶段时间", "状态", "停止时间", "场测时间", "场测mac", "备注",
    "50", "200", "1,000", "5,000", "10,000", "完全开放",
]

# Excel header -> DASS column (direct, confirmed mappings).
# NOTE: the human-readable 升级方式 (正常升级/强制升级) goes into DASS 升级类型.
# Some workbooks name columns differently (e.g. 升级模式/时间 instead of
# 升级方式/升级日期). Multiple Excel headers can map to the same DASS column; the
# first one found in the sheet wins.
EXCEL_TO_DASS = {
    "设备类型代码": "方案号",
    "升级方式": "升级类型",
    "升级模式": "升级类型",       # 别名
    "特征码": "featurecode",
    "内部机型信息": "机型",
    "设备扩展信息": "设备扩展信息",
    "品牌组": "品牌",
    "Rule ID": "ruleid",
    "源版本": "源版本",
    "目标完整版本": "目标版本",   # DASS『目标版本』取 Excel 的『目标完整版本』列
    "升级日期": "当前阶段时间",
    "时间": "当前阶段时间",       # 别名
}

# How a Excel "区域组" value maps to DASS 大区域 / 国家. Extend as needed.
# Confirmed rules:
#   区域组=全球   -> 大区域=全球, 国家=全球
#   区域组=不区分 -> 大区域 与 国家 都留空（不需要填写）
REGION_TO_DASS = {
    "全球": {"大区域": "全球", "国家": "全球"},
    "不区分": {"大区域": "", "国家": ""},
}


def map_region(region_group: str) -> dict:
    """Split a 区域组 value into DASS 大区域 / 国家.

    Known values use REGION_TO_DASS; unknown non-empty values default to filling
    国家 with the raw value and leaving 大区域 blank (so nothing is silently lost).
    """
    rg = str(region_group or "").strip()
    if not rg:
        return {"大区域": "", "国家": ""}
    if rg in REGION_TO_DASS:
        return dict(REGION_TO_DASS[rg])
    return {"大区域": "", "国家": rg}


def read_updated_excel(path: str, sheet_name: str = "升级规则") -> list[dict]:
    """Return a list of DASS records (one per data row).

    Each record is a dict keyed by DASS_COLUMNS; unmapped columns are "".
    """
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Excel 文件不存在：{path}")

    wb = load_workbook(str(p), data_only=True)
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.worksheets[0]

    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    col_index = {h: i for i, h in enumerate(headers)}

    records: list[dict] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(v not in (None, "") for v in row):
            continue
        rec = {c: "" for c in DASS_COLUMNS}
        for excel_col, dass_col in EXCEL_TO_DASS.items():
            idx = col_index.get(excel_col)
            if idx is None or idx >= len(row):
                continue
            val = row[idx]
            val = "" if val is None else str(val).strip()
            # 别名列：若该 DASS 列已被先前的别名填上非空值，则不要用空值覆盖它。
            if val == "" and rec[dass_col] != "":
                continue
            rec[dass_col] = val

        # 区域组 -> 大区域 / 国家 (confirmed rule).
        rg_idx = col_index.get("区域组")
        region_group = "" if rg_idx is None or rg_idx >= len(row) else (row[rg_idx] or "")
        region = map_region(region_group)
        rec["大区域"] = region["大区域"]
        rec["国家"] = region["国家"]

        records.append(rec)
    wb.close()
    return records


def latest_updated_excel(directory: str) -> str:
    """Return the most recently modified .xlsx in a directory (helper)."""
    d = Path(directory)
    xlsx = sorted(d.glob("*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True)
    if not xlsx:
        raise FileNotFoundError(f"目录中没有 .xlsx 文件：{directory}")
    return str(xlsx[0])


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("用法：python -m automation.dass_data <更新后的Excel路径>")
        raise SystemExit(1)
    rows = read_updated_excel(sys.argv[1])
    print(f"读取到 {len(rows)} 条记录：")
    for r in rows:
        print({k: v for k, v in r.items() if v})
