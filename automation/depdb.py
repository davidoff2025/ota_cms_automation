"""OTA 部署记录的永久存储（方案 D：JSONL 主存 + 可导出 CSV）。

- 主存储：<OTADATA>/deployments.jsonl，每次部署追加一行 JSON。
- 导出：<OTADATA>/deployments.csv，可随时从 JSONL 重新生成，供人用 Excel 查看。
- 幂等键：设备类型代码 + Rule ID 范围 + 目标版本——已存在同键记录则不重复写入。
- 写入时机：DASS 插入/更新数据完成后（见 dass_fill）。

存储目录默认 /Users/devinwang/Documents/海信运维交接/OTA升级部署/OTADATA，
可用环境变量 OTA_DEPLOY_DB_DIR 覆盖。
"""
from __future__ import annotations

import csv
import json
import os
from datetime import datetime

DEFAULT_DB_DIR = "/Users/devinwang/Documents/海信运维交接/OTA升级部署/OTADATA"
JSONL_NAME = "deployments.jsonl"
CSV_NAME = "deployments.csv"

# JSONL/CSV 的字段顺序（CSV 表头用它）。
FIELDS = [
    "批次号", "部署时间", "状态", "设备类型代码", "设备类型名称",
    "源版本", "目标版本", "升级类型",
    "Rule ID范围", "规则条数", "Rule ID列表", "设备扩展码集合",
    "作废时间", "Excel路径",
]

STATUS_ACTIVE = "有效"
STATUS_VOID = "作废"


def db_dir() -> str:
    return os.environ.get("OTA_DEPLOY_DB_DIR", DEFAULT_DB_DIR)


def _jsonl_path() -> str:
    return os.path.join(db_dir(), JSONL_NAME)


def _csv_path() -> str:
    return os.path.join(db_dir(), CSV_NAME)


def load_all() -> list[dict]:
    """读出所有部署记录（JSONL 每行一条）。文件不存在返回空列表。"""
    path = _jsonl_path()
    out: list[dict] = []
    if not os.path.isfile(path):
        return out
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                out.append(json.loads(line))
            except Exception:
                continue  # 跳过坏行，不让单行损坏影响整体
    return out


def _identity_key(rec: dict) -> tuple:
    """身份键（唯一组合标志，判断是否"同一批规则")：
    设备类型代码 + 源版本 + 目标版本 + 整批设备扩展码集合。【不含 Rule ID】。

    用途：同身份键视为同一批规则——即使重新生成导致 Rule ID 变了，也还是这一批。
    """
    return (
        str(rec.get("设备类型代码", "")).strip(),
        str(rec.get("源版本", "")).strip(),
        str(rec.get("目标版本", "")).strip(),
        str(rec.get("设备扩展码集合", "")).strip(),
    )


def _ruleids(rec: dict) -> str:
    return str(rec.get("Rule ID列表", "")).strip()


def _export_csv() -> None:
    """从 JSONL 全量重新生成 CSV（供人用 Excel 查看；utf-8-sig 让 Excel 正确识别中文）。"""
    rows = load_all()
    with open(_csv_path(), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow(r)


def _write_all(records: list[dict]) -> None:
    """整体重写 JSONL（用于作废标记后回写），并同步导出 CSV。"""
    os.makedirs(db_dir(), exist_ok=True)
    with open(_jsonl_path(), "w", encoding="utf-8") as f:
        for r in records:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
    _export_csv()


def append_record(rec: dict) -> tuple[bool, str]:
    """写入一条部署记录，按身份键做作废+替换。返回 (是否新写入, 说明)。

    三种情况（身份键 = 设备类型代码+源版本+目标版本+设备扩展码集合，不含 Rule ID）：
      1) 同身份键且【有效记录的 Rule ID 列表也相同】→ 完全重复，跳过不写。
      2) 同身份键但 Rule ID 列表不同（重新生成）→ 把该身份键下所有【有效】老记录标『作废』，
         写入一条新记录『有效』。
      3) 新身份键 → 直接写入『有效』。
    老记录不删除，仅标记作废（保留可追溯）。写入后同步导出 CSV。
    """
    os.makedirs(db_dir(), exist_ok=True)
    rec = dict(rec)
    rec.setdefault("部署时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    rec["状态"] = STATUS_ACTIVE

    all_recs = load_all()
    ident = _identity_key(rec)
    # 同身份键的"有效"记录。
    same_active = [r for r in all_recs
                   if _identity_key(r) == ident and str(r.get("状态", STATUS_ACTIVE)) == STATUS_ACTIVE]

    # 情况1：已有有效记录且 Rule ID 列表相同 → 完全重复，跳过。
    if any(_ruleids(r) == _ruleids(rec) for r in same_active):
        return False, f"已存在相同部署记录（同批且 Rule ID 相同），跳过写入。身份键={ident}"

    # 情况2：同身份键但 Rule ID 不同 → 作废旧的有效记录。
    voided = 0
    if same_active:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for r in all_recs:
            if (_identity_key(r) == ident
                    and str(r.get("状态", STATUS_ACTIVE)) == STATUS_ACTIVE):
                r["状态"] = STATUS_VOID
                r["作废时间"] = now
                voided += 1

    all_recs.append(rec)
    _write_all(all_recs)
    if voided:
        return True, (f"检测到同批规则重新生成：已将 {voided} 条旧记录标『作废』，"
                      f"并写入新 Rule ID 记录『有效』→ {_jsonl_path()}")
    return True, f"已写入部署记录『有效』→ {_jsonl_path()}（并同步导出 CSV）。"


def build_record_from_excel(xlsx_path: str) -> dict | None:
    """从更新后的 Excel（含 Rule ID）构造一条部署记录；无 Rule ID 列则返回 None。

    复用 summary 的读取逻辑，取设备码/版本/升级类型/Rule ID 列表与范围。
    """
    from .summary import _read_rows, _rule_id_range
    device_code, rows = _read_rows(xlsx_path)
    if not rows:
        return None
    rids = [r.get("Rule ID", "").strip() for r in rows if r.get("Rule ID", "").strip()]
    lo, hi = _rule_id_range(rows)
    rng = lo if lo == hi else f"{lo} ~ {hi}"
    first = rows[0]
    today = datetime.now().strftime("%Y%m%d")
    # 整批设备扩展码集合（去重排序后拼接），作为身份键的一部分。
    ext_set = sorted({r.get("机型扩展码", "").strip() for r in rows if r.get("机型扩展码", "").strip()})
    return {
        "批次号": f"{device_code}_{rng.replace(' ', '')}_{today}",
        "部署时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "状态": STATUS_ACTIVE,
        "设备类型代码": device_code,
        "设备类型名称": "",  # Excel 升级规则 sheet 的设备名称列可能为空；范围内不强求
        "源版本": first.get("源版本", ""),
        "目标版本": first.get("目标版本", ""),
        "升级类型": first.get("升级类型", ""),
        "Rule ID范围": rng,
        "规则条数": len(rids),
        "Rule ID列表": ",".join(rids),
        "设备扩展码集合": ",".join(ext_set),
        "作废时间": "",
        "Excel路径": os.path.abspath(xlsx_path),
    }


def record_deployment_from_excel(xlsx_path: str) -> tuple[bool, str]:
    """便捷入口：从 Excel 构造记录并（幂等）写入。返回 (是否新写入, 说明)。"""
    rec = build_record_from_excel(xlsx_path)
    if rec is None:
        return False, "该 Excel 无 Rule ID 数据，未记录部署。"
    return append_record(rec)


# —— 查看（deployota history）——
def _print_history(limit: int | None, device: str | None) -> int:
    from ._log import ok, warn, color
    rows = load_all()
    if device:
        rows = [r for r in rows if device.lower() in str(r.get("设备类型代码", "")).lower()]
    if not rows:
        print(warn("暂无部署记录。"))
        print(color(f"（存储位置：{_jsonl_path()}）", "dim"))
        return 0
    rows = rows[-limit:] if limit else rows
    active_n = sum(1 for r in rows if str(r.get("状态", STATUS_ACTIVE)) == STATUS_ACTIVE)
    void_n = len(rows) - active_n
    print(color(f"========== OTA 部署历史（共 {len(rows)} 条：有效 {active_n} / 作废 {void_n}，"
                f"存储：{db_dir()}）==========", "bold"))
    for r in rows:
        status = str(r.get("状态", STATUS_ACTIVE))
        is_void = status == STATUS_VOID
        tag = (color("[作废]", "red") if is_void else color("[有效]", "green"))
        extra = f"  作废于 {r.get('作废时间','')}" if is_void else ""
        line = (f"· {r.get('部署时间','')}  {tag}  {r.get('设备类型代码','')}  "
                f"{r.get('源版本','')} → {r.get('目标版本','')}  "
                f"[{r.get('升级类型','')}]  Rule ID {r.get('Rule ID范围','')}"
                f"（{r.get('规则条数','')} 条）{extra}")
        print(color(line, "dim") if is_void else line)
    print(ok(f"CSV 导出：{_csv_path()}"))
    return 0


def main() -> int:
    """history 查看入口：python -m automation.depdb [--device CODE] [--limit N] [--export]"""
    import argparse
    import sys
    p = argparse.ArgumentParser(prog="deployota history",
                                description="查看 OTA 部署历史（JSONL 主存，CSV 可导出）。")
    p.add_argument("--device", default=None, help="按设备类型代码过滤（子串匹配）。")
    p.add_argument("--limit", type=int, default=20, help="只显示最近 N 条（默认 20；0 表示全部）。")
    p.add_argument("--export", action="store_true", help="重新导出一份 CSV 后退出。")
    args = p.parse_args()
    if args.export:
        _export_csv()
        print(f"已导出 CSV → {_csv_path()}")
        return 0
    return _print_history(args.limit or None, args.device)


if __name__ == "__main__":
    import sys
    sys.exit(main())
