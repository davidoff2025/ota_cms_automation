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

JSONL_NAME = "deployments.jsonl"
CSV_NAME = "deployments.csv"

# JSONL/CSV 的字段顺序（CSV 表头用它）。
FIELDS = [
    "Rule ID范围", "Rule ID列表", "规则条数",
    "设备类型代码", "设备类型名称", "源版本", "目标版本", "目标完整版本", "升级类型",
    "Owner", "状态", "部署时间", "入库时间", "设备扩展码集合",
    "作废时间", "批次号", "Excel路径",
]

STATUS_ACTIVE = "有效"
STATUS_VOID = "作废"


def db_dir() -> str:
    """部署数据库目录：优先 OTA_DEPLOY_DB_DIR；否则取 OTA 工作目录下的 OTADATA 子目录。"""
    explicit = os.environ.get("OTA_DEPLOY_DB_DIR")
    if explicit:
        return explicit
    from .otagen import config as otagen_config
    return os.path.join(otagen_config.WORKSPACE, "OTADATA")


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


def full_target_versions(rec: dict) -> list[str]:
    """批次的『目标完整版本』列表（如 ['0100UQ1104']；一个批次可能有多个）。

    来源依次：记录里的『目标完整版本』字段 → 『规则明细』里各规则的目标版本。都没有返回 []
    （例如 CMS 回补的老批次：CMS 页面只给『版本变化』Q0602_to_Q0724，没有完整版本）。
    """
    v = str(rec.get("目标完整版本", "") or "").strip()
    if v:
        return [x for x in v.split("/") if x]
    seen: list[str] = []
    for d in rec.get("规则明细") or []:
        t = str(d.get("目标版本", "") or "").strip()
        if t and t not in seen:
            seen.append(t)
    return sorted(seen)


def display_target(rec: dict) -> tuple[str, bool]:
    """汇总里显示的目标版本：(文本, 是否为完整版本)。有完整版本显示完整版本（多个用 / 连接），
    否则退回『版本变化』。"""
    full = full_target_versions(rec)
    if full:
        return "/".join(full), True
    return str(rec.get("目标版本", "") or ""), False


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


def append_record(rec: dict, force_status: str | None = None) -> tuple[bool, str]:
    """写入一条部署记录，按身份键做作废+替换。返回 (是否新写入, 说明)。

    三种情况（身份键 = 设备类型代码+源版本+目标版本+设备扩展码集合，不含 Rule ID）：
      1) 同身份键且【有效记录的 Rule ID 列表也相同】→ 完全重复，跳过不写。
      2) 同身份键但 Rule ID 列表不同（重新生成）→ 把该身份键下所有【有效】老记录标『作废』，
         写入一条新记录『有效』。
      3) 新身份键 → 直接写入『有效』。
    老记录不删除，仅标记作废（保留可追溯）。写入后同步导出 CSV。

    force_status：可选，强制写入记录的状态。
      - None（默认）→ 写入『有效』(STATUS_ACTIVE)，与历史行为完全一致（Excel 入口不传该参数，零影响）。
      - STATUS_VOID → 新记录以『作废』状态落库（用于回填 CMS 中备注含 VOID 的历史记录）；
        若 rec 未带『作废时间』则补当前时刻。
    身份键的三分支逻辑（情况1/2/3）不受 force_status 影响。
    """
    os.makedirs(db_dir(), exist_ok=True)
    rec = dict(rec)
    # 入库时间=写库时刻（兜底）；部署时间来自 Excel 升级日期，不在此用 now 冒充。
    rec.setdefault("入库时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    if force_status is None:
        rec["状态"] = STATUS_ACTIVE
    else:
        rec["状态"] = force_status
        if force_status == STATUS_VOID:
            rec.setdefault("作废时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

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


def build_record_from_excel(xlsx_path: str, owner: str = "") -> dict | None:
    """从更新后的 Excel（含 Rule ID）构造一条部署记录；无 Rule ID 列则返回 None。

    复用 summary 的读取逻辑，取设备码/版本/升级类型/Rule ID 列表与范围。
    owner 为本次 OTA 请求发起人（可空）。
    """
    from .summary import _read_rows, _rule_id_range
    device_code, rows = _read_rows(xlsx_path)
    if not rows:
        return None
    rids = [r.get("Rule ID", "").strip() for r in rows if r.get("Rule ID", "").strip()]
    lo, hi = _rule_id_range(rows)
    rng = lo if lo == hi else f"{lo} ~ {hi}"
    first = rows[0]
    # 部署时间（真实部署日期，非入库时刻）优先级：
    #   1) Excel『升级日期』列（新文件有，collect 回写）；
    #   2) 文件名里的 8 位日期 YYYYMMDD（老文件没有升级日期列，但文件名带日期，最可靠）；
    #   3) 文件修改时间（兜底）。
    deploy_time = str(first.get("部署时间", "") or "").strip()
    if not deploy_time:
        import re as _re
        m = _re.search(r"(20\d{6})", os.path.basename(xlsx_path))
        if m:
            deploy_time = m.group(1)
    if not deploy_time:
        try:
            deploy_time = datetime.fromtimestamp(os.path.getmtime(xlsx_path)).strftime("%Y%m%d")
        except Exception:
            deploy_time = datetime.now().strftime("%Y%m%d")
    # 整批设备扩展码集合（去重排序后拼接），作为身份键的一部分。
    ext_set = sorted({r.get("机型扩展码", "").strip() for r in rows if r.get("机型扩展码", "").strip()})
    # 每条规则自己的明细（供 `deployota update <RuleID>` 直接用，不依赖 Excel 还在不在）。
    # 失败不影响写库：拿不到就留空，查询时会回退去读来源 Excel。
    details: list[dict] = []
    try:
        from .dass_data import read_updated_excel
        for d in read_updated_excel(xlsx_path):
            rid = str(d.get("ruleid", "")).strip()
            if rid:
                details.append({k: str(d.get(k, "") or "") for k in
                                ("ruleid", "方案号", "目标版本", "设备扩展信息", "机型", "源版本", "升级类型", "featurecode")})
    except Exception:
        details = []
    return {
        "批次号": f"{device_code}_{rng.replace(' ', '')}_{deploy_time}",
        "部署时间": deploy_time,
        "入库时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "状态": STATUS_ACTIVE,
        "Owner": str(owner or "").strip(),
        "设备类型代码": device_code,
        "设备类型名称": "",  # Excel 升级规则 sheet 的设备名称列可能为空；范围内不强求
        "源版本": first.get("源版本", ""),
        # 去重身份键一直用 Excel『目标版本』列（版本变化，如 Q0305_to_Q1104）；汇总显示改成完整版本后，
        # _read_rows 把它单独放在『版本变化』里。老记录与新记录因此仍能正确去重。
        "目标版本": first.get("版本变化", "") or first.get("目标版本", ""),
        # 汇总显示用：整批的目标完整版本（_read_rows 的『目标版本』已是完整版本；多个用 / 连接）。
        "目标完整版本": "/".join(sorted({r.get("目标版本", "").strip() for r in rows if r.get("目标版本", "").strip()})),
        "升级类型": first.get("升级类型", ""),
        "Rule ID范围": rng,
        "规则条数": len(rids),
        "Rule ID列表": ",".join(rids),
        "设备扩展码集合": ",".join(ext_set),
        "作废时间": "",
        "Excel路径": os.path.abspath(xlsx_path),
        "规则明细": details,
    }


def record_deployment_from_excel(xlsx_path: str, owner: str = "") -> tuple[bool, str]:
    """便捷入口：从 Excel 构造记录并（幂等/作废替换）写入。返回 (是否新写入, 说明)。"""
    rec = build_record_from_excel(xlsx_path, owner=owner)
    if rec is None:
        return False, "该 Excel 无 Rule ID 数据，未记录部署。"
    return append_record(rec)


# —— 查看（deployota history）——
def _rid_in_record(rec: dict, needle: str) -> bool:
    """判断某 Rule ID 或范围文本 needle 是否匹配该记录（范围子串 或 落在数值区间内）。"""
    needle = str(needle).strip()
    if not needle:
        return True
    rng = str(rec.get("Rule ID范围", ""))
    lst = str(rec.get("Rule ID列表", ""))
    # 直接子串命中（范围文本或列表里出现）。
    if needle in rng or needle in lst:
        return True
    # 若 needle 是纯数字且记录范围是 "a ~ b"，判断是否落在区间。
    if needle.isdigit():
        ids = [x.strip() for x in lst.split(",") if x.strip()]
        if needle in ids:
            return True
        m = [int(x) for x in ids if x.isdigit()]
        if m and min(m) <= int(needle) <= max(m):
            return True
    return False


def _print_history(limit: int | None, device: str | None,
                   owner: str | None = None, rule: str | None = None) -> int:
    from ._log import ok, warn, color
    rows = load_all()
    if device:
        rows = [r for r in rows if device.lower() in str(r.get("设备类型代码", "")).lower()]
    if owner:
        rows = [r for r in rows if owner.lower() in str(r.get("Owner", "")).lower()]
    if rule:
        rows = [r for r in rows if _rid_in_record(r, rule)]
    if not rows:
        print(warn("没有匹配的部署记录。"))
        print(color(f"（存储位置：{_jsonl_path()}）", "dim"))
        return 0
    # 按『部署时间』升序排序（稳定排序；同日保持写入顺序），这样末尾 = 真正最新的部署。
    # 注意：不能依赖文件写入顺序——回填是从新往旧抓的，文件末尾反而是最老的记录。
    rows = sorted(rows, key=lambda r: str(r.get("部署时间", "")))
    all_rows = rows                      # 过滤后的全部记录：总体统计基于它，不受 --limit 影响。
    shown = rows[-limit:] if limit else rows   # 仅展示列表截取最近 N 条。

    active_n = sum(1 for r in all_rows if str(r.get("状态", STATUS_ACTIVE)) == STATUS_ACTIVE)
    void_n = len(all_rows) - active_n
    print(color(f"========== OTA 部署历史（显示最近 {len(shown)} 条 / 共 {len(all_rows)} 条："
                f"有效 {active_n} / 作废 {void_n}，存储：{db_dir()}）==========", "bold"))
    for r in shown:
        status = str(r.get("状态", STATUS_ACTIVE))
        is_void = status == STATUS_VOID
        tag = (color("[作废]", "red") if is_void else color("[有效]", "green"))
        extra = f"  作废于 {r.get('作废时间','')}" if is_void else ""
        owner = r.get("Owner", "") or "-"
        tgt, is_full = display_target(r)
        tgt = tgt if is_full else f"{tgt}（版本变化，无完整版本）"
        line = (f"· Rule ID {r.get('Rule ID范围',''):15}（{r.get('规则条数','')} 条）  {tag}  "
                f"{r.get('设备类型代码',''):18} {r.get('源版本','')} → {tgt}  "
                f"[{r.get('升级类型','')}]  owner={owner}  {r.get('部署时间','')}{extra}")
        print(color(line, "dim") if is_void else line)

    # —— 总体统计 summary（基于 all_rows：过滤后的全部记录，不受 --limit 截断影响）——
    platforms = sorted({str(r.get("设备类型代码", "")).strip() for r in all_rows if str(r.get("设备类型代码", "")).strip()})
    owners = sorted({str(r.get("Owner", "")).strip() for r in all_rows if str(r.get("Owner", "")).strip()})
    # 统计去重后的 Rule ID 总数（按『有效』记录的 Rule ID 列表汇总，避免作废/重复重复计数）。
    active_rids = set()
    for r in all_rows:
        if str(r.get("状态", STATUS_ACTIVE)) == STATUS_ACTIVE:
            active_rids.update(x.strip() for x in str(r.get("Rule ID列表", "")).split(",") if x.strip())
    rid_nums = sorted(int(x) for x in active_rids if x.isdigit())
    rid_span = f"{rid_nums[0]} ~ {rid_nums[-1]}" if rid_nums else "-"
    # 部署时间跨度（all_rows 已按部署时间升序）。
    dates = [str(r.get("部署时间", "")) for r in all_rows if str(r.get("部署时间", ""))]
    date_span = f"{dates[0]} ~ {dates[-1]}" if dates else "-"

    print("")
    print(color("────────── 总体统计 / Overall Summary ──────────", "bold"))
    print(f"  平台数（设备类型）：{color(str(len(platforms)), 'bold')}"
          f"    批次数：{color(str(len(all_rows)), 'bold')}"
          f"（有效 {active_n} / 作废 {void_n}）"
          f"    部署时间跨度：{date_span}")
    print(f"  有效 Rule ID 总数：{color(str(len(active_rids)), 'green')}"
          f"    范围：{color(rid_span, 'green')}")
    if owners:
        print(f"  Owner：{', '.join(owners)}")
    if platforms:
        # 每平台的有效批次数，便于一眼看分布。
        per = {}
        for r in all_rows:
            if str(r.get("状态", STATUS_ACTIVE)) == STATUS_ACTIVE:
                per[r.get("设备类型代码", "")] = per.get(r.get("设备类型代码", ""), 0) + 1
        dist = "  ".join(f"{k}:{v}" for k, v in sorted(per.items()))
        print(f"  各平台有效批次：{dist}")

    print(ok(f"CSV 导出：{_csv_path()}"))
    return 0


def _selftest() -> int:
    """最小纯逻辑自测：仅写 OTA_DEPLOY_DB_DIR 指向的临时库，绝不触碰真实部署库。

    覆盖 append_record 的 force_status 语义与身份键三分支：
      - 不传 force_status → 状态『有效』；
      - force_status=STATUS_VOID → 状态『作废』且『作废时间』非空；
      - 同身份键不同 Rule ID → 旧『有效』记录被作废。
    """
    import tempfile

    failures: list[str] = []

    def check(cond: bool, msg: str) -> None:
        if cond:
            print(f"  PASS: {msg}")
        else:
            failures.append(msg)
            print(f"  FAIL: {msg}")

    tmp = tempfile.mkdtemp(prefix="depdb_selftest_")
    old = os.environ.get("OTA_DEPLOY_DB_DIR")
    os.environ["OTA_DEPLOY_DB_DIR"] = tmp
    try:
        print(f"[selftest] 临时库目录 = {tmp}")

        base = {
            "设备类型代码": "DEV_X",
            "源版本": "1.0.0",
            "目标版本": "2.0.0",
            "设备扩展码集合": "A,B",
            "Rule ID列表": "100,101",
            "Rule ID范围": "100 ~ 101",
            "规则条数": 2,
        }

        # 1) 不传 force_status → 有效。
        wrote, _ = append_record(dict(base))
        check(wrote, "默认写入返回新写入")
        recs = load_all()
        check(len(recs) == 1 and recs[-1].get("状态") == STATUS_ACTIVE,
              "默认写入状态为『有效』")

        # 2) force_status=STATUS_VOID → 作废且作废时间非空（新身份键，避免与情况3耦合）。
        void_rec = dict(base)
        void_rec["设备类型代码"] = "DEV_Y"
        wrote, _ = append_record(void_rec, force_status=STATUS_VOID)
        check(wrote, "VOID 写入返回新写入")
        recs = load_all()
        last = recs[-1]
        check(last.get("状态") == STATUS_VOID, "force_status=VOID 写入状态为『作废』")
        check(bool(str(last.get("作废时间", "")).strip()), "VOID 记录『作废时间』非空")

        # 3) 同身份键不同 Rule ID → 旧『有效』记录被作废。
        regen = dict(base)
        regen["Rule ID列表"] = "200,201"
        regen["Rule ID范围"] = "200 ~ 201"
        wrote, _ = append_record(regen)
        check(wrote, "同身份键重新生成返回新写入")
        recs = load_all()
        dev_x = [r for r in recs if r.get("设备类型代码") == "DEV_X"]
        old_active = [r for r in dev_x if r.get("Rule ID列表") == "100,101"]
        new_active = [r for r in dev_x if r.get("Rule ID列表") == "200,201"]
        check(len(old_active) == 1 and old_active[0].get("状态") == STATUS_VOID,
              "同身份键旧 Rule ID 记录被作废")
        check(len(new_active) == 1 and new_active[0].get("状态") == STATUS_ACTIVE,
              "同身份键新 Rule ID 记录为『有效』")

        if failures:
            print(f"[selftest] 失败 {len(failures)} 项。")
            return 1
        print("[selftest] 全部 PASS。")
        return 0
    finally:
        if old is None:
            os.environ.pop("OTA_DEPLOY_DB_DIR", None)
        else:
            os.environ["OTA_DEPLOY_DB_DIR"] = old


def main() -> int:
    """history 查看入口：python -m automation.depdb [--device CODE] [--limit N] [--export]"""
    import argparse
    import sys
    p = argparse.ArgumentParser(prog="deployota history",
                                description="查看 OTA 部署历史（JSONL 主存，CSV 可导出）。")
    p.add_argument("--device", default=None, help="按设备类型代码过滤（子串匹配）。")
    p.add_argument("--owner", default=None, help="按 Owner 过滤（子串匹配）。")
    p.add_argument("--rule", default=None,
                   help="按 Rule ID 或范围过滤（纯数字会判断是否落在某批区间内；也支持范围文本子串）。")
    p.add_argument("--limit", type=int, default=20, help="只显示最近 N 条（默认 20；0 表示全部）。")
    p.add_argument("--export", action="store_true", help="重新导出一份 CSV 后退出。")
    p.add_argument("--selftest", action="store_true",
                   help="运行纯逻辑自测（仅写临时库，绝不触碰真实部署库）后退出。")
    args = p.parse_args()
    if args.selftest:
        return _selftest()
    if args.export:
        _export_csv()
        print(f"已导出 CSV → {_csv_path()}")
        return 0
    return _print_history(args.limit or None, args.device, args.owner, args.rule)


if __name__ == "__main__":
    import sys
    sys.exit(main())
