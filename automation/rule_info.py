"""deployota query <Rule ID... | 月份> —— 查部署信息（只读，查本地部署库）

  ./deployota query 65370                 单个
  ./deployota query 65370-65379,65385     区间 / 逗号，可混用
  ./deployota query 2026.04               一个月：列出所有批次和 Rule ID，统计批次数 / 规则条数
  ./deployota query 2026.04-2026.10       月份区间（也可写 2026/04、2026年04月）

按 Rule ID 查：所属批次（设备类型代码、目标完整版本、升级类型、部署时间、Owner…）、
每个 Rule ID 自己的机型 / 设备扩展信息 / 版本（来源 Excel 在时），以及灰度阶段时间线和 CMS 状态。
退出码：0 全部找到；1 有找不到的 / 该时间段没有部署；2 参数有误。
"""
from __future__ import annotations

import sys

from . import depdb
from ._log import color, err, ok, warn

# 区间展开的数量上限：防止误输入 1-99999999 把终端刷爆。
MAX_IDS = 5000

import re as _re

# 月份写法：2026.04 / 2026/04 / 2026年04月（年份 4 位，月份 1~2 位）。
_MONTH = r"(\d{4})\s*[./年]\s*(\d{1,2})\s*月?"
_MONTH_RE = _re.compile(rf"^{_MONTH}$")
_MONTH_RANGE_RE = _re.compile(rf"^{_MONTH}\s*[-~～至到]\s*{_MONTH}$")


def is_month_spec(spec: str) -> bool:
    s = str(spec).strip()
    return bool(_MONTH_RE.match(s) or _MONTH_RANGE_RE.match(s))


def parse_month_spec(spec: str) -> tuple[str, str]:
    """'2026.04' -> ('202604','202604')；'2026.04-2026.10' -> ('202604','202610')。非法抛 ValueError。"""
    s = str(spec).strip()
    m = _MONTH_RANGE_RE.match(s)
    if m:
        y1, m1, y2, m2 = (int(x) for x in m.groups())
    else:
        m = _MONTH_RE.match(s)
        if not m:
            raise ValueError(f"看不懂的月份：{spec!r}（例：2026.04 或 2026.04-2026.10）")
        y1, m1 = (int(x) for x in m.groups())
        y2, m2 = y1, m1
    for mm in (m1, m2):
        if not 1 <= mm <= 12:
            raise ValueError(f"月份必须在 1~12：{spec!r}")
    a, b = f"{y1:04d}{m1:02d}", f"{y2:04d}{m2:02d}"
    if a > b:
        raise ValueError(f"起始月份晚于结束月份：{spec!r}")
    return a, b


def batches_in_months(rows: list[dict], start: str, end: str) -> list[dict]:
    """部署时间（YYYYMMDD）落在 [start, end] 月份内的批次，按部署时间升序（同日保持库内顺序）。"""
    hit = [(str(r.get("部署时间", ""))[:8], i, r) for i, r in enumerate(rows)
           if start <= str(r.get("部署时间", ""))[:6] <= end]
    return [r for _, _, r in sorted(hit, key=lambda x: (x[0], x[1]))]


def month_stats(batches: list[dict]) -> dict:
    """统计（纯逻辑）：有效/作废批次数、有效规则条数、去重后的有效 Rule ID 数、平台数、Rule ID 范围。"""
    active = [b for b in batches if str(b.get("状态", depdb.STATUS_ACTIVE)) == depdb.STATUS_ACTIVE]
    rules = 0
    uniq: set[str] = set()
    for b in active:
        ids = [x.strip() for x in str(b.get("Rule ID列表", "")).split(",") if x.strip()]
        try:
            rules += int(b.get("规则条数") or len(ids))
        except (TypeError, ValueError):
            rules += len(ids)
        uniq.update(ids)
    nums = sorted(int(x) for x in uniq if x.isdigit())
    return {
        "batches": len(active),
        "void": len(batches) - len(active),
        "rules": rules,
        "unique": len(uniq),
        "platforms": len({str(b.get("设备类型代码", "")).strip() for b in active if str(b.get("设备类型代码", "")).strip()}),
        "span": f"{nums[0]} ~ {nums[-1]}" if nums else "-",
    }


def _batch_line(r: dict) -> str:
    """与 `deployota history` 同一行格式。"""
    void = str(r.get("状态", depdb.STATUS_ACTIVE)) == depdb.STATUS_VOID
    tag = color("[作废]", "red") if void else color("[有效]", "green")
    tgt, is_full = depdb.display_target(r)
    tgt = tgt if is_full else f"{tgt}（版本变化，无完整版本）"
    line = (f"· Rule ID {r.get('Rule ID范围', ''):15}（{r.get('规则条数', '')} 条）  {tag}  "
            f"{r.get('设备类型代码', ''):18} {r.get('源版本', '')} → {tgt}  "
            f"[{r.get('升级类型', '')}]  owner={r.get('Owner', '') or '-'}  {r.get('部署时间', '')}")
    return color(line, "dim") if void else line


def _ym(s: str) -> str:
    return f"{s[:4]}.{s[4:6]}"


def show_months(specs: list[str]) -> int:
    """按月份列出部署批次与 Rule ID，并给出每月小计和总计。返回 0 有数据；1 该时间段没有部署。"""
    try:
        ranges = [parse_month_spec(x) for x in specs]
    except ValueError as e:
        print(err(str(e)), file=sys.stderr)
        return 2
    rows = depdb.load_all()
    seen: set[int] = set()
    batches: list[dict] = []
    for a, b in ranges:
        for r in batches_in_months(rows, a, b):
            if id(r) not in seen:
                seen.add(id(r))
                batches.append(r)
    batches.sort(key=lambda r: str(r.get("部署时间", "")))
    label = "、".join(_ym(a) if a == b else f"{_ym(a)} ~ {_ym(b)}" for a, b in ranges)
    print(color(f"查询部署时间：{label}    部署库：{depdb.db_dir()}（{len(rows)} 个批次）", "bold"))
    if not batches:
        print(warn("这段时间没有部署记录。"))
        return 1
    by_month: dict[str, list[dict]] = {}
    for r in batches:
        by_month.setdefault(str(r.get("部署时间", ""))[:6], []).append(r)
    for mon, items in by_month.items():
        st = month_stats(items)
        void = f"，另有作废 {st['void']} 批" if st["void"] else ""
        print("")
        print(color(f"—— {_ym(mon)}：{st['batches']} 个批次，共 {st['rules']} 条规则{void} ——", "bold"))
        for r in items:
            print(_batch_line(r))
    st = month_stats(batches)
    print("")
    print(color("────────── 汇总 ──────────", "bold"))
    print(f"  时间段：{label}    批次数：{color(str(st['batches']), 'bold')}（有效）"
          + (f" + 作废 {st['void']}" if st["void"] else "")
          + f"    规则条数：{color(str(st['rules']), 'green')}"
          + f"    平台数：{st['platforms']}")
    dup = st["rules"] - st["unique"]
    print(f"  去重后 Rule ID：{st['unique']} 个    范围：{st['span']}"
          + (color(f"    （有 {dup} 条在多个有效批次里重复记录，规则条数按批次累计）", "yellow") if dup > 0 else ""))
    if len(by_month) > 1:
        print("  按月：" + "   ".join(
            f"{_ym(m)} {month_stats(v)['batches']}批/{month_stats(v)['rules']}条" for m, v in by_month.items()))
    return 0


def compress_ids(ids) -> str:
    """把一组数字 Rule ID 压成『a~b, c, d~e』（连续的合并，升序）。非数字原样保留在末尾。"""
    nums = sorted({int(x) for x in ids if str(x).isdigit()})
    others = sorted({str(x) for x in ids if not str(x).isdigit()})
    parts: list[str] = []
    i = 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        parts.append(str(nums[i]) if i == j else f"{nums[i]}~{nums[j]}")
        i = j + 1
    return ", ".join(parts + others)


def lookup(rule_ids: list[str], rows: list[dict]):
    """纯逻辑：把 rule_ids 匹配到 rows 里的批次。

    返回 (groups, missing)：
      groups = [(批次记录, [命中的 Rule ID...]), ...]，按『部署时间』升序、同日保持库内顺序；
      missing = 没有任何批次包含的 Rule ID（保持输入顺序）。
    匹配只认批次的『Rule ID列表』里精确出现的号（不按区间首尾猜中间号）。
    """
    want = list(dict.fromkeys(str(x).strip() for x in rule_ids if str(x).strip()))
    index: dict[str, list[int]] = {}
    for bi, rec in enumerate(rows):
        for x in str(rec.get("Rule ID列表", "")).split(","):
            x = x.strip()
            if x:
                index.setdefault(x, []).append(bi)
    hit: dict[int, list[str]] = {}
    missing: list[str] = []
    for rid in want:
        bis = index.get(rid)
        if not bis:
            missing.append(rid)
            continue
        for bi in bis:
            hit.setdefault(bi, []).append(rid)
    order = sorted(hit, key=lambda bi: (str(rows[bi].get("部署时间", "")), bi))
    return [(rows[bi], hit[bi]) for bi in order], missing


_EXCEL_CACHE: dict[str, dict] = {}


def _excel_rows_by_rule(path: str) -> dict | None:
    """读来源 Excel，返回 {ruleid: 行记录}；不是本机 .xlsx / 读不了 / 没有 Rule ID 列 → None。"""
    import os
    path = str(path or "").strip()
    if not path.lower().endswith((".xlsx", ".xlsm")) or not os.path.isfile(path):
        return None
    if path not in _EXCEL_CACHE:
        try:
            from .dass_data import read_updated_excel
            _EXCEL_CACHE[path] = {r["ruleid"]: r for r in read_updated_excel(path) if r.get("ruleid")}
        except Exception:
            _EXCEL_CACHE[path] = {}
    return _EXCEL_CACHE[path] or None


DETAIL_MAX = 60  # 逐条明细最多显示行数


def _print_rule_rows(rec: dict, hit_ids: list[str]) -> None:
    """逐条明细：来源 Excel 在就按行显示每个 Rule ID 自己的信息，否则给出提示。"""
    by_rule = _excel_rows_by_rule(rec.get("Excel路径", ""))
    if by_rule is None:
        print(color("    （来源 Excel 不可用：上面的版本是批次首行记录，若该批次混有多个版本，"
                    "不一定适用于每个 Rule ID）", "dim"))
        return
    print("    逐条明细（来自来源 Excel）：")
    shown = 0
    for rid in sorted(hit_ids, key=lambda x: int(x) if x.isdigit() else 0):
        r = by_rule.get(rid)
        if shown >= DETAIL_MAX:
            print(color(f"      … 还有 {len(hit_ids) - shown} 个未显示", "dim"))
            break
        shown += 1
        if r is None:
            print(f"      {rid}  " + color("来源 Excel 里没有这一行", "yellow"))
            continue
        print(f"      {rid}  机型={r.get('机型') or '-'}  设备扩展信息={r.get('设备扩展信息') or '-'}  "
              f"{r.get('源版本') or '-'} → {r.get('目标版本') or '-'}  [{r.get('升级类型') or '-'}]")


_STAGE_STATE: dict | None = None


def _print_stage_lines(hit_ids: list[str]) -> None:
    """每个 Rule ID 的灰度阶段时间线与 CMS 在线状态（来自 rule_stages.jsonl）。"""
    global _STAGE_STATE
    from . import stage_db
    if _STAGE_STATE is None:
        _STAGE_STATE = stage_db.build_state(stage_db.load_events())
    print("    灰度阶段：")
    for rid in sorted(hit_ids, key=lambda x: int(x) if x.isdigit() else 0):
        st = _STAGE_STATE.get(rid)
        if not st:
            print(color(f"      {rid}  （没有阶段记录）", "dim"))
            continue
        note = f"，{st['CMS备注']}" if st.get("CMS备注") else ""
        cms = f"  CMS={st['CMS状态']}（{st['CMS状态时间'][:16]}{note}）" if st["CMS状态"] else ""
        print(f"      {rid}  当前={st['当前阶段'] or '-'}  {stage_db.timeline(st) or '-'}{cms}")
        for line in stage_db.conflict_lines(st):
            print(color(f"             冲突：{line}", "yellow"))


def _print_group(rec: dict, hit_ids: list[str]) -> None:
    void = str(rec.get("状态", depdb.STATUS_ACTIVE)) == depdb.STATUS_VOID
    tag = color("[作废]", "red") if void else color("[有效]", "green")
    total = rec.get("规则条数", "")
    print(f"{tag} Rule ID {color(compress_ids(hit_ids), 'bold')}"
          f"（命中 {len(hit_ids)} 个；该批次共 {total} 条，范围 {rec.get('Rule ID范围', '')}）")
    ext = [x.strip() for x in str(rec.get("设备扩展码集合", "")).split(",") if x.strip()]
    rows = [
        ("设备类型代码", rec.get("设备类型代码", "")),
        ("设备类型名称", rec.get("设备类型名称", "") or "-"),
        ("目标完整版本", (lambda t: t[0] if t[1] else "-（无完整版本）")(depdb.display_target(rec))),
        ("版本变化(批次首行)", f"{rec.get('源版本', '')} → {rec.get('目标版本', '')}"),
        ("升级类型", rec.get("升级类型", "") or "-"),
        ("部署时间", rec.get("部署时间", "") or "-"),
        ("Owner", rec.get("Owner", "") or "-"),
        ("设备扩展码", (", ".join(ext) + f"（{len(ext)} 个）") if ext else "-"),
        ("批次号", rec.get("批次号", "") or "-"),
        ("来源", rec.get("Excel路径", "") or "-"),
        ("入库时间", rec.get("入库时间", "") or "-"),
    ]
    if void:
        rows.insert(4, ("作废时间", rec.get("作废时间", "") or "-"))
    for k, v in rows:
        print(f"    {k}：{v}")
    _print_rule_rows(rec, hit_ids)
    _print_stage_lines(hit_ids)


def show(spec_args: list[str]) -> int:
    """入口：月份参数 → 按月份列出；否则按 Rule ID 查。两种不能混用。"""
    import os
    home = os.path.expanduser("~")
    spec_args = ["~" if a in (home, home + "/") else a for a in spec_args]
    joined = " ".join(spec_args).strip()
    if is_month_spec(joined):            # 带空格的月份区间：2026.04 ~ 2026.10
        return show_months([joined])
    months = [x for x in spec_args if is_month_spec(x)]
    if months:
        if len(months) != len(spec_args):
            print(err("月份和 Rule ID 不能混在一次查询里，请分开查。"), file=sys.stderr)
            return 2
        return show_months(months)
    return show_rules(spec_args)


def show_rules(spec_args: list[str]) -> int:
    """解析 Rule ID 规格 → 查部署库 → 打印。返回 0 全部找到；1 有未找到的；2 参数错误。"""
    from .dass_perf_query import parse_ruleid_args  # 复用：区间（- ~ ～）/ 逗号 / 空格 / 千分位
    ids: list[str] = []
    try:
        ids = parse_ruleid_args(spec_args)
    except ValueError as e:
        print(err(f"Rule ID 参数有误：{e}"), file=sys.stderr)
        return 2
    ids = list(dict.fromkeys(ids))
    if not ids:
        print(err("没有解析到任何 Rule ID。"), file=sys.stderr)
        return 2
    if len(ids) > MAX_IDS:
        print(err(f"一次最多查 {MAX_IDS} 个 Rule ID，当前解析出 {len(ids)} 个，请缩小区间。"), file=sys.stderr)
        return 2
    short = [x for x in ids if len(x) < 5]
    if short:
        print(warn(f"注意：{short[:5]} 位数少于 5 位，Rule ID 通常是 5 位数字。"
                   "若本意是带逗号的单个号（如 64,700），请去掉逗号写成 64700。"))

    rows = depdb.load_all()
    print(color(f"查询 Rule ID：{compress_ids(ids)}（共 {len(ids)} 个）    部署库：{depdb.db_dir()}（{len(rows)} 个批次）",
                "bold"))
    groups, missing = lookup(ids, rows)
    print("")
    for rec, hit in groups:
        _print_group(rec, hit)
        print("")
    found = {x for _, hit in groups for x in hit}
    print(color("────────── 汇总 ──────────", "bold"))
    print(f"  查询 {len(ids)} 个：" + ok(f"库中找到 {len(found)} 个") +
          (("，" + warn(f"未找到 {len(missing)} 个")) if missing else ""))
    if missing:
        print(f"  未找到：{compress_ids(missing)}")
        print(color("  （部署库只含已入库的批次：publish / dass init 完成后自动入库）",
                    "dim"))
    return 1 if missing else 0


def _selftest() -> int:
    good = True

    def check(name, cond):
        nonlocal good
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        good = good and bool(cond)

    check("compress 连续合并", compress_ids(["65372", "65370", "65371"]) == "65370~65372")
    check("compress 断开", compress_ids(["65370", "65372", "65373", "65380"]) == "65370, 65372~65373, 65380")
    check("compress 去重", compress_ids(["1", "1"]) == "1")
    rows = [
        {"批次号": "A", "Rule ID列表": "65370,65371,65372", "部署时间": "20261109", "状态": "有效"},
        {"批次号": "B", "Rule ID列表": "65380,65381", "部署时间": "20261110", "状态": "有效"},
        {"批次号": "C", "Rule ID列表": "65371", "部署时间": "20261001", "状态": "作废"},
    ]
    g, m = lookup(["65371", "65380", "99999"], rows)
    check("lookup 命中与未命中", [r["批次号"] for r, _ in g] == ["C", "A", "B"] and m == ["99999"])
    check("lookup 同号落在作废+有效两批", {r["批次号"] for r, h in g if "65371" in h} == {"A", "C"})
    g2, m2 = lookup(["65373"], rows)  # 在 A 的区间之外：不能按首尾猜中间号
    check("lookup 只认列表里的号", g2 == [] and m2 == ["65373"])
    g3, _ = lookup(["65370", "65371", "65372"], rows[:1])
    check("lookup 同批合并成一组", len(g3) == 1 and g3[0][1] == ["65370", "65371", "65372"])
    check("月份：单月", parse_month_spec("2026.04") == ("202604", "202604"))
    check("月份：区间与多种写法", parse_month_spec("2026.04-2026.10") == ("202604", "202610")
          and parse_month_spec("2026/4~2026/10") == ("202604", "202610")
          and parse_month_spec("2026年04月-2026年10月") == ("202604", "202610"))
    check("月份：Rule ID 不会被当成月份", not is_month_spec("65370-65379") and not is_month_spec("65370"))
    for bad in ("2026.13", "2026.10-2026.04"):
        try:
            parse_month_spec(bad); raised = False
        except ValueError:
            raised = True
        check(f"月份：非法 {bad} 报错", raised)
    mrows = [
        {"部署时间": "20260427", "状态": "有效", "规则条数": 6, "Rule ID列表": "1,2,3,4,5,6", "设备类型代码": "A"},
        {"部署时间": "20260427", "状态": "有效", "规则条数": 1, "Rule ID列表": "7", "设备类型代码": "B"},
        {"部署时间": "20260428", "状态": "有效", "规则条数": 2, "Rule ID列表": "6,8", "设备类型代码": "A"},
        {"部署时间": "20260501", "状态": "作废", "规则条数": 3, "Rule ID列表": "9,10,11", "设备类型代码": "A"},
        {"部署时间": "20260601", "状态": "有效", "规则条数": 1, "Rule ID列表": "12", "设备类型代码": "C"},
    ]
    apr = batches_in_months(mrows, "202604", "202604")
    st = month_stats(apr)
    check("月份统计：批次数/规则条数/去重/平台", (st["batches"], st["rules"], st["unique"], st["platforms"]) == (3, 9, 8, 2))
    st2 = month_stats(batches_in_months(mrows, "202604", "202605"))
    check("月份统计：作废批次单列、不计入规则条数", st2["void"] == 1 and st2["rules"] == 9)
    print("\n自测结果：", "全部通过 ✅" if good else "存在失败 ❌")
    return 0 if good else 1


def main() -> int:
    argv = sys.argv[1:]
    if "--selftest" in argv:
        return _selftest()
    if not argv or any(a in ("-h", "--help") for a in argv):
        print(__doc__)
        return 0 if argv else 2
    return show(argv)


if __name__ == "__main__":
    sys.exit(main())
