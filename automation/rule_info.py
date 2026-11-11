"""按 Rule ID 查询部署信息（只读，查本地部署库）。

用法：
    ./deployota query 65370                    # 单个
    ./deployota query 65370-65379              # 区间（闭区间）
    ./deployota query 65370,65375,65380        # 逗号分隔
    ./deployota query 65370-65372,65380        # 区间与逗号可混用
    ./deployota query 65370 65380-65383        # 也可给多个参数

对每个 Rule ID 找到它所属的部署批次（可能同时落在一个『作废』批次和一个『有效』批次里，都会列出），
同一批次里的多个 Rule ID 合并成一段显示（如 65370~65379），批次信息只打印一次：
设备类型代码 / 源版本 → 目标版本 / 升级类型 / 状态 / 部署时间 / Owner / 设备扩展码 / 来源。
库里没有的 Rule ID 单独列在末尾。

数据来自 depdb（每次 dass 更新后自动入库；历史批次可用 `deployota backfill` 回补）。
部署库按『批次』存，只记批次第一行的源版本/目标版本；所以当来源 Excel 还在时，会再按行读出
每个 Rule ID 自己的 机型 / 设备扩展信息 / 源版本 → 目标版本 / 升级类型（逐条明细）。
来源不是本机 Excel（如 CMS 回补的批次）或文件已不在时，只显示批次信息并给出提示。
"""
from __future__ import annotations

import sys

from . import depdb
from ._log import color, err, ok, warn

# 区间展开的数量上限：防止误输入 1-99999999 把终端刷爆。
MAX_IDS = 5000


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


def show(spec_args: list[str]) -> int:
    """解析 Rule ID 规格 → 查部署库 → 打印。返回 0 全部找到；1 有未找到的；2 参数错误。"""
    from .dass_perf_query import parse_ruleid_spec  # 复用：区间 / 逗号 / 千分位 解析

    ids: list[str] = []
    try:
        for spec in spec_args:
            ids.extend(parse_ruleid_spec(spec))
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
        print(color("  （部署库只含已入库的批次：dass 更新后自动入库；更早的批次可用 `./deployota backfill` 回补）",
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
