"""deployota update —— 按 Rule ID 更新 DASS 的 OTA 阶段（原来 `dass otaXX <Excel>` 的 Rule ID 版）。

用法：
    ./deployota update ota200 65370                  # 单个 Rule ID
    ./deployota update ota200 65370-65379            # 区间
    ./deployota update ota1k  65370,65375,65380      # 逗号分隔（区间可混用）
    ./deployota update ota50  65370-65372 65380      # 也可给多个参数

OPTION：ota50 / ota200 / ota1k / ota5k / ota20k / ota-all（init 仍用 `./deployota dass init`，它要 Excel）。

做法：
  1) 把参数解析成 Rule ID 列表；
  2) 在部署库（depdb）里找到每个 Rule ID 所属的有效批次，取出它的 方案号(设备类型代码) /
     完整目标版本 / 设备扩展信息 等：优先用批次里存的『规则明细』；老批次没有明细就去读来源 Excel；
  3) 打印解析结果，然后交给 dass_fill 现有的更新逻辑（ota200 起先过成功率闸门；
     Stage1 方案号+目标版本+上一阶段 粗筛，Stage2 加设备扩展信息精选；写当前阶段/当前阶段时间；
     最后仍由你在终端输入 approve 才提交）。

解析不出来的 Rule ID（库里没有 / 只在作废批次里 / 没有明细且 Excel 不在）会逐个列出原因，
默认整体中止（什么都不改）；加 --skip-missing 则只处理能解析的那些。

成功率闸门（ota200 起）会逐个显示每个 Rule ID 的成功率。不达标时默认中止；在交互终端里会问你是否
『人工批复』继续——只有原样输入 override 才继续（之后仍要你核对并输入 approve 才提交），直接回车就中止。
加 --strict 则完全不询问、不达标一律中止。
"""
from __future__ import annotations

import os
import sys

from . import depdb
from ._log import color, err, ok, warn

OPTIONS = ["ota50", "ota200", "ota1k", "ota5k", "ota20k", "ota-all"]
_DETAIL_KEYS = ("ruleid", "方案号", "目标版本", "设备扩展信息", "机型", "源版本", "升级类型", "featurecode")


def resolve_rules(rule_ids: list[str], batches: list[dict], read_excel=None):
    """纯逻辑：把 Rule ID 解析成 DASS 更新用的记录。

    返回 (resolved, problems)：
      resolved = [(记录dict, 来源说明), ...]，顺序同 rule_ids；
      problems = [(rule_id, 原因), ...]。
    选批次：包含该 Rule ID 的批次里优先『有效』（多个取部署时间最新）；只有『作废』批次 → 记问题。
    取明细：批次『规则明细』里有该 Rule ID 就用；否则 read_excel(来源路径) 里按 ruleid 找；都没有 → 记问题。
    read_excel 默认读本机 Excel（可注入，便于自测）。
    """
    if read_excel is None:
        read_excel = _read_excel_by_rule
    resolved: list[tuple[dict, str]] = []
    problems: list[tuple[str, str]] = []
    excel_cache: dict[str, dict | None] = {}
    for rid in rule_ids:
        owners = [b for b in batches
                  if rid in [x.strip() for x in str(b.get("Rule ID列表", "")).split(",")]]
        if not owners:
            problems.append((rid, "部署库里没有这个 Rule ID（先部署入库，或用 `./deployota backfill` 回补）"))
            continue
        active = [b for b in owners if str(b.get("状态", depdb.STATUS_ACTIVE)) == depdb.STATUS_ACTIVE]
        if not active:
            problems.append((rid, "只出现在『作废』批次里，不能更新"))
            continue
        # 同一个 Rule ID 可能落在多个有效批次里（例如本机 Excel 入库一次、CMS 回补又入库一次，
        # 身份键不同没被合并）。按部署时间从新到旧逐个试，用第一个能取到明细的；都不行才记问题。
        rec, src, tried = None, "", []
        for batch in sorted(active, key=lambda b: str(b.get("部署时间", "")), reverse=True):
            label = batch.get("批次号", "")
            detail = next((d for d in (batch.get("规则明细") or []) if str(d.get("ruleid", "")).strip() == rid), None)
            src_try = f"部署库明细 {label}"
            if detail is None:
                path = str(batch.get("Excel路径", "")).strip()
                if path not in excel_cache:
                    excel_cache[path] = read_excel(path)
                rows = excel_cache[path]
                detail = (rows or {}).get(rid)
                src_try = f"来源 Excel {os.path.basename(path)}"
                if detail is None:
                    why = ("批次没有存规则明细，且来源 Excel 不可用" if rows is None else "来源 Excel 里没有这一行")
                    tried.append(f"{why}（批次 {label}，来源 {path or '-'}）")
                    continue
            cand = {k: str(detail.get(k, "") or "") for k in _DETAIL_KEYS}
            cand["ruleid"] = rid
            # 方案号缺失时用批次的设备类型代码兜底。
            cand["方案号"] = cand["方案号"] or str(batch.get("设备类型代码", ""))
            miss = [k for k in ("方案号", "目标版本") if not cand[k]]
            if miss:
                tried.append(f"明细缺少 {miss}，无法作为 DASS 查询条件（批次 {label}）")
                continue
            rec, src = cand, src_try
            break
        if rec is None:
            problems.append((rid, "；".join(tried) or "没有可用的批次"))
            continue
        resolved.append((rec, src))
    return resolved, problems


def _read_excel_by_rule(path: str) -> dict | None:
    """读来源 Excel → {ruleid: 记录}；不是本机 .xlsx / 读不了 → None。"""
    if not str(path).lower().endswith((".xlsx", ".xlsm")) or not os.path.isfile(path):
        return None
    try:
        from .dass_data import read_updated_excel
        return {str(r["ruleid"]).strip(): r for r in read_updated_excel(path) if str(r.get("ruleid", "")).strip()}
    except Exception:
        return None


def _print_plan(option: str, resolved) -> None:
    from .dass_fill import STAGE_LABELS
    print(color(f"—— 将更新 DASS：{len(resolved)} 条 Rule ID → 阶段 {STAGE_LABELS[option]} ——", "bold"))
    print(f"  {'Rule ID':8} {'方案号':20} {'目标版本':14} {'设备扩展信息':18} 来源")
    for rec, src in resolved:
        print(f"  {rec['ruleid']:8} {rec['方案号']:20} {rec['目标版本']:14} {rec['设备扩展信息'] or '-':18} {src}")


def main() -> int:
    argv = sys.argv[1:]
    if "--selftest" in argv:
        return _selftest()
    if not argv or any(a in ("-h", "--help") for a in argv):
        print(__doc__)
        return 0 if argv else 2

    skip_missing = "--skip-missing" in argv
    strict = "--strict" in argv
    option = None
    specs: list[str] = []
    for a in argv:
        if a in ("--skip-missing", "--strict"):
            continue
        if a == "init":
            print(err("init 需要 Excel（要插入新行），请用：./deployota dass init"), file=sys.stderr)
            return 2
        if a in OPTIONS:
            if option and option != a:
                print(err(f"只能指定一个阶段，已有 {option!r} 又遇到 {a!r}。"), file=sys.stderr)
                return 2
            option = a
        elif a.startswith("-") and not a[1:2].isdigit():
            print(err(f"未知参数：{a!r}。用法见 -h。"), file=sys.stderr)
            return 2
        else:
            specs.append(a)
    if not option:
        print(err(f"缺少阶段。用法：./deployota update <{'|'.join(OPTIONS)}> <Rule ID / 区间 / 逗号分隔>"),
              file=sys.stderr)
        return 2
    if not specs:
        print(err("缺少 Rule ID。例：./deployota update ota200 65370-65379"), file=sys.stderr)
        return 2

    from .dass_perf_query import parse_ruleid_spec
    ids: list[str] = []
    try:
        for s in specs:
            ids.extend(parse_ruleid_spec(s))
    except ValueError as e:
        print(err(f"Rule ID 参数有误：{e}"), file=sys.stderr)
        return 2
    ids = list(dict.fromkeys(ids))

    batches = depdb.load_all()
    resolved, problems = resolve_rules(ids, batches)
    print(color(f"update {option}：查询 {len(ids)} 个 Rule ID，解析到 {len(resolved)} 个。", "bold"))
    if problems:
        print(warn(f"有 {len(problems)} 个 Rule ID 无法解析："))
        for rid, why in problems:
            print(f"  · {rid}：{why}")
        if not skip_missing:
            print(err("已中止，什么都没改。修正后重跑，或加 --skip-missing 只处理能解析的。"), file=sys.stderr)
            return 2
        print(warn("--skip-missing：跳过这些，继续处理其余的。"))
    if not resolved:
        print(err("没有可更新的 Rule ID。"), file=sys.stderr)
        return 2

    _print_plan(option, resolved)
    from .dass_fill import execute
    return execute(option, [rec for rec, _ in resolved], None, allow_gate_override=not strict)


def _selftest() -> int:
    good = True

    def check(name, cond):
        nonlocal good
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        good = good and bool(cond)

    det = {"ruleid": "65370", "方案号": "DEV", "目标版本": "0100UQ1104", "设备扩展信息": "55U73"}
    batches = [
        {"批次号": "A", "Rule ID列表": "65370,65371", "状态": "有效", "部署时间": "20261109",
         "设备类型代码": "DEV", "Excel路径": "/x/a.xlsx", "规则明细": [det]},
        {"批次号": "OLD", "Rule ID列表": "65371", "状态": "作废", "部署时间": "20261001", "设备类型代码": "DEV"},
        {"批次号": "V", "Rule ID列表": "70000", "状态": "作废", "部署时间": "20261001", "设备类型代码": "DEV"},
    ]
    excel = {"65371": {"ruleid": "65371", "方案号": "DEV", "目标版本": "0100UQ1104", "设备扩展信息": "65U73"}}
    r, p = resolve_rules(["65370", "65371", "70000", "99999"], batches, read_excel=lambda path: excel)
    check("有明细直接用", r[0][0]["设备扩展信息"] == "55U73" and "部署库明细" in r[0][1])
    check("无明细回退读 Excel", r[1][0]["设备扩展信息"] == "65U73" and "来源 Excel" in r[1][1])
    check("只在作废批次 → 问题", ("70000", "只出现在『作废』批次里，不能更新") in p)
    check("库里没有 → 问题", any(x[0] == "99999" for x in p))
    r2, p2 = resolve_rules(["65371"], batches, read_excel=lambda path: None)
    check("Excel 不可用 → 问题", r2 == [] and "来源 Excel 不可用" in p2[0][1])
    r3, p3 = resolve_rules(["65371"], batches, read_excel=lambda path: {})
    check("Excel 里没这行 → 问题", r3 == [] and "没有这一行" in p3[0][1])
    bad = [{"批次号": "B", "Rule ID列表": "1", "状态": "有效", "设备类型代码": "",
            "规则明细": [{"ruleid": "1", "方案号": "", "目标版本": ""}]}]
    r4, p4 = resolve_rules(["1"], bad)
    check("缺方案号/目标版本 → 问题", r4 == [] and "缺少" in p4[0][1])
    # 同一 Rule ID 在两个有效批次：较新的没有明细，较旧的有 Excel → 应回退用旧的
    overlap = [
        {"批次号": "NEW_CMS", "Rule ID列表": "7,8", "状态": "有效", "部署时间": "20260811",
         "设备类型代码": "DEV", "Excel路径": "CMS:x"},
        {"批次号": "OLD_XLS", "Rule ID列表": "6,7,8", "状态": "有效", "部署时间": "20260809",
         "设备类型代码": "DEV", "Excel路径": "/x/old.xlsx"},
    ]
    xl = {k: {"ruleid": k, "方案号": "DEV", "目标版本": "T", "设备扩展信息": "E"} for k in ("6", "7", "8")}
    r5, p5 = resolve_rules(["6", "7", "8"], overlap, read_excel=lambda path: xl if path.endswith(".xlsx") else None)
    check("重叠批次：新批次无明细时回退到旧批次的 Excel", len(r5) == 3 and p5 == [] and all("old.xlsx" in x[1] for x in r5))
    r6, p6 = resolve_rules(["7"], overlap, read_excel=lambda path: None)
    check("重叠批次都没有明细 → 问题里列出全部尝试", r6 == [] and "NEW_CMS" in p6[0][1] and "OLD_XLS" in p6[0][1])
    print("\n自测结果：", "全部通过 ✅" if good else "存在失败 ❌")
    return 0 if good else 1


if __name__ == "__main__":
    sys.exit(main())
