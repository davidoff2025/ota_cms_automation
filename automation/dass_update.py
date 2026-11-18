"""deployota update <OPTION> <Rule ID...> —— 按 Rule ID 放量

  ./deployota update ota50  65370-65372 65380      Rule ID：单个 / 区间 / 逗号，可混用
  ./deployota update ota50  65370 --dry-run        测试模式：只在 CMS 走一遍流程，什么都不改
  ./deployota update ota200 65370-65379 --yes      跳过 CMS 执行前的 yes 确认
  ./deployota update 65370-65379                   不写阶段：按 CMS『阶段名称』自动进到下一阶段
                                                   （一批里阶段不同就分组，开始前确认一次，逐组执行）

OPTION（只前进）：ota50 / ota200 / ota1k / ota5k / ota20k / ota-all

流程：
  1) 按 Rule ID 从部署库取方案号、目标完整版本、设备扩展信息（老批次没明细时读来源 Excel）。
  2) 两道闸门（ota200 起；ota50 不检查）：
       成功率：报表里每个 Rule ID 的成功率都 > 50%（正好 50% 不算）；
       间隔：距进入上一阶段 > 48 小时（本地阶段记录；没有记录就用报表开放日期 +2 天）。
     两道都通过 → 全自动：不再问 yes、20K 判断直接执行、DASS 写完自动提交（加 --confirm 恢复逐步确认）。
     任一不通过 → 停下列出原因，原样输入 override 放行整组（--strict 则不询问、直接中止）。
  2b) 放 20K（ota20k）时，再按『源版本用户数』（同一张报表、同一行）整批判断：
     当前 5K：有一条 > 20000 → 放 20K；全部 ≤ 20000 → 直接完全开放（CMS 全量，DASS 完全开放）。
     当前已 20K：有一条 > 20000 → 再放一轮 20K（CMS 累计检测数 +20000，DASS 保持 OTA20000 不改）；
                 全部 ≤ 20000 → 完全开放。当前阶段取自阶段记录；判断结果要你输入 yes（--yes 跳过）。
     ota-all 不检查用户数。
  3) CMS「在线规则」灰度放开：按设备类型代码 + 批量规则ID 搜索，累计检测数必须是上一阶段的值
     （ota50 要求 5、ota200 要求 50 …），批量灰度 → 配置最大升级检测数选目标阶段 → 确定 →
     你输入 yes → 执行批量操作 → 回读累计检测数 = 目标值才算成功，并记录灰度阶段。
  4) DASS 更新『当前阶段 / 当前阶段时间』（只更新第 3 步成功的）。需要逐组处理时，每组写完会停下，
     输入 approve 提交本组并继续下一组。最后仍由你核对、输入 approve 提交。

选项：
  --dry-run       测试模式（不做闸门、不执行批量操作、不改 DASS）
  --confirm       每一步都要你确认（CMS 执行前 yes、20K 判断 yes、DASS approve）；默认闸门通过就全自动
  --yes           跳过 CMS 执行前的确认
  --auto-approve  跳过 DASS 提交前的 approve：全部写入成功才自动提交（含逐组提交）；有没成功的仍要你确认。
                  不会绕过成功率闸门（闸门不达标照样要 override 或中止）
  --no-cms        不动 CMS，只更新 DASS（以 DASS 更新成功记录阶段）
  --strict        闸门不达标直接中止，不询问人工批复
  --skip-missing  部署库里解析不出的 Rule ID 跳过，其余照常（默认整体中止）
  --mail-to "…"   放量汇总邮件额外收件人（固定 To 数据统计报告、Bcc devinw）
  --no-mail       不发放量汇总邮件（照样记录、照样存 HTML）

结束后：放量汇总（带灰度进度条）存到 OTADATA/reports/，发邮件给 数据统计报告（Bcc devinw）；
  有规则放开才发（失败的条目也写进去）；执行了但一条都没放开 → 发失败通知；测试模式 / 取消 / 无变化不发。
  每次 update 记录到 OTADATA/update_runs.jsonl，query / history 能看到灰度进度。
退出码：0 成功；1 出错；2 参数问题 / 闸门未通过 / 有解析不出的；4 CMS 一个都没放开。
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
            problems.append((rid, "部署库里没有这个 Rule ID（需先经 publish / dass init 入库）"))
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


def _print_plan(option: str, resolved, dry_run: bool = False) -> None:
    from .dass_fill import STAGE_LABELS
    if dry_run:
        print(color(f"—— 测试模式：{len(resolved)} 条 Rule ID（目标 {STAGE_LABELS[option]}），只在 CMS 走流程，不更新 DASS ——", "bold"))
    else:
        print(color(f"—— 将更新 DASS：{len(resolved)} 条 Rule ID → 阶段 {STAGE_LABELS[option]} ——", "bold"))
    print(f"  {'Rule ID':8} {'方案号':20} {'目标版本':14} {'设备扩展信息':18} 来源")
    for rec, src in resolved:
        print(f"  {rec['ruleid']:8} {rec['方案号']:20} {rec['目标版本']:14} {rec['设备扩展信息'] or '-':18} {src}")


def main() -> int:
    argv = sys.argv[1:]
    if "--selftest" in argv:
        return _selftest()
    from .mailer import pop_mail_to
    argv = pop_mail_to(argv)          # --mail-to "…"：放量汇总邮件额外收件人
    if not argv or any(a in ("-h", "--help") for a in argv):
        print(__doc__)
        return 0 if argv else 2

    skip_missing = "--skip-missing" in argv
    strict = "--strict" in argv
    no_cms = "--no-cms" in argv
    dry_run = "--dry-run" in argv
    assume_yes = "--yes" in argv
    no_mail = "--no-mail" in argv
    confirm = "--confirm" in argv
    if not confirm and not dry_run:
        # 全自动：两道闸门（成功率 > 50%、间隔 > 48 小时）都通过就不再问任何确认；不通过照样停下等 override。
        assume_yes = True
        from .dass_fill import enable_auto_approve
        enable_auto_approve()
        print(color("（全自动：成功率 > 50%、间隔 > 48 小时两道闸门都通过就直接执行 CMS 灰度和 DASS 更新并提交；"
                    "不通过会停下等你 override。加 --confirm 每一步都要你确认。）", "dim"))
    if "--auto-approve" in argv:
        from .dass_fill import enable_auto_approve
        enable_auto_approve()
    option = None
    specs: list[str] = []
    for a in argv:
        if a in ("--skip-missing", "--strict", "--no-cms", "--yes", "--dry-run", "--auto-approve", "--no-mail",
                 "--confirm"):
            continue
        if a == "init":
            print(err("init 需要 Excel（要插入新行），请用：./deployota dass init"), file=sys.stderr)
            return 2
        if a in OPTIONS:
            if option and option != a:
                print(err(f"只能指定一个阶段，已有 {option!r} 又遇到 {a!r}。"), file=sys.stderr)
                return 2
            option = a
        elif a.startswith("-") and a not in ("-", "--") and not a[1:2].isdigit():
            print(err(f"未知参数：{a!r}。用法见 -h。"), file=sys.stderr)
            return 2
        else:
            specs.append(a)
    if not specs:
        print(err("缺少 Rule ID。例：./deployota update ota200 65370-65379"), file=sys.stderr)
        return 2

    from .dass_perf_query import parse_ruleid_args
    ids: list[str] = []
    try:
        ids = parse_ruleid_args(specs)
    except ValueError as e:
        print(err(f"Rule ID 参数有误：{e}"), file=sys.stderr)
        return 2
    ids = list(dict.fromkeys(ids))
    if not option:
        # 没写阶段：按 CMS『阶段名称』自动进到下一阶段（一批里阶段不同就按阶段分组，一次确认后逐组执行）。
        if no_cms:
            print(err("--no-cms 时必须写阶段（自动判断要读 CMS）。"), file=sys.stderr)
            return 2
        groups = auto_plan(ids, assume_yes, dry_run)
        if groups is None:
            return 2
        reports, rc = [], 0
        for opt, gids in groups:
            r, report = _one(opt, gids, skip_missing, strict, no_cms, assume_yes, dry_run,
                             cms_yes=True, auto=True, title=f"update {opt}（自动下一阶段）")
            reports.append(report)
            rc = rc or r
        from .update_report import finalize
        finalize(reports, send=not no_mail)
        return rc
    r, report = _one(option, ids, skip_missing, strict, no_cms, assume_yes, dry_run)
    from .update_report import finalize
    finalize([report], send=not no_mail)
    return r


def _one(option, ids, skip_missing, strict, no_cms, assume_yes, dry_run, cms_yes=None, auto=False,
         title: str = "") -> tuple[int, dict | None]:
    """跑一组（一个阶段），返回 (退出码, 放量汇总)。汇总要在操作总结打印（finish）之前收集。"""
    from . import action_summary as A
    from . import cms_grey
    from .dass_perf import LAST_USERS
    from .update_report import collect
    cms_grey.LAST_RUN.clear()
    LAST_USERS.clear()
    A.start((title or f"update {option}") + ("（测试模式）" if dry_run else ""), ids)
    try:
        rc = _run(option, ids, skip_missing, strict, no_cms, assume_yes, dry_run, cms_yes=cms_yes, auto=auto)
        return rc, collect(option, ids, dry_run=dry_run, no_cms=no_cms)
    finally:
        A.finish()


def auto_plan(ids: list[str], assume_yes: bool, dry_run: bool) -> list[tuple[str, list[str]]] | None:
    """读 CMS 每条的阶段 → 决定下一阶段，按阶段分组；打印计划、确认一次。返回 [(内部阶段键, Rule IDs)]；None = 不执行。"""
    from playwright.sync_api import sync_playwright
    from .active_offline import device_codes_for
    from .cms_grey import read_stages
    batches = depdb.load_all()
    code_of, _no = device_codes_for(ids, batches)
    for rec, _ in resolve_rules(ids, batches)[0]:
        code_of.setdefault(rec["ruleid"], rec.get("方案号", ""))
    print(color(f"update（自动下一阶段）：{len(ids)} 个 Rule ID —— 先在 CMS『在线规则』读当前阶段（只读）", "bold"))
    with sync_playwright() as p:
        cms = read_stages(p, {r: c for r, c in code_of.items() if c})
    from .stage_db import build_state, load_events
    plan, lines, skipped = plan_next(ids, cms, code_of, build_state(load_events()))
    for ln in lines:
        print(ln)
    if not plan:
        print(color("\n没有要放量的 Rule ID。", "yellow"))
        return None
    from .update_report import TRANSITION
    print(color("\n计划（按阶段分组，逐组执行；每组照常：成功率闸门 → CMS 灰度 → DASS）：", "bold"))
    for k, g in plan:
        frm, to = TRANSITION[k]
        tgt = "20K / 完全开放（看源版本用户数）" if k in ("ota20k", "ota20k-again") else to
        print(f"  · {k:13} {frm} → {tgt}：{', '.join(g)}")
    if skipped:
        print(color(f"  不处理：{', '.join(skipped)}", "yellow"))
    if dry_run:
        print(color("（测试模式：不需要确认，下面每组只在 CMS 走一遍流程）", "dim"))
        return plan
    if assume_yes:
        return plan
    if not sys.stdin.isatty():
        print(err("非交互终端无法确认，什么都没做（需要无人值守请加 --yes）。"))
        return None
    try:
        go = input(f"  输入 yes 按上面的计划执行（{sum(len(g) for _, g in plan)} 个 Rule ID，"
                   f"{len(plan)} 组），其它任意输入取消：").strip().lower() == "yes"
    except EOFError:
        go = False
    if not go:
        print("已取消，什么都没改。")
        return None
    return plan


def plan_next(ids, cms: dict, code_of: dict, local: dict):
    """纯逻辑：CMS 阶段 → 下一阶段分组。返回 (plan=[(内部阶段键, ids)], 打印行, 不处理的 ids)。"""
    from .cms_grey import CMS_NEXT
    from .update_report import TRANSITION
    order = ["ota50", "ota200", "ota1k", "ota5k", "ota20k", "ota20k-again"]
    groups: dict[str, list[str]] = {}
    lines = [f"\n  {'Rule ID':8} {'设备类型代码':20} {'CMS 阶段（累计检测数）':22} {'本地记录':8}  本次"]
    skipped = []
    for rid in ids:
        row = cms.get(rid)
        loc = local.get(rid, {}).get("当前阶段", "") or "-"
        dev = code_of.get(rid) or "-"
        if not row:
            why = "查不到设备类型代码" if dev == "-" else "不在线（CMS『在线规则』里搜不到）"
            skipped.append(rid)
            lines.append(f"  {rid:8} {dev:20} {'-':22} {loc:8}  ✗ {why}，不处理")
            continue
        shown = f"{row.get('阶段名称') or '?'}（{row.get('累计检测数') or '?'}）"
        stage = row.get("stage", "")
        if stage not in CMS_NEXT:
            skipped.append(rid)
            lines.append(f"  {rid:8} {dev:20} {shown:22} {loc:8}  ✗ 看不出是哪个阶段，不处理")
            continue
        loc_std, nxt = CMS_NEXT[stage]
        mark = "" if loc in (loc_std, "-") else f"   ⚠ 本地记录是 {loc}，以 CMS 为准"
        if nxt is None:
            skipped.append(rid)
            lines.append(f"  {rid:8} {dev:20} {shown:22} {loc:8}  已完全开放，跳过{mark}")
            continue
        groups.setdefault(nxt, []).append(rid)
        frm, to = TRANSITION[nxt]
        tgt = "20K / 完全开放" if nxt in ("ota20k", "ota20k-again") else to
        lines.append(f"  {rid:8} {dev:20} {shown:22} {loc:8}  → {tgt}{mark}")
    return [(k, groups[k]) for k in order if k in groups], lines, skipped


def _run(option, ids, skip_missing, strict, no_cms, assume_yes, dry_run, cms_yes=None, auto=False) -> int:
    from . import action_summary as A
    batches = depdb.load_all()
    resolved, problems = resolve_rules(ids, batches)
    print(color(f"update {option}：查询 {len(ids)} 个 Rule ID，解析到 {len(resolved)} 个。", "bold"))
    if problems:
        print(warn(f"有 {len(problems)} 个 Rule ID 无法解析："))
        for rid, why in problems:
            print(f"  · {rid}：{why}")
        if not skip_missing:
            print(err("已中止，什么都没改。修正后重跑，或加 --skip-missing 只处理能解析的。"), file=sys.stderr)
            A.step("从部署库解析 Rule ID", A.FAIL,
                   f"{len(problems)} 个解析不出（{', '.join(r for r, _ in problems[:8])}），已中止，什么都没改")
            return 2
        print(warn("--skip-missing：跳过这些，继续处理其余的。"))
    if not resolved:
        print(err("没有可更新的 Rule ID。"), file=sys.stderr)
        A.step("从部署库解析 Rule ID", A.FAIL, "没有可更新的 Rule ID")
        return 2
    A.step("从部署库解析 Rule ID", A.WARN if problems else A.OK,
           f"{len(resolved)} 个取到方案号/目标版本/设备扩展信息"
           + (f"；跳过解析不出的 {len(problems)} 个（--skip-missing）" if problems else ""))

    records = [rec for rec, _ in resolved]
    if option == "ota20k" and not auto:
        # 放 20K：整批都在 20K → 再放一轮 20K；都在 5K（或没有记录）→ 第一次。混着就停。
        from .big_stage import start_option
        from .stage_db import build_state, load_events
        state = build_state(load_events())
        cur = {r["ruleid"]: state.get(r["ruleid"], {}).get("当前阶段", "") for r in records}
        opt2, why = start_option(cur)
        if opt2 is None:
            print(err(f"无法判断这批在 5K 还是 20K：{why}。什么都没改。"), file=sys.stderr)
            A.step("判断当前阶段", A.FAIL, why + "，已中止，什么都没改")
            return 2
        if opt2 == "ota20k-again":
            print(color("这批当前已在 20K：本次是『再放一轮 20K』（之后按源版本用户数决定再一轮 20K 还是完全开放）。", "bold"))
        option = opt2
    _print_plan("ota20k" if option == "ota20k-again" else option, resolved, dry_run)
    from .dass_fill import execute
    if dry_run:
        # 测试模式：只在 CMS 走一遍灰度放开流程（找规则、核对检测数、打开配置窗口后取消），
        # 不做成功率闸门、不执行批量操作、不更新 DASS。
        from playwright.sync_api import sync_playwright
        from .active_offline import device_codes_for
        from .cms_grey import run_grey_release
        code_of, _ = device_codes_for([r["ruleid"] for r in records], batches)
        for r in records:
            code_of.setdefault(r["ruleid"], r.get("方案号", ""))
        with sync_playwright() as p:
            run_grey_release(p, option, code_of, dry_run=True)
        print(color("（测试模式结束：成功率闸门和 DASS 更新都没有执行）", "dim"))
        A.step("成功率闸门", A.SKIP, "测试模式不检查")
        A.step(f"DASS 更新", A.SKIP, "测试模式不更新 DASS")
        return 0
    before = None
    if not no_cms:
        # CMS 灰度放开：成功率闸门通过之后、更新 DASS 之前执行。设备类型代码取自部署库（与 offline 相同）。
        from .active_offline import device_codes_for
        from .cms_grey import run_grey_release
        code_of, _no_dev = device_codes_for([r["ruleid"] for r in records], batches)
        for r in records:  # 明细里的方案号就是设备类型代码，部署库查不到时用它兜底
            code_of.setdefault(r["ruleid"], r.get("方案号", ""))
        _cy = assume_yes if cms_yes is None else cms_yes   # 自动模式：开始前已确认过，CMS 执行前不再问
        before = lambda p, opt: run_grey_release(p, opt, code_of, assume_yes=_cy)
        print(color("（流程：成功率闸门 → CMS 在线规则灰度放开 → DASS 更新阶段；加 --no-cms 只更新 DASS）", "dim"))
    else:
        A.step("CMS 灰度放开", A.SKIP, "--no-cms：不动 CMS")
    return execute(option, records, None, allow_gate_override=not strict, before_update=before,
                   assume_yes=assume_yes)


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
    # —— 不写阶段：按 CMS 阶段自动进到下一阶段 ——
    cms = {"1": {"阶段名称": "测试阶段", "累计检测数": "5", "stage": "测试阶段"},
           "2": {"阶段名称": "5000阶段", "累计检测数": "6255", "stage": "5000阶段"},
           "3": {"阶段名称": "20000阶段", "累计检测数": "26255", "stage": "20000阶段"},
           "4": {"阶段名称": "全量", "累计检测数": "-1", "stage": "全量"},
           "5": {"阶段名称": "5000阶段", "累计检测数": "6255", "stage": "5000阶段"},
           "7": {"阶段名称": "奇怪", "累计检测数": "9", "stage": ""}}
    plan, lines, skipped = plan_next(["1", "2", "3", "4", "5", "6", "7"], cms, {r: "DEV" for r in "1234567"},
                                     {"2": {"当前阶段": "1K"}, "5": {"当前阶段": "5K"}})
    check("自动下一阶段：按 CMS 阶段分组、按阶段顺序",
          plan == [("ota50", ["1"]), ("ota20k", ["2", "5"]), ("ota20k-again", ["3"])])
    check("已完全开放 / 不在线 / 看不出阶段 → 不处理", skipped == ["4", "6", "7"])
    check("本地记录和 CMS 不一致 → 提醒（以 CMS 为准）", any("本地记录是 1K，以 CMS 为准" in l for l in lines)
          and not any("本地记录是 5K" in l for l in lines))
    print("\n自测结果：", "全部通过 ✅" if good else "存在失败 ❌")
    return 0 if good else 1


if __name__ == "__main__":
    sys.exit(main())
