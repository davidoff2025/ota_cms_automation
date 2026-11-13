"""DASS 升级成功率『独立查询 CLI』：命令行传入 方案号 + 目标版本 + 一组 Rule ID，
复用 dass_perf 的绩效页查询/解析逻辑，做【只读】的 Stage1 模糊查询，对每个目标 Rule ID
打印成功率、是否 >= 50%、来源阶段，最后给出汇总。

用法：
    cd ota_cms_automation
    ./.venv/bin/python -m automation.dass_perf_query \
        --plancode MT9603AndrGTV4K --target 0100AQ0725 65142-65154

Rule ID 位置参数支持区间与逗号混合：
    65142-65154   展开为 65142..65154 连续 13 个
    65142,65150   逗号列举两个
    65,142        千分位逗号：归一化为单个 65142（见下方『逗号语义』）
可给多个位置参数，最终合并、去重保序。

逗号语义（与页面侧 norm_ruleid 去分隔符方向一致）：
  DASS 页面把 ruleid 当数字展示、带千分位逗号（如 '65,142' 本意是单个 65142）。
  本 CLI 的逗号解析据此设计——只有当某一整段恰好是『合法千分位分组的单个数字』
  （首组 1~3 位、其后每组恰 3 位，如 65,142 / 1,234,567）时，逗号才被当作千分位
  分隔符、整段归一为一个 Rule ID；其余情形（如 65142,65150）逗号才作为列表分隔符。
  这样从页面直接复制粘贴 '65,142' 得到的是单个 65142，而不会被静默拆成 65 与 142。

本 CLI 只做 Stage1（方案号+目标版本 模糊查询，没有设备扩展信息，不跑 Stage2）；
某 Rule ID Stage1 查不到就标『未找到（Stage1）』继续。全程只读：不点升级/提交/插入行，
不注册任何 prompt 应答。关键节点用 _shot 截图。

返回码：
    0  全部找到且全部通过（成功率 >= 50%）
    1  有不通过 / 未找到 / 成功率解析失败
    2  参数非法 / 绩效页列缺失（安全失败）
"""
from __future__ import annotations

import argparse
import re
import sys

from playwright.sync_api import sync_playwright

from . import config
from ._log import color
from .dass_login import login, _shot
from .dass_fill import norm_ruleid, _report_frame
from .dass_perf import (
    navigate_to_perf,
    query_rates_stage1,
    PASS_THRESHOLD,
)

# 区间段：两端为纯数字、中间一个连字符（允许两侧空白）。
_RANGE_RE = re.compile(r"^\s*(\d+)\s*-\s*(\d+)\s*$")
# 区间连接符：半角 - ~，全角 ～ －，波浪 〜，en/em dash – —。统一归一成 '-'。
_RANGE_SEPS = "~～〜－–—"


def _normalize_spec(text: str) -> str:
    """统一区间连接符为 '-'，去掉连接符两侧的空格（'65229 ~ 65231' → '65229-65231'）。"""
    t = str(text or "")
    for ch in _RANGE_SEPS:
        t = t.replace(ch, "-")
    return re.sub(r"\s*-\s*", "-", t).strip()


def parse_ruleid_args(args: list[str]) -> list[str]:
    """把命令行上的 Rule ID 参数（可能被 shell 按空格拆成多段）解析成去重保序的 Rule ID 列表。

    支持：65148-65154 / 65229~65231 / 65229 ~ 65231 / 65229～65231 / 逗号 / 空格分隔，可混用。
    · shell 会把单独的 ~ 展开成家目录（'65229 ~ 65231' 到程序里是 65229、/Users/xxx、65231），这里还原成 ~。
    · 参数先用空格拼回一行，再统一连接符，所以 '65148 - 65154' 这类被拆开的写法也能还原。
    · 归一后剩下的空格当作列表分隔（同逗号）。
    """
    import os
    home = os.path.expanduser("~")
    toks = ["~" if a in (home, home + "/") else a for a in args]
    line = _normalize_spec(" ".join(toks))
    line = re.sub(r"\s+", ",", line)
    return parse_ruleid_spec(line)

# 千分位分组的单个数字：首组 1~3 位，其后每组恰为逗号(半/全角)+3 位，至少含一个逗号。
# 命中即说明整段是一个带千分位逗号的单个数字（如 65,142 / 1,234,567），逗号为千分位
# 分隔符而非列表分隔符——与页面侧 norm_ruleid 去分隔符方向一致（见模块 docstring）。
_THOUSANDS_RE = re.compile(r"^\s*\d{1,3}(?:[,，]\d{3})+\s*$")


def parse_ruleid_spec(spec: str) -> list[str]:
    """把 Rule ID 规格串解析成 norm_ruleid 形态（纯数字串）的目标列表，去重保序。

    规则：
      - 若整段恰好是『合法千分位分组的单个数字』（首组 1~3 位、其后每组恰 3 位，
        如 65,142 / 1,234,567），逗号视为千分位分隔符，整段归一为单个 Rule ID
        （norm_ruleid 去逗号）。这样从页面复制粘贴 '65,142' 得到的是单个 65142，
        而不会被拆成 65 与 142（与页面结果侧 norm_ruleid 的去分隔符方向一致）。
      - 否则按逗号（半角 ',' 与全角 '，'）分段作列表分隔，再逐段处理：
        - 段内匹配 '^数字-数字$'（连字符区间，两端纯数字）-> 闭区间 a..b，要求 a <= b，
          展开为 a, a+1, ..., b；a > b 视为非法输入（raise ValueError）。
        - 否则对该段取 norm_ruleid（去一切非数字）；归一化后为空则跳过并告警。
      - 结果去重但保留首次出现顺序。
    空规格（无任何有效段）-> raise ValueError。
    """
    spec = _normalize_spec(spec)
    out: list[str] = []
    seen: set[str] = set()

    def _add(rid: str) -> None:
        if rid and rid not in seen:
            seen.add(rid)
            out.append(rid)

    # 整段是带千分位逗号的单个数字 -> 当作一个 Rule ID（逗号是千分位分隔符）。
    if _THOUSANDS_RE.match(spec):
        _add(norm_ruleid(spec))
        if not out:
            raise ValueError("未解析到任何有效的 Rule ID。")
        return out

    segments = re.split(r"[,，]", spec)
    for seg in segments:
        if not seg.strip():
            continue
        m = _RANGE_RE.match(seg)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if a > b:
                raise ValueError(f"非法区间 {seg.strip()!r}：起始值大于结束值（{a} > {b}）。")
            for v in range(a, b + 1):
                _add(str(v))
            continue
        if not re.fullmatch(r"\s*\d+\s*", seg):
            # 以前会把非数字字符直接删掉再拼起来（'65229~65231' → '6522965231'），这里改成报错，不猜。
            raise ValueError(f"看不懂的 Rule ID 片段：{seg.strip()!r}（写法：65370 / 65370-65379 / 65370~65379 / 逗号分隔）")
        _add(norm_ruleid(seg))

    if not out:
        raise ValueError("未解析到任何有效的 Rule ID。")
    return out


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="python -m automation.dass_perf_query",
        description="DASS 升级成功率独立查询 CLI（只读 Stage1 模糊查询）。",
    )
    p.add_argument("--plancode", required=True, help="方案号（填 PLANCODE 过滤框）。")
    p.add_argument("--target", required=True, help="目标版本（填 NEXTVERSION 过滤框）。")
    p.add_argument(
        "ruleids",
        nargs="*",
        help="一组 Rule ID，支持区间与逗号混合，如 65142-65154 或 65142,65150；"
             "与 --all 互斥，给了 --all 时留空。",
    )
    p.add_argument(
        "--all",
        action="store_true",
        dest="list_all",
        help="不限定目标 Rule ID，列出该 方案号+目标版本 下 Stage1 查到的全部行"
             "（与位置参数 ruleids 互斥）。",
    )
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="只解析 Rule ID 列表并打印，不启动浏览器（供校验用）。",
    )
    return p


def _validate_mode(list_all: bool, ruleids: list[str]) -> str:
    """校验 --all 与位置参数 ruleids 的互斥/必填关系，返回运行模式。

    - list_all 且 ruleids 非空 -> 抛 ValueError（互斥，二选一）。
    - 两者皆空 -> 抛 ValueError（必填其一）。
    - 仅 list_all -> 返回 "all"。
    - 仅 ruleids -> 返回 "targets"。
    """
    has_ruleids = bool(ruleids)
    if list_all and has_ruleids:
        raise ValueError("--all 与位置参数 ruleids 互斥，二选一。")
    if not list_all and not has_ruleids:
        raise ValueError("请给出一组 Rule ID，或用 --all 列出全部。")
    return "all" if list_all else "targets"


def _collect_targets(ruleid_args: list[str]) -> list[str]:
    """把多个位置参数合并解析为去重保序的目标 Rule ID 列表。"""
    seen: set[str] = set()
    targets: list[str] = []
    for rid in parse_ruleid_args(ruleid_args):
        if rid not in seen:
            seen.add(rid)
            targets.append(rid)
    return targets


def _print_report(targets: list[str], rates: dict, pages: int = 1) -> tuple[list[str], list[str]]:
    """按目标顺序逐条打印，返回 (不通过列表, 未找到列表)。

    rates 为 query_rates_stage1 的返回 dict（键为 norm_ruleid）。pages 为跨页遍历
    实际翻过的页数，用于『未找到（Stage1，已遍历全部 N 页）』提示。
    判定：命中且 rate >= PASS_THRESHOLD -> 通过；命中但 rate < 阈值 -> 不通过；
    命中但 rate 为 None（解析失败）-> 不通过（标成功率解析失败）；未命中 -> 未找到。
    """
    failed: list[str] = []
    not_found: list[str] = []
    print("\n======== 查询结果（Stage1 模糊查询）========")
    for rid in targets:
        info = rates.get(rid)
        if info is None:
            print(color(f"  Rule ID={rid}  成功率=未找到（Stage1，已遍历全部 {pages} 页）  未找到", "yellow"))
            not_found.append(rid)
            continue
        rate = info.get("rate")
        if rate is None:
            print(color(
                f"  Rule ID={rid}  成功率解析失败（原文 {info.get('rate_text', '')!r}）  不通过 ❌  （Stage1）",
                "red"))
            failed.append(rid)
            continue
        passed = rate >= PASS_THRESHOLD
        mark = color("通过 ✅", "green") if passed else color("不通过 ❌", "red")
        print(f"  Rule ID={rid}  成功率={rate:.2f}%  {mark}  （Stage1）")
        if not passed:
            failed.append(rid)
    return failed, not_found


def _print_summary(targets: list[str], failed: list[str], not_found: list[str], pages: int = 1) -> None:
    print("\n======== 汇总 ========")
    total = len(targets)
    passed_n = total - len(failed) - len(not_found)
    print(f"  目标 Rule ID 总数：{total}")
    print(f"  通过：{passed_n}")
    if not failed and not_found == []:
        print(color("  全部通过 ✅", "green"))
        return
    print(color("  以下不通过/未找到：", "red"))
    if failed:
        print(color(f"    不通过（< {PASS_THRESHOLD:.0f}% 或解析失败）：{failed}", "red"))
    if not_found:
        print(color(f"    未找到（Stage1，已遍历全部 {pages} 页）：{not_found}", "yellow"))


def _ruleid_sort_key(rid: str):
    """norm_ruleid 排序键：纯数字按数值升序，非纯数字退化为字符串比较（排在数字之后）。"""
    return (0, int(rid)) if rid.isdigit() else (1, rid)


def _print_all(rates: dict, pages: int = 1) -> None:
    """--all 模式：把 Stage1 查到的【每一行】都打印出来（不做任何过滤、不触发 abort）。

    rates 为 query_rates_stage1 的返回 dict（键为 norm_ruleid）。按 norm_ruleid 升序
    （数字优先）逐行打印 raw_ruleid / norm_ruleid / 成功率文本 / 解析 float / 是否 >=50%，
    末尾打印汇总：总行数、Rule ID 范围（最小 ~ 最大）。
    """
    print(f"\n======== 全部 Rule ID 及成功率（Stage1 模糊查询，已遍历全部 {pages} 页）========")
    keys = sorted(rates, key=_ruleid_sort_key)
    if not keys:
        print(f"  该 方案号+目标版本 下 Stage1 未查到任何行（0 行）。")
        print("\n======== 汇总 ========")
        print("  总行数：0")
        print("  Rule ID 范围：（无）")
        return
    for rid in keys:
        info = rates[rid]
        rate = info.get("rate")
        raw = info.get("raw_ruleid", "")
        rate_text = info.get("rate_text", "")
        if rate is None:
            parsed = "解析失败"
            ge = "否"
        else:
            parsed = f"{rate:.2f}"
            ge = "是" if rate >= PASS_THRESHOLD else "否"
        print(
            f"  Rule ID 原始显示值={raw!r} 归一值={rid}  "
            f"成功率={rate_text!r}（解析={parsed}）  >=50%? {ge}"
        )
    print("\n======== 汇总 ========")
    print(f"  总行数：{len(keys)}")
    print(f"  Rule ID 范围：{keys[0]} ~ {keys[-1]}")


def main() -> int:
    args = _build_parser().parse_args()

    try:
        mode = _validate_mode(args.list_all, args.ruleids)
    except ValueError as e:
        print(color(f"✗ 参数错误：{e}", "red"))
        return 2

    if mode == "all":
        print(f"方案号：{args.plancode}")
        print(f"目标版本：{args.target}")
        print("模式：--all（列出该 方案号+目标版本 下 Stage1 查到的全部行）")
        if args.dry_run:
            print("（--dry-run：仅解析参数，不启动浏览器。）")
            return 0
        return _run_all(args)

    try:
        targets = _collect_targets(args.ruleids)
    except ValueError as e:
        print(color(f"✗ Rule ID 解析失败：{e}", "red"))
        return 2

    print(f"方案号：{args.plancode}")
    print(f"目标版本：{args.target}")
    print(f"目标 Rule ID（{len(targets)} 个）：{targets}")

    if args.dry_run:
        print("（--dry-run：仅解析参数，不启动浏览器。）")
        return 0

    return _run_targets(args, targets)


def _run_targets(args, targets: list[str]) -> int:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not config.HEADED)
        context = browser.new_context()
        page = context.new_page()
        # 只读查询：不注册任何 dialog 应答。
        try:
            login(page)
            print("\n======== 导航『安卓升级展示报表』========")
            navigate_to_perf(page)
            page.wait_for_timeout(2500)
            fr = _report_frame(page, "perf")
            if fr is None:
                print(color("✗ 未找到绩效页 frame，无法查询。", "red"))
                _shot(page, "perf_query_no_frame")
                return 2

            # 传入目标 Rule ID 集合作提前停止：跨页遍历（_perf_query 内），目标全部命中即停。
            # stats_out 收集实际遍历页数，用于『未找到（Stage1，已遍历全部 N 页）』提示。
            stats: dict = {}
            rates = query_rates_stage1(
                fr, page, args.plancode, args.target, set(targets), stats
            )
            _shot(page, "perf_query_done")
            if rates is None:
                print(color("✗ 绩效页未解析到成功率/ruleid 列字母，安全失败。", "red"))
                _shot(page, "perf_query_missing_columns")
                return 2

            pages = stats.get("pages", 1)
            failed, not_found = _print_report(targets, rates, pages)
            _print_summary(targets, failed, not_found, pages)
            return 0 if (not failed and not not_found) else 1
        except Exception as e:
            print(color(f"✗ 出错：{e}", "red"))
            _shot(page, "perf_query_error")
            return 1
        finally:
            context.close()
            browser.close()


def _run_all(args) -> int:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not config.HEADED)
        context = browser.new_context()
        page = context.new_page()
        # 只读查询：不注册任何 dialog 应答，不点升级/提交/插入行。
        try:
            login(page)
            print("\n======== 导航『安卓升级展示报表』========")
            navigate_to_perf(page)
            page.wait_for_timeout(2500)
            fr = _report_frame(page, "perf")
            if fr is None:
                print(color("✗ 未找到绩效页 frame，无法查询。", "red"))
                _shot(page, "perf_query_no_frame")
                return 2

            # --all：不限定目标 Rule ID（target_ruleids=None），跨页遍历查询全集。
            # stats_out 收集实际遍历页数，供汇总展示。
            stats: dict = {}
            rates = query_rates_stage1(
                fr, page, args.plancode, args.target, None, stats
            )
            _shot(page, "perf_query_done")
            if rates is None:
                print(color("✗ 绩效页未解析到成功率/ruleid 列字母，安全失败。", "red"))
                _shot(page, "perf_query_missing_columns")
                return 2

            _print_all(rates, stats.get("pages", 1))
            return 0
        except Exception as e:
            print(color(f"✗ 出错：{e}", "red"))
            _shot(page, "perf_query_error")
            return 1
        finally:
            context.close()
            browser.close()


# —— 不依赖浏览器的纯函数自测（python -m automation.dass_perf_query --selftest）——
def _selftest() -> int:
    ok = True

    def check(name, cond):
        nonlocal ok
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        if not cond:
            ok = False

    # 区间展开：65142-65154 -> 13 个连续
    import os as _os
    want = [str(i) for i in range(65148, 65155)] + ["65229", "65230", "65231"]
    home = _os.path.expanduser("~")
    for label, args in (("65148-65154 65229~65231", ["65148-65154", "65229~65231"]),
                        ("全角 ～", ["65148-65154", "65229～65231"]),
                        ("空格 + shell 展开的 ~", ["65148-65154", "65229", home, "65231"]),
                        ("被拆开的 - ", ["65148", "-", "65154", "65229", "~", "65231"]),
                        ("引号里带空格", ["65148 - 65154", "65229 ~ 65231"]),
                        ("逗号 + 空格", ["65148-65154, 65229 ~65231"]),
                        ("en dash", ["65148–65154,65229—65231"])):
        check(f"同等解析：{label}", parse_ruleid_args(args) == want)
    bad = False
    try:
        parse_ruleid_spec("65229x65231")
    except ValueError:
        bad = True
    check("无法识别的片段报错（不再拼成一个大数）", bad)
    r = parse_ruleid_spec("65142-65154")
    check("65142-65154 展开为 13 个", len(r) == 13)
    check("区间首项=65142", r[0] == "65142")
    check("区间末项=65154", r[-1] == "65154")
    check("区间连续无缺项", r == [str(x) for x in range(65142, 65155)])

    # 逗号列举
    check("65142,65150 -> 两个", parse_ruleid_spec("65142,65150") == ["65142", "65150"])
    # 全角逗号（列表分隔：两侧均非千分位分组，故按逗号列举）
    check("全角逗号 65142，65150 -> 两个", parse_ruleid_spec("65142，65150") == ["65142", "65150"])
    # 千分位逗号：'65,142' 本意是单个 65142（页面把 ruleid 当数字展示、带千分位逗号），
    # 与页面结果侧 norm_ruleid 去分隔符方向一致，不被拆成 65 与 142。
    check("65,142 -> ['65142']（千分位逗号归一为单个）", parse_ruleid_spec("65,142") == ["65142"])
    check("全角千分位 65，142 -> ['65142']", parse_ruleid_spec("65，142") == ["65142"])
    check("多组千分位 1,234,567 -> ['1234567']", parse_ruleid_spec("1,234,567") == ["1234567"])
    # 列表分隔不被误判为千分位：groups 不是『1~3 位 + 每组恰 3 位』。
    check("65142,65150 -> 两个（非千分位分组，列表分隔）", parse_ruleid_spec("65142,65150") == ["65142", "65150"])
    # 单点区间
    check("65142-65142 -> ['65142']", parse_ruleid_spec("65142-65142") == ["65142"])
    # 混合
    check(
        "65142-65144,65150 混合",
        parse_ruleid_spec("65142-65144,65150") == ["65142", "65143", "65144", "65150"],
    )
    # 去重保序
    check(
        "重复去重保序",
        parse_ruleid_spec("65142,65142,65143") == ["65142", "65143"],
    )

    # 非法区间
    try:
        parse_ruleid_spec("65154-65142")
        check("非法区间 65154-65142 应抛错", False)
    except ValueError:
        check("非法区间 65154-65142 抛 ValueError", True)

    # 空规格
    try:
        parse_ruleid_spec("")
        check("空规格应抛错", False)
    except ValueError:
        check("空规格抛 ValueError", True)

    # 多参数合并去重保序
    check(
        "_collect_targets 合并去重保序",
        _collect_targets(["65142-65144", "65143,65150"]) == ["65142", "65143", "65144", "65150"],
    )

    # —— --all 与位置参数互斥/必填的模式校验 ——
    check("--all 且无 ruleids -> 'all'", _validate_mode(True, []) == "all")
    check("无 --all 且有 ruleids -> 'targets'", _validate_mode(False, ["65142"]) == "targets")
    try:
        _validate_mode(True, ["65142"])
        check("--all 与 ruleids 并存应抛错（互斥）", False)
    except ValueError:
        check("--all 与 ruleids 并存抛 ValueError（互斥）", True)
    try:
        _validate_mode(False, [])
        check("两者皆空应抛错（必填）", False)
    except ValueError:
        check("两者皆空抛 ValueError（必填）", True)

    # —— _build_parser：ruleids 变为可选、--all 可解析 ——
    parser = _build_parser()
    ns_all = parser.parse_args(["--plancode", "P", "--target", "T", "--all"])
    check("解析 --all 时 list_all=True 且 ruleids 为空", ns_all.list_all and ns_all.ruleids == [])
    ns_tgt = parser.parse_args(["--plancode", "P", "--target", "T", "65142"])
    check("解析位置参数时 list_all=False 且 ruleids=['65142']",
          (not ns_tgt.list_all) and ns_tgt.ruleids == ["65142"])

    # —— _print_all 冒烟 + 汇总范围（捕获 stdout 校验总行数/范围）——
    import io
    from contextlib import redirect_stdout

    sample = {
        "64807": {"rate": 100.0, "rate_text": "100.00%", "raw_ruleid": "64,807"},
        "64813": {"rate": 0.0, "rate_text": "0.00%", "raw_ruleid": "64813"},
        "64923": {"rate": None, "rate_text": "--", "raw_ruleid": "64923"},
    }
    buf = io.StringIO()
    with redirect_stdout(buf):
        _print_all(sample, 3)
    out = buf.getvalue()
    check("_print_all 汇总总行数=3", "总行数：3" in out)
    check("_print_all 汇总范围=64807 ~ 64923", "Rule ID 范围：64807 ~ 64923" in out)
    check("_print_all 对解析失败行标注『解析失败』", "解析=解析失败" in out)
    check("_print_all 对 >=50% 行标『是』", ">=50%? 是" in out)
    check("_print_all 对 <50% 行标『否』", ">=50%? 否" in out)

    # 空 dict 不崩，且汇总显示 0 行。
    buf2 = io.StringIO()
    with redirect_stdout(buf2):
        _print_all({}, 1)
    out2 = buf2.getvalue()
    check("_print_all 空结果不崩且总行数=0", "总行数：0" in out2)

    print("\n自测结果：", "全部通过 ✅" if ok else "存在失败 ❌")
    return 0 if ok else 1


if __name__ == "__main__":
    if "--selftest" in sys.argv[1:]:
        sys.exit(_selftest())
    sys.exit(main())
