"""deployota offline —— 在 CMS「规则上下线编排」里把指定 Rule ID 下线（online 的反向操作）。

用法：
    ./deployota offline 50000                 # 单个
    ./deployota offline 50000-50005           # 区间
    ./deployota offline 50000,50003,50010     # 逗号分隔（可与区间混用、可给多个参数）
    ./deployota offline 50000-50005 --yes     # 跳过开始前的确认
    ./deployota offline 50000 --device MT9655AMAndrGTV   # 部署库里查不到设备类型代码时手工指定

流程（不做单独的状态检查，逐个直接处理）：
  1) 从部署库按 Rule ID 取设备类型代码（--device 可覆盖）；取不到的不处理。
  2) 开始前确认一次（--yes 跳过）。
  3) 打开「升级事件 → 规则上下线编排」，按设备类型分组，在『已选规则』区设一次设备类型。
  4) 每个 Rule ID：在『已选规则』（在线区）按规则ID精确搜索——
       搜到且有『下 线』按钮 → 点『下 线』→ 确认框『确认下线该规则?』里点『确定』
         （点之前确认背后的搜索结果仍是目标行；绝不点编排弹窗底部的『确 定』）→
         重新搜索，『已选规则』里没有了才算成功；
       搜不到 → 已经下线或不存在，直接跳过。
     每次搜索都核对输入框里确实是目标号（防止填进其它框）。

结果：下线成功 / 跳过（已下线或找不到）/ 查不到设备类型代码 / 失败（附原因），
再为下线成功的规则打印『下线信息汇总』表。
退出码：0 没有失败；1 有失败或有查不到设备类型代码的；2 参数问题或你取消。
"""
from __future__ import annotations

import sys
from datetime import datetime

from playwright.sync_api import sync_playwright

from . import config, depdb
from ._log import color, err, ok, warn
from .active_ota import (
    _click_button_in,
    _detect_ui_error,
    _find_pending_row_by_ruleid,
    _search_selected_zone,
    _selected_table,
    _set_device_type_once,
    open_orchestration_modal,
)
from .dass_fill import norm_ruleid
from .login import _shot, login

# 每次搜索后等目标行出现的最长时间（毫秒）。
SEARCH_WAIT_MS = 6000



# ======================================================================
# 纯逻辑（--selftest 覆盖）
# ======================================================================

def device_codes_for(rule_ids: list[str], batches: list[dict]) -> tuple[dict, list[str]]:
    """从部署库批次里给每个 Rule ID 找设备类型代码。

    返回 ({rule_id: 设备类型代码}, [查不到的 rule_id])。
    同一 Rule ID 出现在多个批次时：优先『有效』批次，其次部署时间最新。
    """
    found: dict[str, str] = {}
    missing: list[str] = []
    for rid in rule_ids:
        owners = [b for b in batches
                  if rid in [x.strip() for x in str(b.get("Rule ID列表", "")).split(",")]
                  and str(b.get("设备类型代码", "")).strip()]
        if not owners:
            missing.append(rid)
            continue
        owners.sort(key=lambda b: (str(b.get("状态", depdb.STATUS_ACTIVE)) == depdb.STATUS_ACTIVE,
                                   str(b.get("部署时间", ""))))
        found[rid] = str(owners[-1]["设备类型代码"]).strip()
    return found, missing


def group_by_device(code_of: dict) -> dict[str, list[str]]:
    """{rule_id: device} → {device: [rule_id 升序]}（按设备类型代码排序）。"""
    out: dict[str, list[str]] = {}
    for rid, dev in code_of.items():
        out.setdefault(dev, []).append(rid)
    return {d: sorted(v, key=lambda x: int(x) if x.isdigit() else 0) for d, v in sorted(out.items())}


def build_offline_rows(rids: list[str], details: dict, cms_cells: dict, times: dict) -> list[dict]:
    """组装下线汇总表的行（纯逻辑）。

    details：{rid: 部署库/Excel 解析出的记录（dass_update.resolve_rules 的格式：
              升级类型/机型/设备扩展信息/源版本/目标版本=目标完整版本）}；
    cms_cells：{rid: 下线前在 CMS『已选规则』里读到的该行 {列名: 文本}}，details 缺失时兜底；
    times：{rid: 下线时间}。
    """
    rows = []
    for rid in sorted(rids, key=lambda x: int(x) if x.isdigit() else 0):
        d = details.get(rid) or {}
        c = cms_cells.get(rid) or {}
        rows.append({
            "Rule ID": rid,
            "升级类型": d.get("升级类型") or c.get("升级方式", "") or "-",
            "机型信息": d.get("机型") or "-",
            "机型扩展码": d.get("设备扩展信息") or c.get("设备扩展信息", "") or "-",
            "源版本": d.get("源版本") or c.get("源版本", "") or "-",
            "目标版本": d.get("目标版本") or c.get("目标版本", "") or "-",
            "下线时间": times.get(rid, ""),
        })
    return rows


def offline_time_label(times: list[str]) -> str:
    """汇总头部的下线时间：同一分钟内显示一个时间，否则显示『开始 ~ 结束』。"""
    ts = sorted(t for t in times if t)
    if not ts:
        return ""
    return ts[0] if ts[0][:16] == ts[-1][:16] else f"{ts[0]} ~ {ts[-1]}"


def format_offline_summary(success, skipped, no_device, failed) -> str:
    """最终汇总（纯逻辑）。failed = [(rid, 原因), ...]。"""
    def ids(xs):
        return ", ".join(xs) if xs else "-"
    lines = [color("======== offline 下线汇总 ========", "bold"),
             ok(f"下线成功（{len(success)}）：{ids(success)}"),
             f"  跳过：已下线或找不到（{len(skipped)}）：{ids(skipped)}"]
    if no_device:
        lines.append(warn(f"部署库查不到设备类型代码、未处理（{len(no_device)}）：{ids(no_device)}"
                          "（可用 --device 指定）"))
    if failed:
        lines.append(err(f"失败（{len(failed)}）："))
        for rid, why in failed:
            lines.append(color(f"    {rid}：{why}", "red"))
    total = len(success) + len(skipped) + len(no_device) + len(failed)
    lines.append(f"  —— 总数：{total}  下线成功：{len(success)}  跳过：{len(skipped)}  "
                 f"未处理：{len(no_device)}  失败：{len(failed)}")
    return "\n".join(lines)


# ======================================================================
# 页面操作
# ======================================================================

def _wait_row(page, table, rule_id: str, timeout_ms: int = SEARCH_WAIT_MS) -> int:
    """等表里出现规则ID列精确等于 rule_id 的唯一一行，返回行索引；超时返回 -1。"""
    waited = 0
    while waited <= timeout_ms:
        idx = _find_pending_row_by_ruleid(table, rule_id)
        if idx >= 0:
            return idx
        page.wait_for_timeout(250)
        waited += 250
    return -1


def _row_has_button(row, keyword: str) -> bool:
    """行内是否有可见的、去空白后包含 keyword 的按钮（兼容 `下 线` / `上 线` 中间的空格）。"""
    btns = row.locator("button")
    for i in range(btns.count()):
        try:
            b = btns.nth(i)
            if b.is_visible() and keyword in "".join((b.inner_text() or "").split()):
                return True
        except Exception:
            continue
    return False


def _row_cells(table, idx: int) -> dict:
    """读表第 idx 行：{表头文字(去空白): 单元格文字}。读失败返回空 dict。"""
    try:
        return table.evaluate(
            """(t, idx) => {
                const heads = [...t.querySelectorAll('thead th, thead td')]
                    .map(h => (h.innerText || '').replace(/\\s/g, ''));
                const row = t.querySelectorAll('tbody tr')[idx];
                if (!row) return {};
                const cells = [...row.querySelectorAll('td, th')];
                const out = {};
                heads.forEach((h, i) => { if (h && cells[i]) out[h] = (cells[i].innerText || '').trim(); });
                return out;
            }""",
            idx,
        ) or {}
    except Exception:
        return {}


def _search_selected(modal, page, rid: str) -> tuple[int, str]:
    """在『已选规则』区搜 rid。返回 (行索引或-1, 问题说明)。问题说明非空表示这次搜索不可信。"""
    want = norm_ruleid(rid)
    if not _search_selected_zone(modal, page, want):
        return -1, "已选规则区的『搜索』按钮没点中"
    box = modal.locator("input#ruleId").first
    try:
        got = (box.input_value() or "").strip()
    except Exception:
        got = ""
    if got != want:
        return -1, f"已选规则区『规则ID』框里是 {got!r}，不是 {want!r}（可能填错了框）"
    return _wait_row(page, _selected_table(modal), want), ""


def _confirm_offline(page, modal, rid: str) -> tuple[bool, str]:
    """在『确认下线该规则?』确认框里点『确定』。点之前确认背后的搜索结果仍是目标行。"""
    waited = 0
    while waited <= 3000 and page.get_by_text("确认下线该规则", exact=False).count() == 0:
        page.wait_for_timeout(250)
        waited += 250
    if page.get_by_text("确认下线该规则", exact=False).count() == 0:
        return False, "点『下 线』后没有出现『确认下线该规则?』确认框"
    # 确认框文案不带 ID：确认背后的已选区搜索结果里仍然是目标行，才点确定。
    if _find_pending_row_by_ruleid(_selected_table(modal), rid) < 0:
        return False, "确认框出现时，背后的搜索结果已经不是目标行，未点确定"
    for sel in (".ant-popover", ".ant-popconfirm", ".ant-modal-confirm"):
        scope = page.locator(sel).last
        try:
            if scope.count() and scope.is_visible() and "确认下线该规则" in (scope.inner_text() or ""):
                if _click_button_in(scope, page, "确定", f"确认下线-确定({rid})"):
                    return True, ""
                return False, "确认框里没找到『确定』按钮"
        except Exception:
            continue
    return False, "没定位到确认框本身的作用域；为避免误点编排弹窗底部的『确 定』，未点击"


def _verify_gone(modal, page, rid: str, timeout_ms: int = SEARCH_WAIT_MS) -> tuple[bool, str]:
    """重新在『已选规则』搜 rid，等它不在表里。返回 (已不在线, 问题说明)。"""
    want = norm_ruleid(rid)
    if not _search_selected_zone(modal, page, want):
        return False, "回读时已选规则区的『搜索』按钮没点中"
    waited = 0
    while waited <= timeout_ms:
        page.wait_for_timeout(250)
        waited += 250
        if _find_pending_row_by_ruleid(_selected_table(modal), want) < 0:
            return True, ""
    return False, "点了确定，但『已选规则』里仍能搜到（未下线）"


def offline_one(modal, page, rid: str) -> tuple[str, str, dict]:
    """处理一个 Rule ID。返回 (结果, 说明, 下线前该行各列)；结果 = 'done' | 'skipped' | 'failed'。"""
    idx, problem = _search_selected(modal, page, rid)
    if problem:
        return "failed", problem, {}
    if idx < 0:
        return "skipped", "已选规则里搜不到（已下线或不存在）", {}
    table = _selected_table(modal)
    cells = _row_cells(table, idx)
    row = table.locator("tbody tr").nth(idx)
    if not _row_has_button(row, "下线"):
        return "failed", "目标行没有『下 线』按钮", cells
    if not _click_button_in(row, page, "下线", f"下线({rid})"):
        return "failed", "点『下 线』失败", cells
    ok_, why = _confirm_offline(page, modal, rid)
    if not ok_:
        _shot(page, f"offline_confirm_fail_{rid}")
        try:  # 关掉可能残留的确认框，不让它挡住后面的操作
            page.keyboard.press("Escape")
        except Exception:
            pass
        return "failed", why, cells
    page.wait_for_timeout(1200)
    ui = _detect_ui_error(page)
    if ui:
        _shot(page, f"offline_ui_error_{rid}")
        return "failed", "CMS 报错：" + " ".join(ui.split()), cells
    gone, why = _verify_gone(modal, page, rid)
    if not gone:
        _shot(page, f"offline_still_online_{rid}")
        return "failed", why, cells
    return "done", "", cells


# ======================================================================
# CLI
# ======================================================================

def _print_offline_tables(success: list[str], code_of: dict, cms_cells: dict, off_time: dict) -> None:
    """为下线成功的规则按设备类型代码打印『下线信息汇总』表。"""
    if not success:
        return
    from .dass_update import resolve_rules
    from .summary import render_summary
    resolved, _problems = resolve_rules(success, depdb.load_all())
    details = {rec["ruleid"]: rec for rec, _src in resolved}
    for dev, rids in group_by_device({r: code_of.get(r, "") for r in success}).items():
        rows = build_offline_rows(rids, details, cms_cells, off_time)
        print("")
        print(render_summary(dev, rows, "offline", offline_time_label([off_time.get(r, "") for r in rids])))
    missing = [r for r in success if r not in details]
    if missing:
        print(color(f"  （{', '.join(missing)} 在部署库/来源 Excel 里没有明细，表中版本等取自下线前 CMS 页面的显示）", "dim"))


def _parse_args(argv: list[str]):
    specs, device, yes = [], "", False
    i = 0
    while i < len(argv):
        a = argv[i]
        if a in ("--yes", "-y"):
            yes = True
        elif a == "--device":
            if i + 1 >= len(argv):
                raise ValueError("--device 后面要跟设备类型代码")
            device = argv[i + 1].strip()
            i += 1
        elif a.startswith("--device="):
            device = a.split("=", 1)[1].strip()
        elif a.startswith("-") and not a[1:2].isdigit():
            raise ValueError(f"未知参数：{a!r}")
        else:
            specs.append(a)
        i += 1
    return specs, device, yes


def main() -> int:
    argv = sys.argv[1:]
    if "--selftest" in argv:
        return _selftest()
    if not argv or any(a in ("-h", "--help") for a in argv):
        print(__doc__)
        return 0 if argv else 2
    try:
        specs, device_override, assume_yes = _parse_args(argv)
    except ValueError as e:
        print(err(str(e)), file=sys.stderr)
        return 2
    if not specs:
        print(err("缺少 Rule ID。例：./deployota offline 50000-50005"), file=sys.stderr)
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

    # 1) 设备类型代码：--device > 部署库 > prep Excel；还没找到的登录后去 CMS『在线规则』页查
    from .device_lookup import from_cms_online_rules, print_sources, resolve_local
    code_of, src, missing = resolve_local(ids, device_override)
    print(color(f"offline：{len(ids)} 个 Rule ID", "bold"))
    print_sources(code_of, src, missing, color, warn)
    no_device: list[str] = []

    # 开始前确认一次（不做单独的状态检查：在线的直接下线，已下线/找不到的跳过）
    todo = list(ids)
    print(color(f"\n将下线 {len(todo)} 个 Rule ID（在线的直接下线，已下线或找不到的跳过）：{', '.join(todo)}", "bold"))
    print("  （下线会立即影响线上设备。）")
    if not assume_yes:
        if not sys.stdin.isatty():
            print(err("非交互终端无法确认，已取消。需要无人值守请加 --yes。"), file=sys.stderr)
            return 2
        try:
            go = input("  输入 yes 开始下线，其它任意输入取消：").strip().lower() == "yes"
        except EOFError:
            go = False
        if not go:
            print("已取消，未下线任何规则。")
            return 2

    success, skipped, failed = [], [], []
    cms_cells: dict[str, dict] = {}   # 下线前在『已选规则』读到的行（汇总兜底用）
    off_time: dict[str, str] = {}     # 每条的下线时间
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not config.HEADED)
        context = browser.new_context()
        page = context.new_page()
        try:
            login(page)
            if missing:
                found = from_cms_online_rules(page, missing)
                code_of.update(found)
                for r in found:
                    print(f"  {r}：CMS『在线规则』查到设备类型代码 {found[r]}")
                gone = [r for r in missing if r not in found]
                if gone:  # 不在『在线规则』里 = 不在线（或不存在），按约定直接跳过
                    print(f"  跳过（『在线规则』里没有，已下线或不存在）：{', '.join(gone)}")
                    skipped += gone
            groups = group_by_device(code_of)
            modal = open_orchestration_modal(page)
            for dev, rids in groups.items():
                if not _set_device_type_once(page, modal, dev):
                    failed += [(r, f"设置设备类型 {dev!r} 失败，未处理") for r in rids]
                    continue
                for rid in rids:
                    print(f"\n---- 下线 Rule ID={rid} ----")
                    try:
                        res, why, cells = offline_one(modal, page, rid)
                    except Exception as e:
                        res, why, cells = "failed", f"异常：{e}", {}
                    if cells:
                        cms_cells[rid] = cells
                    if res == "done":
                        off_time[rid] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        print(ok(f"{rid} 已下线"))
                        success.append(rid)
                    elif res == "skipped":
                        print(f"  {rid} 跳过：{why}")
                        skipped.append(rid)
                    else:
                        print(err(f"{rid} 下线失败：{why}"))
                        failed.append((rid, why))
            _shot(page, "offline_done")
        except Exception as e:
            print(err(f"运行出错：{e}"))
            _shot(page, "offline_error")
            handled = set(success) | set(skipped) | {r for r, _ in failed}
            failed += [(r, f"未处理：{e}") for r in todo if r not in handled]
        finally:
            if config.HEADED:
                print("\n浏览器将在 15 秒后关闭（可自行查看页面）...")
                page.wait_for_timeout(15000)
            context.close()
            browser.close()

    print("\n" + format_offline_summary(success, skipped, no_device, failed))
    _print_offline_tables(success, code_of, cms_cells, off_time)
    return 0 if not (failed or no_device) else 1


def _selftest() -> int:
    good = True

    def check(name, cond):
        nonlocal good
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        good = good and bool(cond)

    batches = [
        {"Rule ID列表": "50000,50001", "设备类型代码": "OLD", "状态": "作废", "部署时间": "20260901"},
        {"Rule ID列表": "50000,50001,50002", "设备类型代码": "DEV_A", "状态": "有效", "部署时间": "20260801"},
        {"Rule ID列表": "50003", "设备类型代码": "DEV_B", "状态": "有效", "部署时间": "20260902"},
    ]
    code_of, missing = device_codes_for(["50000", "50002", "50003", "59999"], batches)
    check("有效批次优先于作废批次（即使作废更新）", code_of["50000"] == "DEV_A")
    check("取到各自的设备类型代码", code_of["50003"] == "DEV_B" and code_of["50002"] == "DEV_A")
    check("库里没有的列为 missing", missing == ["59999"])
    g = group_by_device({"50003": "DEV_B", "50002": "DEV_A", "50000": "DEV_A"})
    check("按设备分组且组内升序", g == {"DEV_A": ["50000", "50002"], "DEV_B": ["50003"]})
    s = format_offline_summary(["50000"], ["50001", "50004"], ["59999"], [("50002", "确认框没出现")])
    check("汇总四类都列出", all(x in s for x in ("下线成功（1）", "跳过：已下线或找不到（2）",
                                               "查不到设备类型代码", "50002：确认框没出现")))
    check("汇总计数", "总数：5" in s and "跳过：2" in s and "未处理：1" in s and "失败：1" in s)
    try:
        _parse_args(["50000", "--device"])
        bad = False
    except ValueError:
        bad = True
    check("--device 缺值报错", bad)
    check("参数解析", _parse_args(["50000-50002", "--device", "X", "--yes"]) == (["50000-50002"], "X", True))
    rows = build_offline_rows(
        ["65371", "65370"],
        {"65370": {"升级类型": "强制升级", "机型": "55U73SUA(01)", "设备扩展信息": "55U73SUA01299",
                   "源版本": "0100QQ0305", "目标版本": "0100UQ1104"}},
        {"65371": {"升级方式": "强制", "设备扩展信息": "X1", "源版本": "0100SQ0929", "目标版本": "Q0929_to_Q1104"}},
        {"65370": "2026-11-10 15:30:01", "65371": "2026-11-10 15:30:20"})
    check("汇总行升序、优先用部署库明细（完整目标版本）", rows[0]["Rule ID"] == "65370" and rows[0]["目标版本"] == "0100UQ1104")
    check("缺明细时用 CMS 行兜底", rows[1]["源版本"] == "0100SQ0929" and rows[1]["机型信息"] == "-")
    check("下线时间：同一分钟显示一个时间", offline_time_label(["2026-11-10 15:30:01", "2026-11-10 15:30:20"]) == "2026-11-10 15:30:01")
    check("下线时间：跨分钟显示区间", offline_time_label(["2026-11-10 15:30:01", "2026-11-10 15:41:00"])
          == "2026-11-10 15:30:01 ~ 2026-11-10 15:41:00")
    print("\n自测结果：", "全部通过 ✅" if good else "存在失败 ❌")
    return 0 if good else 1


if __name__ == "__main__":
    sys.exit(main())
