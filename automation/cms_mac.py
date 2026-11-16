"""CMS「维度管理 → MAC组管理」：确保场测 MAC 都在『测试+设备类型代码』这个 MAC 组里（prep 调用）。

页面结构（2026-11-16 只读实测）：
  · 列表页 /basicservice/dimension-manage/fbmsop-mac：过滤『MAC组名称』(input#groupName，模糊匹配)、
    『设备类型』(antd v3 下拉)、『搜索』；列 MAC组名称 / 设备类型 / MAC组 / 编辑 / 删除。
    『MAC组』列直接列出区间：'28cddbd8-28cddbd8,7de6d0e5-7de6d0f4'（后 8 位，起-止）。
  · 『编 辑』/『新建』打开右侧抽屉（.ant-drawer，不是弹窗）：设备类型（下拉）、MAC组名称（input#groupName）、
    『增加』/『删除』、vxe 表格两列 MAC起(后8位) / MAC止(后8位)（每行两个 input），底部『取 消』/『保 存』。
    『增加』在表格末尾加一行空的输入框。『取 消』关闭抽屉，不保存。

流程：
  1) 选设备类型 + 名称过滤『测试<设备类型代码>』→ 搜索 → 按名称【精确】找到这一组，读『MAC组』列。
  2) 组存在：不在组里的 MAC → 编辑 → 每个 MAC 增加一行（起 = 止 = 后 8 位）→ 保存。都在 → OK，不动。
     组不存在：新建 → 设备类型 → 名称『测试<设备类型代码>』→ 每个 MAC 增加一行 → 保存。
  3) 保存前你输入 yes（--yes 跳过）；dry_run=True 时填好后点『取 消』，不保存。
  4) 保存后重新搜索，所有 MAC 都在组里才算成功。
"""
from __future__ import annotations

import re
import sys

from . import config
from ._log import color, err, ok, warn
from .login import _shot, login
from .mac_group import group_name, in_ranges, missing_macs, parse_ranges

GROUP_COL = "MAC组"


# ----------------------------------------------------------------------
# 页面操作
# ----------------------------------------------------------------------

def open_mac_page(page) -> None:
    page.goto(config.CMS_UPEVENT_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(2500)
    page.locator(".ant-menu-submenu-title", has_text="维度管理").first.click()
    page.wait_for_timeout(800)
    page.locator(".ant-menu-item", has_text=re.compile(r"^\s*MAC组管理\s*$")).first.click()
    page.locator("input#groupName").first.wait_for(state="visible", timeout=15000)
    page.wait_for_timeout(1200)
    print("  ✓ 已打开『MAC组管理』页")


def _pick_select(page, select, code: str) -> bool:
    """在一个 antd v3 下拉里输入设备类型代码并选中联想项（候选项文字形如『名称(代码)』）。"""
    try:
        select.click()
        page.wait_for_timeout(300)
        inp = select.locator("input.ant-select-search__field").first
        if inp.count():
            inp.fill("")
            inp.type(code, delay=50)
        else:
            page.keyboard.type(code, delay=50)
        opts = page.locator(".ant-select-dropdown:not(.ant-select-dropdown-hidden) .ant-select-dropdown-menu-item")
        for _ in range(25):
            page.wait_for_timeout(200)
            if opts.count() and opts.first.is_visible():
                break
        for i in range(opts.count()):
            o = opts.nth(i)
            if o.is_visible() and code in (o.inner_text() or ""):
                o.click()
                page.wait_for_timeout(400)
                break
        return code in (select.inner_text() or "")
    except Exception:
        return False


def search_group(page, device_code: str) -> tuple[dict | None, str]:
    """列表页按设备类型 + 名称搜『测试<代码>』。返回 (该组的行 {名称, 设备类型, MAC组} 或 None, 问题说明)。"""
    name = group_name(device_code)
    form_sel = page.locator(".ant-form-item").filter(
        has=page.locator(".ant-form-item-label", has_text="设备类型")).locator(".ant-select").first
    if not _pick_select(page, form_sel, device_code):
        return None, f"列表页『设备类型』选不上 {device_code!r}（CMS 里可能还没有这个设备类型）"
    box = page.locator("input#groupName").first
    box.fill(name)
    if (box.input_value() or "").strip() != name:
        return None, "『MAC组名称』过滤框没填进去"
    page.get_by_role("button", name=re.compile(r"搜\s*索")).first.click()
    page.wait_for_timeout(2500)
    rows = page.evaluate("""() => {
        const vis = e => e.offsetParent !== null;
        const ht = [...document.querySelectorAll('table')]
            .find(t => vis(t) && [...t.querySelectorAll('th')].some(h => (h.innerText||'').includes('MAC组名称')));
        if (!ht) return [];
        const heads = [...ht.querySelectorAll('th')].map(h => (h.innerText||'').replace(/\\s/g,''));
        const out = [];
        for (const t of document.querySelectorAll('table')) {
            if (!vis(t)) continue;
            for (const r of t.querySelectorAll('tbody tr')) {
                if (!vis(r)) continue;
                const c = [...r.querySelectorAll('td')];
                if (c.length < 3) continue;
                const o = {};
                heads.forEach((h, i) => { if (h && c[i]) o[h] = (c[i].innerText||'').trim(); });
                if (o['MAC组名称']) out.push(o);
            }
        }
        return out;
    }""") or []
    hit = [r for r in rows if r.get("MAC组名称", "").strip() == name]
    if len(hit) > 1:
        return None, f"有 {len(hit)} 个同名组『{name}』，不确定用哪个，未处理"
    return (hit[0] if hit else None), ""


def _drawer(page):
    return page.locator(".ant-drawer-open").last


def _drawer_btn(drawer, text: str):
    return drawer.locator("button").filter(has_text=re.compile(rf"^\s*{text}\s*$")).first


def _fill_new_rows(page, drawer, macs: list[str]) -> str:
    """每个 MAC：点『增加』→ 在新增的最后一行填 起 = 止 = 后 8 位。返回问题说明（空 = 成功）。"""
    for m in macs:
        before = drawer.locator("tbody tr").count()
        _drawer_btn(drawer, r"增\s*加").click()
        page.wait_for_timeout(600)
        rows = drawer.locator("tbody tr")
        if rows.count() <= before:
            return "点『增加』后没有出现新的一行"
        inputs = rows.nth(rows.count() - 1).locator("input")
        if inputs.count() < 2:
            return "新增的一行里没有『MAC起 / MAC止』两个输入框"
        for k in (0, 1):
            inputs.nth(k).fill(m)
            page.wait_for_timeout(150)
        got = [(inputs.nth(k).input_value() or "").strip().lower() for k in (0, 1)]
        if got != [m, m]:
            return f"MAC {m} 没填进去（看到 {got}）"
    return ""


def _drawer_values(drawer) -> list[tuple[str, str]]:
    vals = drawer.evaluate("""d => [...d.querySelectorAll('tbody tr')].filter(r => r.offsetParent)
        .map(r => [...r.querySelectorAll('input')].map(i => (i.value||'').trim().toLowerCase()))""") or []
    return [(v[0], v[1]) for v in vals if len(v) >= 2 and v[0]]


def _save_or_cancel(page, drawer, dry_run: bool) -> str:
    if dry_run:
        _drawer_btn(drawer, r"取\s*消").click()
        page.wait_for_timeout(1200)
        return ""
    _drawer_btn(drawer, r"保\s*存").click()
    for _ in range(30):
        page.wait_for_timeout(300)
        if page.locator(".ant-drawer-open").count() == 0:
            break
    from .active_ota import _detect_ui_error
    ui = _detect_ui_error(page)
    if ui:
        return "CMS 报错：" + " ".join(ui.split())
    if page.locator(".ant-drawer-open").count():
        return "点『保 存』后抽屉没有关闭（可能有必填项没填或 CMS 报错）"
    return ""


# ----------------------------------------------------------------------
# 对外入口
# ----------------------------------------------------------------------

def ensure_macs(p, device_code: str, macs: list[str], assume_yes: bool = False,
                dry_run: bool = False, read_only: bool = False) -> dict:
    """确保 macs（后 8 位）都在『测试<device_code>』组里。

    read_only=True：只查不改（prep --no-upload 用）；dry_run=True：填好后取消（测试模式）。
    返回 {"status": "ok" | "added" | "created" | "would_add" | "would_create" | "tested" | "cancelled" | "failed",
          "group": 组名, "missing": [...], "detail": 说明}。
    """
    name = group_name(device_code)
    res = {"status": "failed", "group": name, "missing": [], "detail": ""}
    if not macs:
        res.update(status="ok", detail="没有场测 MAC，不需要处理")
        return res
    mode = "（只查不改）" if read_only else ("（测试模式：填好后取消，不保存）" if dry_run else "")
    print(color(f"\n======== MAC 组检查：{name}{mode} ========", "bold"))
    print(f"  场测 MAC（后 8 位）：{', '.join(macs)}")
    browser = p.chromium.launch(headless=not config.HEADED)
    context = browser.new_context(viewport={"width": 1600, "height": 900})
    page = context.new_page()
    try:
        login(page)
        open_mac_page(page)
        row, problem = search_group(page, device_code)
        if problem:
            res["detail"] = problem
            print(err(problem))
            return res
        existing = parse_ranges(row.get(GROUP_COL, "")) if row else []
        missing = missing_macs(macs, existing) if row else list(macs)
        res["missing"] = missing
        if row:
            print(f"  已有组『{name}』（{row.get('设备类型', '')}），现有区间：{row.get(GROUP_COL, '') or '（空）'}")
            for m in macs:
                print(f"    {m}  " + (ok("已在组内") if m not in missing else warn("不在组内，需要加入")))
            if not missing:
                res.update(status="ok", detail=f"{len(macs)} 个 MAC 都已在组内")
                return res
        else:
            print(warn(f"  没有找到组『{name}』，需要新建，并加入 {len(macs)} 个 MAC"))

        action = f"在『{name}』里加入 {', '.join(missing)}" if row else f"新建『{name}』并加入 {', '.join(missing)}"
        if read_only:
            res.update(status="would_add" if row else "would_create", detail=f"需要{action}（--no-upload 只查不改）")
            return res
        if not dry_run and not assume_yes:
            if not sys.stdin.isatty():
                res.update(status="cancelled", detail="非交互终端无法确认，未修改（可加 --yes）")
                print(err(res["detail"]))
                return res
            try:
                go = input(f"  将{action}。输入 yes 执行，其它任意输入跳过：").strip().lower() == "yes"
            except EOFError:
                go = False
            if not go:
                res.update(status="cancelled", detail=f"你没有确认，未{action}")
                return res

        if row:   # —— 编辑已有组 ——
            page.locator("tbody tr").filter(has_text=name).first.locator("button").filter(
                has_text=re.compile(r"编\s*辑")).first.click()
        else:     # —— 新建 ——
            page.get_by_role("button", name=re.compile(r"新\s*建")).first.click()
        drawer = _drawer(page)
        drawer.wait_for(state="visible", timeout=10000)
        page.wait_for_timeout(800)
        if not row:
            sel = drawer.locator(".ant-form-item").filter(
                has=page.locator(".ant-form-item-label", has_text="设备类型")).locator(".ant-select").first
            if not _pick_select(page, sel, device_code):
                _drawer_btn(drawer, r"取\s*消").click()
                res["detail"] = f"新建窗口里『设备类型』选不上 {device_code!r}，未保存"
                print(err(res["detail"]))
                return res
            nm = drawer.locator("input#groupName").first
            nm.fill(name)
            if (nm.input_value() or "").strip() != name:
                _drawer_btn(drawer, r"取\s*消").click()
                res["detail"] = "新建窗口里『MAC组名称』没填进去，未保存"
                print(err(res["detail"]))
                return res
        else:
            shown = (drawer.locator("input#groupName").first.input_value() or "").strip()
            if shown != name:
                _drawer_btn(drawer, r"取\s*消").click()
                res["detail"] = f"打开的编辑窗口是『{shown}』，不是『{name}』，未保存"
                print(err(res["detail"]))
                return res
        problem = _fill_new_rows(page, drawer, missing)
        if problem:
            _shot(page, "mac_fill_fail")
            _drawer_btn(drawer, r"取\s*消").click()
            res["detail"] = problem + "，未保存"
            print(err(res["detail"]))
            return res
        vals = _drawer_values(drawer)
        print(f"  抽屉里的区间：{', '.join(f'{a}-{b}' for a, b in vals)}")
        _shot(page, "mac_before_save")
        problem = _save_or_cancel(page, drawer, dry_run)
        if problem:
            _shot(page, "mac_save_fail")
            res["detail"] = problem
            print(err(problem))
            return res
        if dry_run:
            res.update(status="tested", detail=f"测试模式：已填好（{action}），点了取消，CMS 未修改")
            print(ok(res["detail"]))
            return res

        # —— 回读校验 ——
        row2, problem = search_group(page, device_code)
        ranges = parse_ranges((row2 or {}).get(GROUP_COL, ""))
        still = [m for m in macs if not in_ranges(m, ranges)]
        if problem or not row2 or still:
            res["detail"] = problem or (f"保存后回读：组里仍缺 {', '.join(still)}" if row2 else "保存后搜不到这个组")
            print(err(res["detail"]))
            return res
        res.update(status="added" if row else "created",
                   detail=f"已{action}，回读确认 {len(macs)} 个 MAC 都在组内")
        print(ok(res["detail"]))
        return res
    except Exception as e:
        _shot(page, "mac_error")
        res["detail"] = f"出错：{e}"
        print(err(res["detail"]))
        return res
    finally:
        context.close()
        browser.close()
