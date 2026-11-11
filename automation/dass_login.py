"""M4 step 1: log in to DASS (FineReport) and inspect the fill-in page.

Run:
    cd ota_cms_automation
    ./.venv/bin/python -m automation.dass_login

Reads DASS_USERNAME / DASS_PASSWORD from config/.env. Logs in, then reports the
page structure (iframes, the fill-in grid, buttons) so we can write the actual
fill-in script. FineReport fill-in tables are usually inside an <iframe> and use
non-standard cell widgets, so this inspection is required before automating.
"""
from __future__ import annotations

import sys
import time
from datetime import datetime

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

from . import config


def _shot(page, name: str) -> None:
    config.ARTIFACTS_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = config.ARTIFACTS_DIR / f"dass_{stamp}_{name}.png"
    try:
        page.screenshot(path=str(path), full_page=True)
        print(f"  截图：{path}")
    except Exception as e:
        print(f"  截图失败（{name}）：{e}")


def _first_visible(scope, *locators):
    for loc in locators:
        for i in range(loc.count()):
            item = loc.nth(i)
            try:
                if item.is_visible():
                    return item
            except Exception:
                continue
    return locators[0].first if locators else None


def login(page) -> None:
    print(f"打开 DASS 登录页：{config.DASS_LOGIN_URL}")
    page.goto(config.DASS_LOGIN_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(2000)
    _shot(page, "login_page")

    # Login form: 'User Name' / 'Password' / 'Sign In' (observed).
    user = _first_visible(
        page,
        page.get_by_placeholder("User Name"),
        page.locator("input[type='text']"),
    )
    pwd = _first_visible(
        page,
        page.get_by_placeholder("Password"),
        page.locator("input[type='password']"),
    )
    user.fill(config.dass_username())
    pwd.fill(config.dass_password())
    print("已填入 DASS 用户名/密码。")

    btn = _first_visible(
        page,
        page.get_by_role("button", name="Sign In"),
        page.get_by_text("Sign In", exact=False),
        page.locator("button, input[type='submit']"),
    )
    btn.click()
    print("已点击 Sign In，等待登录...")
    try:
        page.wait_for_url(lambda url: "signin" not in url.lower(), timeout=15000)
    except PWTimeout:
        print("⚠ 未检测到登录跳转（可能登录失败或页面结构不同）。")
    page.wait_for_timeout(3000)
    _shot(page, "after_login")
    print(f"当前地址：{page.url}")


def _click_menu_in_any_frame(page, text: str) -> bool:
    """Click the menu item that best matches `text`, searching every frame
    (the menu lives inside an iframe; labels may carry an icon glyph like \ue61f).
    Prefers the candidate whose trimmed text most closely equals `text` (so 'OTA'
    does not wrongly match 'Mobile Phone OTA'). Returns True if clicked."""
    best = None            # (score, element)
    for fr in page.frames:
        try:
            loc = fr.get_by_text(text, exact=False)
            for i in range(loc.count()):
                el = loc.nth(i)
                try:
                    if not el.is_visible():
                        continue
                    raw = (el.inner_text() or "").strip()
                    # Strip leading icon glyph(s) in the private-use area.
                    cleaned = "".join(ch for ch in raw if ch < "\ue000" or ch > "\uf8ff").strip()
                    # score: 0 = exact match (best), else extra-length penalty.
                    score = 0 if cleaned == text else len(cleaned)
                    if best is None or score < best[0]:
                        best = (score, el)
                except Exception:
                    continue
        except Exception:
            continue
    if best is not None:
        try:
            best[1].scroll_into_view_if_needed(timeout=2000)
            best[1].click()
            return True
        except Exception:
            return False
    return False


def menu_item_visible(page, text: str) -> bool:
    """菜单项（去掉图标字符后文字【恰好等于】text）当前是否可见。用来判断分组是否已经展开。"""
    for fr in page.frames:
        try:
            loc = fr.get_by_text(text, exact=False)
            for i in range(loc.count()):
                el = loc.nth(i)
                try:
                    if not el.is_visible():
                        continue
                    raw = (el.inner_text() or "").strip()
                    cleaned = "".join(ch for ch in raw if ch < "\ue000" or ch > "\uf8ff").strip()
                    if cleaned == text:
                        return True
                except Exception:
                    continue
        except Exception:
            continue
    return False


def expand_menu_path(page, groups: list[str], leaf_candidates: list[str]) -> str | None:
    """按当前菜单状态展开到叶子，返回实际可见的叶子文案（都不可见返回 None）。

    ★ 分组项（Android New / OTA）点一下是【展开/折叠切换】。以前无条件依次点击，菜单已经是展开状态
      时（例如刚用完另一个报表），再点反而把它折叠，叶子就找不到了。这里改成：只有下一层不可见时
      才点击当前分组；叶子已经可见就不再动分组。
    """
    def visible_leaf():
        return next((lf for lf in leaf_candidates if menu_item_visible(page, lf)), None)

    leaf = visible_leaf()
    if leaf:
        return leaf
    chain = list(groups)
    for i, g in enumerate(chain):
        nxt_visible = (menu_item_visible(page, chain[i + 1]) if i + 1 < len(chain) else bool(visible_leaf()))
        if nxt_visible:
            continue
        print(f"\n点击菜单：{g} ...")
        if _click_menu_in_any_frame(page, g):
            print(f"  ✓ 已点击 {g}")
        else:
            print(f"  ⚠ 未找到菜单项：{g}")
            _dump_visible_menu(page)
        page.wait_for_timeout(1500)
    return visible_leaf()


def _dump_visible_menu(page) -> None:
    print("  未点中。当前可见的菜单/文字项（供校准）：")
    seen = set()
    for fr in page.frames:
        try:
            loc = fr.locator("a, li, span, div[role='menuitem']")
            for i in range(min(loc.count(), 200)):
                el = loc.nth(i)
                try:
                    if not el.is_visible():
                        continue
                    t = (el.inner_text() or "").strip().replace("\n", " ")
                    if t and len(t) <= 24 and t not in seen:
                        seen.add(t)
                except Exception:
                    continue
        except Exception:
            continue
    for t in list(seen)[:60]:
        print(f"    - {t!r}")


def navigate_to_fillin(page) -> None:
    """Auto-navigate the 3-level left menu:
    『Android New』(expand) -> 『OTA』(expand) -> 『安卓升级状态填报』(open report).
    """
    leaf = "安卓升级状态填报"
    # 状态感知展开：菜单已展开就不再点分组（分组点击是切换，重复点会折叠）。
    expand_menu_path(page, ["Android New", "OTA"], [leaf])
    print(f"\n点击菜单：{leaf} ...")
    if _click_menu_in_any_frame(page, leaf):
        print(f"  ✓ 已点击 {leaf}")
    else:
        print(f"  ⚠ 未找到菜单项：{leaf}")
        _dump_visible_menu(page)
    page.wait_for_timeout(9000)  # 填报表打开较慢


def inspect_fillin(page) -> None:
    print("\n==================== 页面结构探测 ====================")
    frames = page.frames
    print(f"iframe 数量（含所有层级）：{len(frames)}")
    for i, fr in enumerate(frames):
        print(f"  frame[{i}] name={fr.name!r} url={fr.url[:90]!r}")

    for i, fr in enumerate(frames):
        try:
            # Report/fill-in cells: FineReport renders a big <table> of <td>.
            tds = fr.locator("td")
            inputs = fr.locator("input, textarea")
            editors = fr.locator("div.x-table-cell, div[class*='cell'], div[class*='editor']")
            print(f"\n  frame[{i}] <td>={tds.count()}  input/textarea={inputs.count()}  "
                  f"cell/editor-div={editors.count()}")

            # Toolbar / action buttons in this frame.
            btns = fr.locator("button, a, div[role='button'], input[type='button']")
            labels = []
            for j in range(min(btns.count(), 120)):
                el = btns.nth(j)
                try:
                    if not el.is_visible():
                        continue
                    t = (el.inner_text() or "").strip().replace("\n", " ")
                    if t and len(t) <= 24:
                        labels.append(t)
                except Exception:
                    continue
            if labels:
                print(f"  frame[{i}] 按钮：{list(dict.fromkeys(labels))[:40]}")

            # If this frame has a real grid, print the first row's cell texts to
            # help map columns.
            if tds.count() > 20:
                sample = []
                for j in range(min(tds.count(), 40)):
                    try:
                        t = (tds.nth(j).inner_text() or "").strip().replace("\n", " ")
                        sample.append(t[:16])
                    except Exception:
                        continue
                print(f"  frame[{i}] 前若干单元格文本：{sample}")
        except Exception as e:
            print(f"  frame[{i}] 探测失败：{e}")

    # Specifically hunt for the 『插入行』button (the only way to add rows).
    print("\n---- 查找『插入行』按钮 ----")
    for i, fr in enumerate(frames):
        try:
            for sel in ("text=插入行", "*:has-text('插入行')"):
                loc = fr.locator(sel)
                cnt = loc.count()
                if cnt:
                    # Report the smallest (leaf) matches.
                    shown = 0
                    for j in range(cnt):
                        el = loc.nth(j)
                        try:
                            if not el.is_visible():
                                continue
                            t = (el.inner_text() or "").strip().replace("\n", " ")
                            if t == "插入行" or (len(t) <= 8 and "插入行" in t):
                                tag = el.evaluate("e=>e.tagName.toLowerCase()")
                                cls = el.get_attribute("class") or ""
                                print(f"  frame[{i}] <{tag}> class={cls[:50]!r} text={t!r}")
                                shown += 1
                                if shown >= 3:
                                    break
                        except Exception:
                            continue
                    if shown:
                        break
        except Exception:
            continue


def main() -> int:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not config.HEADED)
        context = browser.new_context()
        page = context.new_page()

        # FineReport's 『插入行』shows a JS prompt ("请输入插入行数"). Playwright
        # auto-dismisses dialogs by default, which cancels it. Log it instead so
        # it stays visible / doesn't disrupt your manual exploration.
        def _on_dialog(dialog):
            print(f"  [对话框] type={dialog.type} message={dialog.message!r}")
            # Do NOT dismiss automatically during inspection — leave it be.
        page.on("dialog", _on_dialog)

        try:
            login(page)
            # Try auto-navigation first; it's fine if it doesn't fully work —
            # you then have time to finish manually during the wait below.
            try:
                navigate_to_fillin(page)
            except Exception as e:
                print(f"  自动导航出错（忽略，改为手动）：{e}")

            print("\n" + "=" * 60)
            print("请确认『安卓升级状态填报』表格已显示；并把表格横向滚动到最右，")
            print("让每行末尾的『插入行』按钮露出来。脚本将等待 90 秒后探测结构。")
            print("=" * 60)
            if config.HEADED:
                for remaining in range(90, 0, -15):
                    print(f"  还剩 {remaining} 秒...")
                    time.sleep(15)
            _shot(page, "fillin_page")
            inspect_fillin(page)
            print("\n✅ 探测完成，请把以上输出和 artifacts/ 截图发回。")
            # Keep the browser OPEN so you can keep exploring (e.g. click 插入行,
            # type a row count, see the dialog). It closes only when you press
            # Enter here in the terminal.
            if config.HEADED:
                try:
                    input("\n浏览器保持打开。查看完毕后，回到此终端按回车键关闭... ")
                except EOFError:
                    time.sleep(120)
            return 0
        except Exception as e:
            print(f"✗ 出错：{e}")
            _shot(page, "error")
            return 1
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    sys.exit(main())
