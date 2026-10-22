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


def inspect_fillin(page) -> None:
    print("\n==================== 页面结构探测 ====================")
    frames = page.frames
    print(f"iframe 数量（含主框架）：{len(frames)}")
    for i, fr in enumerate(frames):
        print(f"  frame[{i}] name={fr.name!r} url={fr.url[:80]!r}")

    # Try to find the fill-in grid and toolbar in each frame.
    for i, fr in enumerate(frames):
        try:
            btns = fr.locator("button, a, div[role='button']")
            labels = []
            for j in range(min(btns.count(), 60)):
                el = btns.nth(j)
                try:
                    if not el.is_visible():
                        continue
                    t = (el.inner_text() or "").strip().replace("\n", " ")
                    if t and len(t) <= 20:
                        labels.append(t)
                except Exception:
                    continue
            if labels:
                uniq = list(dict.fromkeys(labels))
                print(f"\n  frame[{i}] 可见按钮/链接：{uniq[:40]}")
            # FineReport fill-in cells are usually <td> inside a report table.
            tds = fr.locator("td")
            print(f"  frame[{i}] <td> 数量：{tds.count()}")
        except Exception as e:
            print(f"  frame[{i}] 探测失败：{e}")


def main() -> int:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not config.HEADED)
        context = browser.new_context()
        page = context.new_page()
        try:
            login(page)
            print("\n请在浏览器里手动进入『OTA → 安卓升级状态填报』页面（如未自动到达）。")
            print("40 秒后自动探测当前页面结构...")
            if config.HEADED:
                time.sleep(40)
            _shot(page, "fillin_page")
            inspect_fillin(page)
            print("\n✅ 探测完成，请把以上输出和 artifacts/ 截图发回。")
            if config.HEADED:
                time.sleep(10)
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
