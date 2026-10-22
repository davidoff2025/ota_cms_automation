"""M1: log in to the CMS and open the OTA upgrade page.

Run:
    cd ota_cms_automation
    ./.venv/bin/python -m automation.login

Reads CMS_USERNAME / CMS_PASSWORD from config/.env. Opens a visible browser by
default (CMS_HEADED=true) so you can watch it log in. Saves screenshots to
artifacts/ and leaves the browser open briefly so you can confirm.
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
    path = config.ARTIFACTS_DIR / f"{stamp}_{name}.png"
    try:
        page.screenshot(path=str(path), full_page=True)
        print(f"  截图：{path}")
    except Exception as e:
        print(f"  截图失败（{name}）：{e}")


def login(page) -> None:
    """Fill the login form and submit. Selectors match the observed login page:
    inputs with placeholders 'User name' / 'Password' and a 'confirm' button.
    """
    print(f"打开登录页：{config.CMS_LOGIN_URL}")
    page.goto(config.CMS_LOGIN_URL, wait_until="domcontentloaded")
    _shot(page, "login_page")

    # Locate fields by placeholder (robust to id/class changes). Fall back to
    # positional inputs if placeholders differ.
    user = page.get_by_placeholder("User name")
    pwd = page.get_by_placeholder("Password")
    if user.count() == 0:
        user = page.locator("input[type='text'], input:not([type])").first
    if pwd.count() == 0:
        pwd = page.locator("input[type='password']").first

    user.fill(config.cms_username())
    pwd.fill(config.cms_password())
    print("已填入用户名/密码。")

    # Click the confirm button (by visible text, case-insensitive).
    btn = page.get_by_role("button", name="confirm")
    if btn.count() == 0:
        btn = page.get_by_text("confirm", exact=False)
    btn.first.click()
    print("已点击 confirm，等待登录完成...")

    # Wait for navigation away from the login page.
    try:
        page.wait_for_url(lambda url: "login" not in url, timeout=15000)
    except PWTimeout:
        print("⚠ 登录后未检测到页面跳转（可能登录失败或页面结构不同）。")
    _shot(page, "after_login")
    print(f"当前地址：{page.url}")


def open_ota_page(page) -> None:
    print(f"打开 OTA 升级页：{config.CMS_OTA_URL}")
    page.goto(config.CMS_OTA_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(2000)
    _shot(page, "ota_page")
    print(f"当前地址：{page.url}")


def main() -> int:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not config.HEADED)
        context = browser.new_context()
        page = context.new_page()
        try:
            login(page)
            open_ota_page(page)
            print("\n✅ M1 完成：已登录并打开 OTA 页面。请查看 artifacts/ 下的截图确认。")
            if config.HEADED:
                print("浏览器将在 20 秒后关闭（可自行查看页面）...")
                time.sleep(20)
            return 0
        except Exception as e:
            print(f"✗ 运行出错：{e}")
            _shot(page, "error")
            return 1
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    sys.exit(main())
