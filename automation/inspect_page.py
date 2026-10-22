"""探测工具：登录后打印页面上的菜单/按钮/输入框，帮助编写精准的点击脚本。

Run:
    cd ota_cms_automation
    ./.venv/bin/python -m automation.inspect_page

它会：登录 -> 打开 OTA 页 -> 列出所有可见的可点击元素（菜单、按钮、链接）和
输入框（含 id / name / placeholder / 文本），并保存一张整页截图。把输出和截图发回，
即可据此写导入流程脚本。浏览器保持打开 40 秒，便于你手动展开左侧菜单后再观察。
"""
from __future__ import annotations

import sys
import time

from playwright.sync_api import sync_playwright

from . import config
from .login import login, _shot


def dump_clickables(page) -> None:
    print("\n==================== 可点击元素（菜单/按钮/链接）====================")
    # Collect menu items, buttons, links, and ant-design menu nodes.
    selectors = [
        "a", "button",
        "[role='menuitem']", ".ant-menu-item", ".ant-menu-submenu-title",
        "li.ant-menu-item", "span.ant-menu-title-content",
    ]
    seen = set()
    for sel in selectors:
        loc = page.locator(sel)
        n = min(loc.count(), 300)
        for i in range(n):
            el = loc.nth(i)
            try:
                if not el.is_visible():
                    continue
                text = (el.inner_text() or "").strip().replace("\n", " ")
                if not text or len(text) > 40:
                    continue
                el_id = el.get_attribute("id") or ""
                key = (sel, text, el_id)
                if key in seen:
                    continue
                seen.add(key)
                extra = f" id={el_id}" if el_id else ""
                print(f"  [{sel}] text={text!r}{extra}")
            except Exception:
                continue


def dump_inputs(page) -> None:
    print("\n==================== 输入框 ====================")
    loc = page.locator("input, textarea")
    n = min(loc.count(), 100)
    for i in range(n):
        el = loc.nth(i)
        try:
            if not el.is_visible():
                continue
            print(
                f"  <{el.evaluate('e=>e.tagName.toLowerCase()')}> "
                f"id={el.get_attribute('id')!r} "
                f"name={el.get_attribute('name')!r} "
                f"type={el.get_attribute('type')!r} "
                f"placeholder={el.get_attribute('placeholder')!r}"
            )
        except Exception:
            continue


def main() -> int:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not config.HEADED)
        context = browser.new_context()
        page = context.new_page()
        try:
            login(page)
            page.goto(config.CMS_OTA_URL, wait_until="domcontentloaded")
            page.wait_for_timeout(3000)
            _shot(page, "inspect_ota_page")

            dump_clickables(page)
            dump_inputs(page)

            print("\n浏览器保持打开 40 秒。你可以手动点开左侧"
                  "『整机升级管理 → OTA升级模版 → 导入模版』，"
                  "我们下一步据此写自动点击。")
            if config.HEADED:
                time.sleep(40)
            return 0
        except Exception as e:
            print(f"✗ 出错：{e}")
            _shot(page, "inspect_error")
            return 1
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    sys.exit(main())
