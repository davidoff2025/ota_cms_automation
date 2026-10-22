"""探测导入弹窗：登录 -> 点『导入模板』-> 打印弹窗里的元素（文件框/名称/备注/按钮）。

Run:
    cd ota_cms_automation
    ./.venv/bin/python -m automation.inspect_import
"""
from __future__ import annotations

import sys
import time

from playwright.sync_api import sync_playwright

from . import config
from .login import login, _shot


def dump_dialog(page) -> None:
    print("\n==================== 弹窗/页面输入框 ====================")
    loc = page.locator("input, textarea")
    for i in range(min(loc.count(), 120)):
        el = loc.nth(i)
        try:
            if not el.is_visible():
                continue
            tag = el.evaluate("e=>e.tagName.toLowerCase()")
            print(f"  <{tag}> id={el.get_attribute('id')!r} "
                  f"name={el.get_attribute('name')!r} type={el.get_attribute('type')!r} "
                  f"placeholder={el.get_attribute('placeholder')!r}")
        except Exception:
            continue

    print("\n==================== 文件上传 input（含隐藏） ====================")
    files = page.locator("input[type='file']")
    print(f"  共 {files.count()} 个 input[type=file]")
    for i in range(files.count()):
        el = files.nth(i)
        try:
            print(f"  #{i} accept={el.get_attribute('accept')!r} "
                  f"visible={el.is_visible()} id={el.get_attribute('id')!r}")
        except Exception:
            continue

    print("\n==================== 弹窗里的按钮/标签 ====================")
    for sel in [".ant-modal button", ".ant-modal-title", ".ant-modal label",
                ".ant-modal .ant-form-item-label", "button"]:
        loc = page.locator(sel)
        for i in range(min(loc.count(), 40)):
            el = loc.nth(i)
            try:
                if not el.is_visible():
                    continue
                t = (el.inner_text() or "").strip().replace("\n", " ")
                if t and len(t) <= 30:
                    print(f"  [{sel}] {t!r}")
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
            page.wait_for_timeout(2500)

            print("点击『导入模板』...")
            btn = page.get_by_role("button", name="导入模板")
            if btn.count() == 0:
                btn = page.get_by_text("导入模板", exact=True)
            btn.first.click()
            page.wait_for_timeout(2000)
            _shot(page, "import_dialog")

            dump_dialog(page)

            print("\n浏览器保持打开 40 秒，便于你查看导入弹窗并截图。")
            if config.HEADED:
                time.sleep(40)
            return 0
        except Exception as e:
            print(f"✗ 出错：{e}")
            _shot(page, "inspect_import_error")
            return 1
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    sys.exit(main())
