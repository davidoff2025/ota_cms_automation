"""Probe how a fill-in cell is edited: insert 1 row, click a cell, see the editor.

Run:
    cd ota_cms_automation
    ./.venv/bin/python -m automation.dass_probe_cell

Confirmed facts this builds on:
- The fill-in report lives in the iframe whose URL has 'entry_report' (frame[2]).
- '插入行' is <button class="fr-btn-text x-emb-add">; it opens a JS prompt
  '请输入插入行数'. We auto-answer it.
- New rows appear at the TOP, starting at table row 2 (row 1 = header).
- Cells are edited by a single click, plain text.

This probe inserts 1 row, clicks the first data cell, dumps whatever editor
element appears, types a test value, and screenshots — so we can confirm the
exact edit mechanism before writing the full fill-in script.
"""
from __future__ import annotations

import sys
import time

from playwright.sync_api import sync_playwright

from . import config
from .dass_login import login, navigate_to_fillin, _shot


def _report_frame(page):
    """Return the FineReport fill-in frame (URL contains 'entry_report')."""
    for fr in page.frames:
        if "entry_report" in (fr.url or ""):
            return fr
    # Fallback: the frame with the most <td>.
    best, best_n = None, -1
    for fr in page.frames:
        try:
            n = fr.locator("td").count()
            if n > best_n:
                best, best_n = fr, n
        except Exception:
            continue
    return best


def main() -> int:
    rows_to_insert = 1  # probe with a single row
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not config.HEADED)
        context = browser.new_context()
        page = context.new_page()

        # Auto-answer the '请输入插入行数' prompt with our row count.
        def _on_dialog(dialog):
            print(f"  [对话框] type={dialog.type} message={dialog.message!r}")
            if dialog.type == "prompt":
                dialog.accept(str(rows_to_insert))
            else:
                dialog.accept()
        page.on("dialog", _on_dialog)

        try:
            login(page)
            navigate_to_fillin(page)
            page.wait_for_timeout(3000)

            fr = _report_frame(page)
            if fr is None:
                print("✗ 未找到填报表 frame")
                _shot(page, "probe_no_frame")
                return 1
            print(f"填报表 frame：{fr.url[:80]!r}")

            tds_before = fr.locator("td").count()
            print(f"插入前 <td> 数量：{tds_before}")

            print("点击『插入行』(将自动应答行数为 1) ...")
            fr.locator("button.x-emb-add", has_text="插入行").first.click()
            page.wait_for_timeout(2500)

            tds_after = fr.locator("td").count()
            print(f"插入后 <td> 数量：{tds_after}（变化 {tds_after - tds_before}）")
            _shot(page, "probe_after_insert")

            # Click the first data cell (row 2). We locate editable cells by the
            # FineReport editable-cell class; fall back to the 2nd row's cells.
            print("尝试单击第 2 行第 1 个数据单元格，观察编辑控件 ...")
            # Common FineReport editable cell containers.
            candidates = [
                "td.x-editor-cell", "td[class*='editor']", "td[class*='x-cell']",
            ]
            clicked = False
            for sel in candidates:
                loc = fr.locator(sel)
                if loc.count() > 0:
                    print(f"  找到可编辑单元格选择器：{sel}（{loc.count()} 个）")
                    loc.first.click()
                    clicked = True
                    break
            if not clicked:
                # Just click a cell in the body table near the top.
                body_tds = fr.locator("td")
                print("  用通用 td 兜底：单击靠前的一个单元格")
                body_tds.nth(min(60, body_tds.count() - 1)).click()

            page.wait_for_timeout(1200)
            # After clicking, an editor input/textarea usually appears.
            editors = fr.locator("input:visible, textarea:visible")
            print(f"单击后可见 input/textarea：{editors.count()}")
            for i in range(min(editors.count(), 6)):
                el = editors.nth(i)
                try:
                    print(f"  editor[{i}] tag={el.evaluate('e=>e.tagName.toLowerCase()')} "
                          f"class={el.get_attribute('class')!r} "
                          f"id={el.get_attribute('id')!r}")
                except Exception:
                    continue

            # Try typing a test value into the first visible editor.
            if editors.count() > 0:
                try:
                    editors.first.fill("TESTVALUE")
                    page.keyboard.press("Enter")
                    print("  已向第一个编辑框输入 'TESTVALUE' 并回车。")
                except Exception as e:
                    print(f"  输入测试值失败：{e}")
            _shot(page, "probe_after_type")

            print("\n✅ 预演完成。请把终端输出和 artifacts/probe_* 截图发回。")
            if config.HEADED:
                try:
                    input("\n浏览器保持打开，查看完毕按回车关闭... ")
                except EOFError:
                    time.sleep(120)
            return 0
        except Exception as e:
            print(f"✗ 出错：{e}")
            _shot(page, "probe_error")
            return 1
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    sys.exit(main())
