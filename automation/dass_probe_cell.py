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

            # The report body cells are <td> that carry a value span; the first
            # data row is the newly inserted blank row near the top. We locate the
            # grid body table and click a cell, then edit via the SHARED
            # 'fr-texteditor' floating editor (NOT the read-only sortable-* ones).
            print("\n定位数据单元格并用 fr-texteditor 填值 ...")

            # FineReport body cells usually live in table.x-table td, and each
            # editable cell has an inner div. Try to click the cell that currently
            # shows empty (the inserted row). As a robust approach, click the cell
            # just below the first visible 记录id value cell.
            body = fr.locator("table td")
            print(f"  body td 总数：{body.count()}")

            # Strategy: click a mid-grid cell (data area) and see the editor.
            target = None
            for idx in range(body.count()):
                try:
                    el = body.nth(idx)
                    if not el.is_visible():
                        continue
                    txt = (el.inner_text() or "").strip()
                    # The inserted blank data row has empty cells; pick the first
                    # empty, editable-looking body cell after the header area.
                    if idx > 40 and txt == "":
                        target = el
                        print(f"  选中空单元格：index={idx}")
                        break
                except Exception:
                    continue
            if target is None:
                target = body.nth(min(70, body.count() - 1))
                print("  未找到空单元格，兜底选一个靠中部的单元格")

            target.click()
            page.wait_for_timeout(800)

            editor = fr.locator("input.fr-texteditor, textarea.fr-texteditor")
            vis = None
            for i in range(editor.count()):
                try:
                    e = editor.nth(i)
                    if e.is_visible():
                        vis = e
                        break
                except Exception:
                    continue
            if vis is None:
                print("  ✗ 单击后未出现可见的 fr-texteditor 编辑框。")
            else:
                print(f"  ✓ 出现 fr-texteditor 编辑框，class={vis.get_attribute('class')!r}")
                try:
                    vis.click()
                    vis.fill("TEST123")
                    page.keyboard.press("Enter")
                    print("  ✓ 已填入 'TEST123' 并回车。请在截图/页面确认值是否进入该单元格。")
                except Exception as e:
                    print(f"  ✗ 填值失败：{e}")
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
