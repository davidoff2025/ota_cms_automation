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

            # Locate the DATA cell purely via DOM table geometry (no screen
            # coordinates, which kept hitting the upper filter boxes).
            #
            # The fill-in grid is a <table>. Find the header row that contains a
            # <td> whose text is EXACTLY '方案号' (the grid header; the filter
            # label is '方案号：'). Then take that same column index in the NEXT
            # table row — that's the first data cell (the inserted blank row).
            print("\n用 DOM 行列关系定位填报网格数据单元格（列=方案号，行=表头下一行）...")

            handle = fr.evaluate_handle(
                """() => {
                    const tables = Array.from(document.querySelectorAll('table'));
                    for (const tbl of tables) {
                        const rows = Array.from(tbl.rows);
                        for (let r = 0; r < rows.length; r++) {
                            const cells = Array.from(rows[r].cells);
                            const ci = cells.findIndex(c => (c.innerText||'').trim() === '方案号');
                            if (ci >= 0 && r + 1 < rows.length) {
                                const dataRow = rows[r + 1];
                                if (dataRow.cells[ci]) return dataRow.cells[ci];
                            }
                        }
                    }
                    return null;
                }"""
            )
            cell = handle.as_element()
            if cell is None:
                print("  ✗ 未能在任何 <table> 中找到『方案号』表头+数据行。")
                _shot(page, "probe_no_grid_cell")
                raise SystemExit(0)
            try:
                print(f"  命中数据单元格，文本={cell.inner_text()!r}")
            except Exception:
                pass
            cell.scroll_into_view_if_needed()
            cell.click()
            page.wait_for_timeout(900)

            # Pick the fr-texteditor that overlaps the clicked grid cell (so we
            # never type into an upper filter box).
            cb = cell.bounding_box()
            editor = fr.locator("input.fr-texteditor, textarea.fr-texteditor")
            vis = None
            for i in range(editor.count()):
                try:
                    e = editor.nth(i)
                    if not e.is_visible():
                        continue
                    eb = e.bounding_box()
                    if cb and eb and abs(eb["y"] - cb["y"]) < 30 and abs(eb["x"] - cb["x"]) < 200:
                        vis = e
                        break
                except Exception:
                    continue
            if vis is None:  # fallback: first visible editor
                for i in range(editor.count()):
                    e = editor.nth(i)
                    try:
                        if e.is_visible():
                            vis = e
                            break
                    except Exception:
                        continue

            if vis is None:
                print("  ✗ 单击后未出现可见的 fr-texteditor 编辑框。")
            else:
                eb = vis.bounding_box()
                print(f"  ✓ 出现 fr-texteditor 编辑框，位置={eb}（单元格位置={cb}）")
                try:
                    vis.fill("TEST123")
                    page.keyboard.press("Enter")
                    print("  ✓ 已填入 'TEST123' 并回车。请确认值进入了表格『方案号』列的数据行。")
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
