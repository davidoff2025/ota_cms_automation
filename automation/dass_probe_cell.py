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

            # IMPORTANT: the page has TWO areas — an upper query/filter strip
            # (大区域/方案号/... search boxes) and the lower fill-in GRID. Both use
            # fr-texteditor, so we must target the GRID only. We anchor on the
            # grid HEADER cell '方案号', get its geometry, then click the cell in
            # the SAME column, ONE data row below the header (the inserted blank
            # row). This never hits the upper filter boxes.
            print("\n用表头『方案号』锚定填报网格，定位其下方的数据单元格 ...")

            # Find the grid header cell whose exact text is 方案号 and is NOT a
            # filter label (filter labels end with '：').
            header = None
            cells = fr.locator("td")
            for idx in range(cells.count()):
                try:
                    el = cells.nth(idx)
                    if not el.is_visible():
                        continue
                    t = (el.inner_text() or "").strip()
                    if t == "方案号":  # grid header (filter label is '方案号：')
                        header = el
                        print(f"  找到表头『方案号』：td index={idx}")
                        break
                except Exception:
                    continue
            if header is None:
                print("  ✗ 未找到表头『方案号』，无法安全定位网格。")
                _shot(page, "probe_no_header")
                raise SystemExit(0)

            hb = header.bounding_box()
            print(f"  表头『方案号』位置：{hb}")
            # The inserted blank row sits DIRECTLY below the header row. Click the
            # same column, just below the header's bottom edge (first data row).
            row_h = hb["height"] if hb else 22
            click_x = hb["x"] + hb["width"] / 2
            click_y = hb["y"] + hb["height"] + row_h * 0.5  # middle of first row under header
            print(f"  单击数据行坐标：x={click_x:.0f} y={click_y:.0f}")
            page.mouse.click(click_x, click_y)
            page.wait_for_timeout(900)

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
