"""DASS 自动填报：读指定 Excel -> 插入 N 行 -> 逐格填入（不点 Submit）。

用法：
    cd ota_cms_automation
    ./.venv/bin/python -m automation.dass_fill "/路径/更新后的文件.xlsx"
    # 不传参数则取 ota_v4 输出目录里最新的 .xlsx（见 OTA_V4_OUTPUTS）

确认过的机制（见探测过程）：
- 填报表在 URL 含 'entry_report' 的 iframe 内。
- 『插入行』是 <button class="x-emb-add">，点后弹 JS prompt『请输入插入行数』，自动应答 N。
- 表头/数据是两张 table.x-table（均 32 列，按列对齐）。
- 单元格 id = 'D{行}-0-{列锚}'：表头行 D1，首条数据 D2，插入的空行从 D3 起。
  列锚每次会话都变，故运行时从表头实时读出。
- 单元格需【双击】进入编辑态，再 fill + 回车。
- 本脚本【不点 Submit】，填完截图并暂停，由你人工核对后手动提交。
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

from . import config
from .dass_login import login, navigate_to_fillin, _shot
from .dass_data import read_updated_excel

# 要填写的 DASS 列（其余列：所有人/是否展示/状态/各时间/灰度档位 一律留空）。
FILL_COLUMNS = [
    "大区域", "国家", "方案号", "OTA类型", "featurecode", "机型",
    "设备扩展信息", "品牌", "ruleid", "源版本", "目标版本", "当前阶段时间",
]

OTA_V4_OUTPUTS = str(Path.home() / "my_python_project" / "ota_v4" / "data" / "outputs")


def _report_frame(page):
    for fr in page.frames:
        if "entry_report" in (fr.url or ""):
            return fr
    best, best_n = None, -1
    for fr in page.frames:
        try:
            n = fr.locator("td").count()
            if n > best_n:
                best, best_n = fr, n
        except Exception:
            continue
    return best


def read_column_anchors(fr) -> dict:
    """{表头文字: 列锚}，从表头 x-table 的单元格 id 'D1-0-<锚>' 提取。"""
    return fr.evaluate(
        """() => {
            const xts = Array.from(document.querySelectorAll('table.x-table'));
            let headerTbl=null;
            for (const t of xts) {
                if (!t.rows.length) continue;
                const texts = Array.from(t.rows[0].cells).map(c => (c.innerText||'').trim());
                if (texts.includes('方案号')) { headerTbl=t; break; }
            }
            if (!headerTbl) return {};
            const map = {};
            for (const c of headerTbl.rows[0].cells) {
                const txt = (c.innerText||'').trim();
                const m = (c.id||'').match(/^D\\d+-\\d+-(.+)$/);
                if (txt && m) map[txt] = m[1];
            }
            return map;
        }"""
    )


def blank_row_numbers(fr, count: int) -> list[int]:
    """返回前 count 个空白插入行的行号 N（取『方案号』列为空的行，自顶向下）。"""
    rows = fr.evaluate(
        """() => {
            const xts = Array.from(document.querySelectorAll('table.x-table'));
            let headerTbl=null, colIndex=-1;
            for (const t of xts) {
                if (!t.rows.length) continue;
                const cells = Array.from(t.rows[0].cells);
                const ci = cells.findIndex(c => (c.innerText||'').trim()==='方案号');
                if (ci>=0){ headerTbl=t; colIndex=ci; break; }
            }
            let dataTbl=null, maxRows=-1;
            for (const t of xts){ if(t===headerTbl) continue;
                if(t.rows.length>maxRows){maxRows=t.rows.length; dataTbl=t;} }
            const res=[];
            if(dataTbl){ for(let r=0;r<dataTbl.rows.length;r++){
                const c=dataTbl.rows[r].cells[colIndex];
                const m=c? (c.id||'').match(/^D(\\d+)-/):null;
                res.push({n: m? parseInt(m[1]):null, text: c?(c.innerText||'').trim():'(x)'});
            }}
            return res;
        }"""
    )
    return [r["n"] for r in rows if r["text"] == "" and r["n"] is not None][:count]


def fill_cell(fr, page, row_n: int, anchor: str, value: str) -> bool:
    """双击 #D{row_n}-0-{anchor} 进入编辑，填 value，回读校验。"""
    if value is None or str(value) == "":
        return True
    cell_id = f"D{row_n}-0-{anchor}"
    cell = fr.locator(f"#{cell_id}")
    try:
        cell.scroll_into_view_if_needed(timeout=3000)
    except Exception:
        pass
    cb = cell.bounding_box()
    cell.dblclick()
    page.wait_for_timeout(250)
    ed = fr.locator("input.fr-texteditor, textarea.fr-texteditor")
    target = None
    for i in range(ed.count()):
        try:
            e = ed.nth(i)
            if not e.is_visible():
                continue
            eb = e.bounding_box()
            if cb and eb and abs(eb["y"] - cb["y"]) < 30 and abs(eb["x"] - cb["x"]) < 260:
                target = e
                break
        except Exception:
            continue
    if target is None:
        print(f"    ✗ {cell_id} 未出现对位编辑框")
        return False
    target.fill(str(value))
    page.keyboard.press("Enter")
    page.wait_for_timeout(120)
    try:
        got = fr.locator(f"#{cell_id}").inner_text().strip()
    except Exception:
        got = "?"
    ok = got == str(value).strip()
    print(f"    {'✓' if ok else '✗'} {cell_id} {value!r}" + ("" if ok else f" -> 实际 {got!r}"))
    return ok


def main() -> int:
    excel = sys.argv[1] if len(sys.argv) > 1 else None
    if not excel:
        from .dass_data import latest_updated_excel
        excel = latest_updated_excel(OTA_V4_OUTPUTS)
    print(f"数据源 Excel：{excel}")
    records = read_updated_excel(excel)
    n = len(records)
    if n == 0:
        print("✗ Excel 中没有数据行。")
        return 1
    print(f"读取到 {n} 条记录，将插入 {n} 行并填写。")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not config.HEADED)
        context = browser.new_context()
        page = context.new_page()

        def _on_dialog(dialog):
            print(f"  [对话框] {dialog.type}: {dialog.message!r}")
            if dialog.type == "prompt":
                dialog.accept(str(n))
            else:
                dialog.accept()
        page.on("dialog", _on_dialog)

        try:
            login(page)
            navigate_to_fillin(page)
            page.wait_for_timeout(2500)
            fr = _report_frame(page)
            if fr is None:
                print("✗ 未找到填报表 frame")
                return 1

            print(f"点击『插入行』，自动应答行数 {n} ...")
            fr.locator("button.x-emb-add", has_text="插入行").first.click()
            page.wait_for_timeout(2500)

            anchors = read_column_anchors(fr)
            print("列锚映射：", {c: anchors.get(c) for c in FILL_COLUMNS})
            if "方案号" not in anchors:
                print("✗ 未读到表头列锚，终止。")
                _shot(page, "fill_no_anchor")
                return 1
            missing = [c for c in FILL_COLUMNS if c not in anchors]
            if missing:
                print(f"⚠ 以下列在表头未找到，将跳过：{missing}")

            rows_n = blank_row_numbers(fr, n)
            print(f"空白插入行（D编号）：{rows_n}")
            if len(rows_n) < n:
                print(f"⚠ 只找到 {len(rows_n)} 个空行，少于 {n}；按可用的填。")

            ok_all = True
            for rec, row_n in zip(records, rows_n):
                print(f"\n填写行 D{row_n}：")
                for col in FILL_COLUMNS:
                    anchor = anchors.get(col)
                    if not anchor:
                        continue
                    val = rec.get(col, "")
                    if val == "":
                        continue
                    if not fill_cell(fr, page, row_n, anchor, val):
                        ok_all = False

            _shot(page, "fill_done")
            print(f"\n{'✅' if ok_all else '⚠'} 填写完成（未点 Submit）。请在浏览器核对。")
            print("确认无误后，请在 DASS 页面手动点击 Submit 提交。")
            if config.HEADED:
                try:
                    input("\n浏览器保持打开。核对完毕后按回车关闭（不会自动提交）... ")
                except EOFError:
                    time.sleep(120)
            return 0
        except Exception as e:
            print(f"✗ 出错：{e}")
            _shot(page, "fill_error")
            return 1
        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    sys.exit(main())
