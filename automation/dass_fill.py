"""DASS 自动填报：读指定 Excel -> 插入 N 行 -> 逐格填入（不点 Submit）。

用法：
    cd ota_cms_automation
    ./.venv/bin/python -m automation.dass_fill "/路径/更新后的文件.xlsx"
    # 不传参数则取 ota_v4 输出目录里最新的 .xlsx（见 OTA_V4_OUTPUTS）

确认过的机制（见探测过程）：
- 填报表在 URL 含 'entry_report' 的 iframe 内。
- 『插入行』是 <button class="x-emb-add">，点后弹 JS prompt『请输入插入行数』，自动应答 N。
- 表头/数据是两张 table.x-table（均 32 列，按列对齐）。
- 单元格 id = '{列字母}{行号}-0-{会话锚}'：列字母标识列（方案号=D、国家=C…），
  行号从 1(表头) 递增（首条数据=2，插入的空行从 3 起），会话锚全表一致、每次会话会变。
  列字母与会话锚都在运行时从表头实时读出。
- 单元格需【双击】进入编辑态，再 fill + 回车。
- 填完后先截图并暂停，等你在终端输入 approve 才会执行『验证/批复 -> 提交』；
  输入其它内容则不提交。绝不在未经确认时自动提交。
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

from . import config
from .dass_login import login, navigate_to_fillin, _shot
from .dass_data import read_updated_excel

# 要填写的 DASS 列（其余列：OTA类型/所有人/是否展示/状态/各时间/灰度档位 一律留空）。
# 『升级类型』填 Excel 的『升级方式』文本（正常升级/强制升级）。
FILL_COLUMNS = [
    "大区域", "国家", "方案号", "升级类型", "featurecode", "机型",
    "设备扩展信息", "品牌", "ruleid", "源版本", "目标版本", "当前阶段时间",
]

# 若实际表头不叫某个标准名，用这里的候选名去表头里匹配（取第一个存在的）。
HEADER_ALIASES = {
    "升级类型": ["升级类型", "OTA类型", "升级方式"],
}

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


def read_column_layout(fr, debug: bool = False) -> dict:
    """解析表头，返回 {col_letters: {列名: 列字母}, session: 会话锚}。

    FineReport 单元格 id 形如 '{列字母}{行号}-0-{会话锚}'，例如：
      记录id=A1-0-1431  国家=C1-0-1431  方案号=D1-0-1431  OTA类型=E1-0-1431
      机型=G1-0-1431   ruleid=L1-0-1431  目标版本=N1-0-1431 ...
    列字母标识列，行号从 1(表头) 递增，会话锚全表一致、每次会话会变。
    """
    known = ["记录id", "大区域", "国家", "方案号", "OTA类型", "升级类型", "升级方式",
             "featurecode", "机型", "设备扩展信息", "品牌", "所有人", "ruleid",
             "源版本", "目标版本", "当前阶段时间"]
    result = fr.evaluate(
        """(known) => {
            const out = {letters: {}, session: null, dump: []};
            const xts = Array.from(document.querySelectorAll('table.x-table'));
            for (const t of xts) {
                for (let r = 0; r < Math.min(t.rows.length, 2); r++) {
                    for (const c of t.rows[r].cells) {
                        const txt = (c.innerText||'').trim();
                        if (!txt || !known.includes(txt)) continue;
                        // id = <letters><rownum>-0-<session>
                        const m = (c.id||'').match(/^([A-Z]+)\\d+-\\d+-(.+)$/);
                        out.dump.push({txt, id: c.id||''});
                        if (m) {
                            if (!(txt in out.letters)) out.letters[txt] = m[1];
                            if (out.session === null) out.session = m[2];
                        }
                    }
                }
            }
            return out;
        }""",
        known,
    )
    if debug:
        print("  表头单元格转储（文字 / id）：")
        for d in result["dump"]:
            print(f"    {d['txt']!r:20} id={d['id']!r}")
    return {"letters": result["letters"], "session": result["session"]}


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


def list_toolbar_buttons(fr) -> list[dict]:
    """列出填报表里可见的按钮（文字 + class），用于定位验证/提交等操作。"""
    return fr.evaluate(
        """() => {
            const out = [];
            const els = document.querySelectorAll(
                "button, input[type=button], input[type=submit], .x-button, [class*='fr-btn'], [class*='x-emb']");
            for (const e of els) {
                const r = e.getBoundingClientRect();
                if (r.width === 0 || r.height === 0) continue;
                const txt = (e.innerText || e.value || '').trim();
                out.push({text: txt, cls: e.className || '', tag: e.tagName.toLowerCase()});
            }
            return out;
        }"""
    )


def _click_button_by_text(fr, page, keywords: list[str], label: str) -> bool:
    """点击文字包含 keywords 之一的可见按钮。成功返回 True。"""
    for kw in keywords:
        loc = fr.locator(
            "button, input[type=button], input[type=submit], .x-button, "
            "[class*='fr-btn'], [class*='x-emb']",
            has_text=kw,
        )
        for i in range(loc.count()):
            el = loc.nth(i)
            try:
                if el.is_visible():
                    el.scroll_into_view_if_needed(timeout=2000)
                    el.click()
                    print(f"  ✓ 已点击『{label}』按钮（匹配 {kw!r}）")
                    page.wait_for_timeout(1500)
                    return True
            except Exception:
                continue
    # input[type=submit] 的文字在 value 上，has_text 匹配不到，补一轮 value 匹配
    for kw in keywords:
        loc = fr.locator(f"input[type=submit][value*='{kw}'], input[type=button][value*='{kw}']")
        for i in range(loc.count()):
            el = loc.nth(i)
            try:
                if el.is_visible():
                    el.click()
                    print(f"  ✓ 已点击『{label}』按钮（value 匹配 {kw!r}）")
                    page.wait_for_timeout(1500)
                    return True
            except Exception:
                continue
    return False


def approve_and_submit(fr, page) -> bool:
    """验证/批复 -> 提交。找不到对应按钮时安全返回 False，绝不盲目提交。"""
    btns = list_toolbar_buttons(fr)
    print("  当前可见按钮：")
    for b in btns:
        print(f"    {b['text']!r:14} <{b['tag']} class={b['cls']!r}>")

    # 1) 验证 / 批复（若存在）
    _click_button_by_text(fr, page, ["验证", "校验", "批复", "审批", "通过"], "验证/批复")

    # 2) 提交 / Submit
    submitted = _click_button_by_text(fr, page, ["提交", "Submit", "submit", "保存"], "提交")
    if not submitted:
        print("  ✗ 未找到『提交』按钮，未提交。请手动在页面提交。")
        return False
    _shot(page, "submitted")
    print("  ✅ 已提交。")
    return True


def fill_cell(fr, page, row_n: int, letter: str, session: str, value: str) -> bool:
    """双击 #{letter}{row_n}-0-{session} 进入编辑，填 value，回读校验。"""
    if value is None or str(value) == "":
        return True
    cell_id = f"{letter}{row_n}-0-{session}"
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

            layout = read_column_layout(fr, debug=True)
            letters, session = layout["letters"], layout["session"]
            print(f"会话锚：{session}")

            def resolve_letter(col):
                """把目标列解析到它的列字母；必要时用表头别名回退。"""
                for name in HEADER_ALIASES.get(col, [col]):
                    if name in letters:
                        return letters[name], name
                return None, None

            print("列字母映射：", {c: resolve_letter(c)[0] for c in FILL_COLUMNS})
            missing0 = [c for c in FILL_COLUMNS if resolve_letter(c)[0] is None]
            if session is None or "方案号" not in letters:
                print("✗ 未读到表头列布局，终止。")
                _shot(page, "fill_no_anchor")
                return 1
            if missing0:
                print(f"⚠ 以下列在表头未找到，将跳过：{missing0}")

            rows_n = blank_row_numbers(fr, n)
            print(f"空白插入行（行号）：{rows_n}")
            if len(rows_n) < n:
                print(f"⚠ 只找到 {len(rows_n)} 个空行，少于 {n}；按可用的填。")

            ok_all = True
            for rec, row_n in zip(records, rows_n):
                print(f"\n填写行 {row_n}：")
                for col in FILL_COLUMNS:
                    letter, header_name = resolve_letter(col)
                    if not letter:
                        continue
                    val = rec.get(col, "")
                    if val == "":
                        continue
                    if not fill_cell(fr, page, row_n, letter, session, val):
                        ok_all = False

            _shot(page, "fill_done")
            print(f"\n{'✅' if ok_all else '⚠'} 填写完成（尚未提交）。请在浏览器核对每一行。")

            if config.HEADED:
                print("\n核对无误后，输入 approve 回车即执行『验证/批复 -> 提交』；")
                print("输入其他任意内容（或直接回车）则不提交、仅关闭浏览器。")
                try:
                    ans = input("是否批复并提交？[approve/提交 确认，其它取消]：").strip().lower()
                except EOFError:
                    ans = ""
                if ans in ("approve", "提交", "submit", "yes", "y"):
                    print("\n开始批复并提交 ...")
                    ok_submit = approve_and_submit(fr, page)
                    if ok_submit:
                        print("\n✅ 记录已提交。可在浏览器确认结果。")
                    else:
                        print("\n⚠ 未能自动提交，请在浏览器手动完成。")
                    try:
                        input("\n按回车关闭浏览器 ... ")
                    except EOFError:
                        time.sleep(60)
                else:
                    print("已取消提交（未做任何批复/提交操作）。")
                    try:
                        input("按回车关闭浏览器 ... ")
                    except EOFError:
                        time.sleep(30)
            else:
                print("非可见模式（HEADED=0），不进行交互提交；请人工核对后提交。")
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
