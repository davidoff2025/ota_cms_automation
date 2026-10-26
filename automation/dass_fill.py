"""DASS 自动填报：支持 init（插入并填报）与 update（匹配已存在行并更新字段）两种模式。

用法：
    cd ota_cms_automation
    ./.venv/bin/python -m automation.dass_fill [--update OPTION] "/路径/更新后的文件.xlsx"
    # 不传位置参数则取 ota_v4 输出目录里最新的 .xlsx（见 OTA_V4_OUTPUTS）

--update OPTION 取值：
    init    插入 N 行并逐格填入（默认，等同于不传 --update）。
    ota50 / ota200 / ota1k / ota5k / ota20k / ota-all
            数据行已存在，不插入行：按 方案号/设备扩展信息/机型/目标版本/ruleid
            五字段精确匹配定位已存在行，更新『当前阶段』『当前阶段时间』（以及
            ota50 另写『是否展示』=1）。

确认过的机制（见探测过程）：
- 填报表在 URL 含 'entry_report' 的 iframe 内。
- 『插入行』是 <button class="x-emb-add">，点后弹 JS prompt『请输入插入行数』，
  自动应答 N —— 仅 init 路径注册此对话框处理，update 路径不插入行、不应答 prompt。
- 表头/数据是两张 table.x-table（均 32 列，按列对齐）。
- 单元格 id = '{列字母}{行号}-0-{会话锚}'：列字母标识列（方案号=D、国家=C…），
  行号从 1(表头) 递增（首条数据=2，插入的空行从 3 起），会话锚全表一致、每次会话会变。
  列字母与会话锚都在运行时从表头实时读出。
- 单元格需【双击】进入编辑态，再 fill + 回车。
- 提交按钮为 FineReport 控件：校验 button.fr-btn-text.x-emb-verify（Data validation）、
  提交 button.fr-btn-text.x-emb-submit（Submit）。
- 填完后先截图并暂停，等你在终端输入 approve 才会执行『验证 -> 提交』；
  输入其它内容则不提交。绝不在未经确认时自动提交。
"""
from __future__ import annotations

import argparse
import sys
import time
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

from . import config
from .dass_login import login, navigate_to_fillin, _shot
from .dass_data import read_updated_excel

# --update 合法取值；init 为默认（插入并填报），其余为更新已存在行。
VALID_OPTIONS = ["init", "ota50", "ota200", "ota1k", "ota5k", "ota20k", "ota-all"]

# update 路径下，各选项写入『当前阶段』的阶段标签。
STAGE_LABELS = {
    "ota50": "OTA50",
    "ota200": "OTA200",
    "ota1k": "OTA1000",
    "ota5k": "OTA5000",
    "ota20k": "OTA20000",
    "ota-all": "完全开放",
}

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
             "源版本", "目标版本", "当前阶段", "当前阶段时间", "是否展示"]
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


# update 路径用于匹配的五个字段（Excel 值 vs DASS 单元格显示文本，精确相等）。
MATCH_COLUMNS = ["方案号", "设备扩展信息", "机型", "目标版本", "ruleid"]


def existing_row_info(fr, letters: dict) -> list[dict]:
    """返回每个已存在数据行的 {n: 行号, 方案号/设备扩展信息/机型/目标版本/ruleid: 文本}。

    类比 blank_row_numbers：从『方案号』列 id 解析行号 N，按各匹配列的列字母读取
    该行对应单元格 .trim() 后的 innerText。仅返回『方案号』列非空的真实数据行
    （排除空白插入行与表头行）。letters 由 read_column_layout 提供列字母映射。
    """
    plan_letter = letters.get("方案号")
    if not plan_letter:
        return []
    col_letters = {col: letters.get(col) for col in MATCH_COLUMNS}
    rows = fr.evaluate(
        """({planLetter, colLetters}) => {
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
                const pc=dataTbl.rows[r].cells[colIndex];
                const m=pc? (pc.id||'').match(new RegExp('^'+planLetter+'(\\\\d+)-')):null;
                if(!m) continue;
                const n=parseInt(m[1]);
                const rec={n:n};
                let ok=true;
                for (const col in colLetters){
                    const letter=colLetters[col];
                    if(!letter){ ok=false; break; }
                    let cell=null;
                    for(const c of dataTbl.rows[r].cells){
                        if((c.id||'').indexOf(letter+n+'-')===0){ cell=c; break; }
                    }
                    rec[col]=cell? (cell.innerText||'').trim() : '';
                }
                if(ok) res.push(rec);
            }}
            return res;
        }""",
        {"planLetter": plan_letter, "colLetters": col_letters},
    )
    # 仅保留『方案号』非空的真实数据行。
    return [r for r in rows if r.get("方案号", "") != ""]


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


def _click_exact_button(fr, page, selector: str, label: str) -> bool:
    """点击指定 class 的可见按钮（FineReport 控件）。成功返回 True。

    绝不回退到其它按钮；找不到/不可见时打印清晰信息并返回 False。
    """
    loc = fr.locator(selector)
    for i in range(loc.count()):
        el = loc.nth(i)
        try:
            if el.is_visible():
                el.scroll_into_view_if_needed(timeout=2000)
                el.click()
                print(f"  ✓ 已点击『{label}』按钮（{selector}）")
                page.wait_for_timeout(1500)
                return True
        except Exception:
            continue
    print(f"  ✗ 未找到可见的『{label}』按钮（{selector}）。")
    return False


def approve_and_submit(fr, page) -> bool:
    """校验 -> 提交（精确 class 定位）。找不到提交按钮时安全返回 False，绝不盲目点击其它按钮。"""
    btns = list_toolbar_buttons(fr)
    print("  当前可见按钮：")
    for b in btns:
        print(f"    {b['text']!r:14} <{b['tag']} class={b['cls']!r}>")

    # 1) 数据校验（Data validation）—— 若按钮不存在/不可见，打印信息后继续尝试提交。
    _click_exact_button(fr, page, "button.fr-btn-text.x-emb-verify", "数据校验/Data validation")

    # 2) 提交（Submit）—— 找不到则返回 False，绝不点击 Save/删除行等其它按钮。
    submitted = _click_exact_button(fr, page, "button.fr-btn-text.x-emb-submit", "提交/Submit")
    if not submitted:
        print("  ✗ 未找到『提交/Submit』按钮，未提交。请手动在页面提交。")
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


def run_init(fr, page, records) -> bool:
    """init 路径：点『插入行』-> 定位 N 个空行 -> 逐格填入。返回是否全部成功。

    行为与原脚本一致。插入行的点击与 prompt 应答在 main() 中完成（仅 init 注册对话框）。
    """
    n = len(records)
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
        return False
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
    return ok_all


def run_update(fr, page, records, option: str) -> dict:
    """update 路径：按五字段精确匹配已存在行并更新字段。返回统计字典。

    不插入行、不应答任何 prompt。逐条 Excel 记录：
      - 五字段（方案号/设备扩展信息/机型/目标版本/ruleid）精确相等匹配已存在行；
      - 恰好一个 -> 更新；零个 -> WARNING 跳过；多于一个 -> WARNING（歧义）跳过。
    更新字段：当前阶段=STAGE_LABELS[option]，当前阶段时间=今日 YYYYMMDD，
    ota50 额外写 是否展示=1（ota-all 及其它 otaXX 不触碰 是否展示）。
    """
    stats = {"updated": 0, "skipped_no_match": 0, "skipped_ambiguous": 0}

    layout = read_column_layout(fr, debug=True)
    letters, session = layout["letters"], layout["session"]
    print(f"会话锚：{session}")
    if session is None or "方案号" not in letters:
        print("✗ 未读到表头列布局，终止。")
        _shot(page, "update_no_anchor")
        return stats

    # 必须能在表头定位所有五个匹配列（含 ruleid），否则报错终止，绝不退化为四字段匹配。
    missing_match = [c for c in MATCH_COLUMNS if c not in letters]
    if missing_match:
        print(f"✗ 表头缺少匹配列 {missing_match}（需要全部五字段含 ruleid），无法安全匹配，终止更新。")
        return stats

    # 确定本次要写入的目标列。
    write_cols = ["当前阶段", "当前阶段时间"]
    if option == "ota50":
        write_cols.append("是否展示")
    missing_write = [c for c in write_cols if c not in letters]
    if missing_write:
        print(f"✗ 表头缺少待写列 {missing_write}，无法更新对应字段，终止更新。")
        return stats

    stage_label = STAGE_LABELS[option]
    stage_time = datetime.now().strftime("%Y%m%d")

    existing = existing_row_info(fr, letters)
    print(f"已存在数据行数：{len(existing)}")

    for rec in records:
        key = {c: rec.get(c, "") for c in MATCH_COLUMNS}
        matches = [
            row for row in existing
            if all(row.get(c, "") == key[c] for c in MATCH_COLUMNS)
        ]
        if len(matches) == 0:
            print(f"⚠ 未匹配到任何已存在行，跳过。匹配键：{key}")
            stats["skipped_no_match"] += 1
            continue
        if len(matches) > 1:
            print(f"⚠ 匹配到多行（歧义 {len(matches)} 行），跳过。匹配键：{key}")
            stats["skipped_ambiguous"] += 1
            continue

        row_n = matches[0]["n"]
        print(f"\n更新行 {row_n}（匹配键：{key}）：")
        ok_row = True
        updates = {"当前阶段": stage_label, "当前阶段时间": stage_time}
        if option == "ota50":
            updates["是否展示"] = "1"
        for col, val in updates.items():
            letter = letters.get(col)
            if not letter:
                print(f"    ✗ 列『{col}』的列字母未解析到，未写入该字段。")
                ok_row = False
                continue
            if not fill_cell(fr, page, row_n, letter, session, val):
                ok_row = False
        if ok_row:
            stats["updated"] += 1

    print("\n======== 更新汇总 ========")
    print(f"  已更新：{stats['updated']}")
    print(f"  未匹配跳过：{stats['skipped_no_match']}")
    print(f"  歧义跳过：{stats['skipped_ambiguous']}")
    return stats


def main() -> int:
    parser = argparse.ArgumentParser(
        prog="automation.dass_fill",
        description="DASS 填报：init 插入并填报 / otaXX 匹配已存在行并更新阶段字段。",
    )
    parser.add_argument(
        "--update",
        choices=VALID_OPTIONS,
        default="init",
        metavar="OPTION",
        help="操作模式：%(choices)s（默认 init）。",
    )
    parser.add_argument(
        "excel",
        nargs="?",
        default=None,
        help="可选 Excel 路径；不传则取 OTA_V4_OUTPUTS 下最新的 .xlsx。",
    )
    args = parser.parse_args()
    option = args.update

    # ——— 浏览器启动之前：先完成选项校验与 Excel 读取，保证静态快速失败 ———
    excel = args.excel
    if not excel:
        from .dass_data import latest_updated_excel
        excel = latest_updated_excel(OTA_V4_OUTPUTS)
    print(f"操作模式：{option}")
    print(f"数据源 Excel：{excel}")
    records = read_updated_excel(excel)  # 文件不存在时抛 FileNotFoundError，快速失败
    n = len(records)
    if n == 0:
        print("✗ Excel 中没有数据行。")
        return 1
    if option == "init":
        print(f"读取到 {n} 条记录，将插入 {n} 行并填写。")
    else:
        print(f"读取到 {n} 条记录，将按五字段精确匹配已存在行并更新（阶段：{STAGE_LABELS[option]}）。")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not config.HEADED)
        context = browser.new_context()
        page = context.new_page()

        # 仅 init 路径注册插入行 prompt 的应答；update 路径不插入行、不应答任何 prompt。
        if option == "init":
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

            if option == "init":
                ok_all = run_init(fr, page, records)
                _shot(page, "fill_done")
                print(f"\n{'✅' if ok_all else '⚠'} 填写完成（尚未验证、尚未提交）。")
            else:
                run_update(fr, page, records, option)
                _shot(page, "update_done")
                print("\n✅ 更新完成（尚未验证、尚未提交）。")

            if config.HEADED:
                print("\n======== 人工核对环节 ========")
                print("请先在浏览器里逐行检查刚填入/更新的数据是否正确。")
                print("确认无误后，输入 approve 回车，脚本才会执行『验证 -> 提交』；")
                print("若发现问题，输入其它任意内容（或直接回车）即取消，不做任何验证/提交。")
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
