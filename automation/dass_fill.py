"""DASS 自动填报：支持 init（插入并填报）与 update（匹配已存在行并更新字段）两种模式。

用法：
    cd ota_cms_automation
    ./.venv/bin/python -m automation.dass_fill [OPTION] "/路径/更新后的文件.xlsx"
    # OPTION 与 Excel 路径都是位置参数，顺序不限、均可省略。
    # 不传 Excel 则取 ota_v4 输出目录里最新的 .xlsx（见 OTA_V4_OUTPUTS）。

OPTION 取值（直接写关键字，不加 --update 前缀）：
    init    插入 N 行并逐格填入（默认，等同于不传 OPTION）。
    ota50 / ota200 / ota1k / ota5k / ota20k / ota-all
            数据行已存在，不插入行：用『粗筛批量匹配 + 精准回退』两阶段定位已存在行。
            Stage1（粗筛）按 方案号 ∈ Excel集合、目标版本 ∈ Excel集合、
            （递进选项再加）当前阶段 == 上一阶段(PREV_STAGE) 圈出候选行，再用 ruleid
            确认是否属于本批 Excel（在集合内才更新，不在则不动）。Stage2（回退）对
            Stage1 未命中的 ruleid，按 方案号+目标版本+设备扩展信息 三字段精确匹配
            （恰好一行才更新；零行报警跳过；多行歧义跳过不猜）。ota50 无上一阶段，
            Stage1 不过滤『当前阶段』。命中行更新『当前阶段』『当前阶段时间』（以及
            ota50 另写『是否展示』=1；ota-all 等其它选项不触碰『是否展示』）。

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

# 各递进选项在 DASS『当前阶段』里要找的上一阶段；ota50 无上一阶段，不参与此过滤。
PREV_STAGE = {
    "ota200": "OTA50",
    "ota1k": "OTA200",
    "ota5k": "OTA1000",
    "ota20k": "OTA5000",
    "ota-all": "OTA20000",
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


# existing_row_info 读取每个数据行的这几列文本；Stage2 精准回退键是其中的子集
# （方案号+目标版本+设备扩展信息）。『当前阶段』另行附带（见 existing_row_info），不并入此表。
MATCH_COLUMNS = ["方案号", "设备扩展信息", "机型", "目标版本", "ruleid"]


def existing_row_info(fr, letters: dict) -> list[dict]:
    """返回每个已存在数据行的 {n, 方案号/设备扩展信息/机型/目标版本/ruleid/当前阶段: 文本}。

    类比 blank_row_numbers：从『方案号』列 id 解析行号 N，按各匹配列的列字母读取
    该行对应单元格 .trim() 后的 innerText。仅返回『方案号』列非空的真实数据行
    （排除空白插入行与表头行）。letters 由 read_column_layout 提供列字母映射。

    除 MATCH_COLUMNS 外，额外附带『当前阶段』文本（Stage1 递进选项需要用它过滤）；
    若表头没有『当前阶段』列字母，则该字段置空串（不丢弃行，是否终止由 run_update 决定）。
    """
    plan_letter = letters.get("方案号")
    if not plan_letter:
        return []
    # MATCH_COLUMNS 之外再读『当前阶段』；该列不并入 MATCH_COLUMNS（后者是 Stage2 精准键）。
    col_letters = {col: letters.get(col) for col in MATCH_COLUMNS}
    col_letters["当前阶段"] = letters.get("当前阶段")
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
                for (const col in colLetters){
                    const letter=colLetters[col];
                    if(!letter){ rec[col]=''; continue; }  // 列字母缺失 -> 该字段置空，不丢弃行
                    let cell=null;
                    for(const c of dataTbl.rows[r].cells){
                        if((c.id||'').indexOf(letter+n+'-')===0){ cell=c; break; }
                    }
                    rec[col]=cell? (cell.innerText||'').trim() : '';
                }
                res.push(rec);
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
    """update 路径：两阶段『粗筛批量匹配 + 精准回退』更新已存在行。返回统计字典。

    不插入行、不应答任何 prompt。本批次的权威标识是 Excel 所有记录的 ruleid 集合。

    Stage1（粗筛批量匹配 + ruleid 确认）：扫描 DASS 已存在行，候选条件须全部满足——
      · 方案号 非空且 ∈ Excel 方案号集合；
      · 目标版本 ∈ Excel 目标版本集合；
      · （递进选项 option != ota50）当前阶段 == PREV_STAGE[option]；ota50 跳过此条件。
    对候选行按 ruleid 确认：ruleid ∈ Excel 集合 -> 更新（Stage1 命中）；
    ∉ 集合 -> 不动（仅 debug 提示）。

    Stage2（精准回退）：对 Stage1 未命中的每个 Excel ruleid，用其 Excel 记录的
      方案号 + 目标版本 + 设备扩展信息 三字段精确相等匹配已存在行——
      恰好一行 -> 更新（Stage2 命中）；零行 -> WARNING 跳过；多于一行 -> WARNING 歧义跳过不猜。

    命中行写入：当前阶段=STAGE_LABELS[option]，当前阶段时间=今日 YYYYMMDD；
    ota50 额外写 是否展示=1（ota-all 及其它选项不触碰 是否展示）。
    每个 Excel ruleid 至多更新一次（Stage1 已满足者不进入 Stage2）。
    所有比较均为对单元格显示文本的精确相等（仅沿用 existing_row_info 的 .trim()）。
    """
    stats = {
        "stage1_updated": 0,
        "stage2_updated": 0,
        "skipped_no_match": 0,
        "skipped_ambiguous": 0,
        "candidates_skipped_not_in_excel": 0,
    }

    layout = read_column_layout(fr, debug=True)
    letters, session = layout["letters"], layout["session"]
    print(f"会话锚：{session}")
    if session is None or "方案号" not in letters:
        print("✗ 未读到表头列布局，终止。")
        _shot(page, "update_no_anchor")
        return stats

    # —— 匹配所需列的表头保证：缺任何一个都报错终止，绝不退化为更弱的匹配键 ——
    # Stage1 用 方案号+目标版本[+当前阶段]，Stage2 用 方案号+目标版本+设备扩展信息，
    # 两阶段都需要 ruleid 做确认/回退键；机型不参与新匹配，故不要求。
    required_match = ["方案号", "目标版本", "设备扩展信息", "ruleid"]
    if option != "ota50":
        required_match.append("当前阶段")  # 递进选项 Stage1 需按上一阶段过滤
    missing_match = [c for c in required_match if c not in letters]
    if missing_match:
        print(f"✗ 表头缺少匹配所需列 {missing_match}，无法安全匹配，终止更新。")
        return stats

    # —— 待写列的表头保证：缺任何一个都报错终止，绝不静默跳过该字段 ——
    write_cols = ["当前阶段", "当前阶段时间"]
    if option == "ota50":
        write_cols.append("是否展示")
    missing_write = [c for c in write_cols if c not in letters]
    if missing_write:
        print(f"✗ 表头缺少待写列 {missing_write}，无法更新对应字段，终止更新。")
        return stats

    stage_label = STAGE_LABELS[option]
    stage_time = datetime.now().strftime("%Y%m%d")

    # —— 构建 Excel 侧集合与按 ruleid 的记录索引（整批构建一次，不逐行重算）——
    excel_targets = {rec.get("目标版本", "") for rec in records if rec.get("目标版本", "")}
    excel_scheme_nos = {rec.get("方案号", "") for rec in records if rec.get("方案号", "")}
    excel_by_ruleid: dict[str, dict] = {}
    for rec in records:
        rid = rec.get("ruleid", "")
        if not rid:
            print(f"⚠ 跳过 ruleid 为空的 Excel 记录（无法作为可靠匹配键）：方案号={rec.get('方案号', '')!r}")
            continue
        if rid not in excel_by_ruleid:  # 去重：同一 ruleid 取首条记录
            excel_by_ruleid[rid] = rec
    excel_ruleids = set(excel_by_ruleid)

    if not excel_ruleids:
        print("✗ Excel 中没有可用的 ruleid（全为空或无数据），没有可靠匹配键，终止更新。")
        return stats
    if not excel_targets:
        print("⚠ Excel 的『目标版本』集合为空，Stage1 粗筛无法按目标版本圈定，"
              "所有 ruleid 将直接走 Stage2 精准回退。")

    existing = existing_row_info(fr, letters)
    print(f"已存在数据行数：{len(existing)}")

    satisfied_ruleids: set[str] = set()

    def _write_row(row_n: int, tag: str, ruleid: str) -> bool:
        """对一行写入本次更新字段；全部 fill_cell 成功才算更新成功。"""
        updates = {"当前阶段": stage_label, "当前阶段时间": stage_time}
        if option == "ota50":
            updates["是否展示"] = "1"
        print(f"\n更新行 {row_n}（{tag}，ruleid={ruleid!r}）：")
        ok_row = True
        for col, val in updates.items():
            if not fill_cell(fr, page, row_n, letters[col], session, val):
                ok_row = False
        return ok_row

    # ———— Stage1：粗筛批量匹配 + ruleid 确认 ————
    for row in existing:
        scheme = row.get("方案号", "")
        if not (scheme and scheme in excel_scheme_nos):
            continue
        if row.get("目标版本", "") not in excel_targets:
            continue
        if option != "ota50" and row.get("当前阶段", "") != PREV_STAGE[option]:
            continue

        rid = row.get("ruleid", "")
        if rid in satisfied_ruleids:
            # 同一 ruleid 已更新过（DASS 可能存在重复行），不重复写。
            print(f"  · 候选行 {row['n']} 的 ruleid={rid!r} 已更新过，跳过。")
            continue
        if rid in excel_ruleids:
            if _write_row(row["n"], "Stage1", rid):
                stats["stage1_updated"] += 1
                satisfied_ruleids.add(rid)
        else:
            stats["candidates_skipped_not_in_excel"] += 1
            print(f"  · 候选行 {row['n']} 的 ruleid={rid!r} 不在 Excel 集合，跳过不动。")

    # ———— Stage2：对 Stage1 未命中的 ruleid 精准回退 ————
    for rid in excel_ruleids - satisfied_ruleids:
        rec = excel_by_ruleid[rid]
        k_scheme = rec.get("方案号", "")
        k_target = rec.get("目标版本", "")
        k_ext = rec.get("设备扩展信息", "")
        matches = [
            r for r in existing
            if r.get("方案号", "") == k_scheme
            and r.get("目标版本", "") == k_target
            and r.get("设备扩展信息", "") == k_ext
        ]
        if len(matches) == 1:
            if _write_row(matches[0]["n"], "Stage2", rid):
                stats["stage2_updated"] += 1
                satisfied_ruleids.add(rid)
        elif len(matches) == 0:
            print(f"⚠ Rule ID={rid!r} 精准匹配未找到任何行，跳过。"
                  f"方案号={k_scheme!r} 目标版本={k_target!r} 设备扩展信息={k_ext!r}")
            stats["skipped_no_match"] += 1
        else:
            print(f"⚠ Rule ID={rid!r} 精准匹配到多行（歧义 {len(matches)} 行），跳过不猜。"
                  f"方案号={k_scheme!r} 目标版本={k_target!r} 设备扩展信息={k_ext!r}")
            stats["skipped_ambiguous"] += 1

    print("\n======== 更新汇总 ========")
    print(f"  Stage-1 批量更新：{stats['stage1_updated']}")
    print(f"  Stage-2 精准回退更新：{stats['stage2_updated']}")
    print(f"  未匹配跳过：{stats['skipped_no_match']}")
    print(f"  歧义跳过：{stats['skipped_ambiguous']}")
    print(f"  候选但不在 Excel（未动）：{stats['candidates_skipped_not_in_excel']}")
    return stats


def _parse_cli(argv: list[str]) -> tuple[str, str | None]:
    """解析位置参数：可给『选项关键字』和/或『Excel 路径』，顺序不限。

    用法：
        python -m automation.dass_fill [OPTION] [EXCEL]
    OPTION 取值见 VALID_OPTIONS（init/ota50/ota200/ota1k/ota5k/ota20k/ota-all），
    不给则默认 init。EXCEL 不给则取 OTA_V4_OUTPUTS 下最新的 .xlsx。
    不再使用 --update 前缀，直接写关键字即可，例如：
        ./run_dass.sh ota50 "/路径/文件.xlsx"
    返回 (option, excel_path_or_None)。非法/多余参数直接报错退出。
    """
    option = "init"
    excel: str | None = None
    for arg in argv:
        if arg in ("-h", "--help"):
            print(
                "用法：python -m automation.dass_fill [OPTION] [EXCEL]\n"
                f"  OPTION 取值：{' / '.join(VALID_OPTIONS)}（默认 init）\n"
                "  EXCEL  可选 Excel 路径；不给则取 OTA_V4_OUTPUTS 下最新的 .xlsx。\n"
                "示例：\n"
                "  ./run_dass.sh \"/路径/文件.xlsx\"          # init，填空\n"
                "  ./run_dass.sh ota50 \"/路径/文件.xlsx\"    # 更新已存在行\n"
            )
            raise SystemExit(0)
        if arg in VALID_OPTIONS:
            option = arg
        elif arg.startswith("-"):
            print(f"✗ 未知参数：{arg!r}。用法：[OPTION] [EXCEL]，OPTION 取值 {VALID_OPTIONS}。")
            raise SystemExit(2)
        elif excel is None:
            excel = arg
        else:
            print(f"✗ 多余参数：{arg!r}。只接受一个 Excel 路径。")
            raise SystemExit(2)
    return option, excel


def main() -> int:
    option, excel = _parse_cli(sys.argv[1:])

    # ——— 浏览器启动之前：先完成选项校验与 Excel 读取，保证静态快速失败 ———
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
        print(f"读取到 {n} 条记录，将按『粗筛批量匹配 + 精准回退』更新已存在行（阶段：{STAGE_LABELS[option]}）。")

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
