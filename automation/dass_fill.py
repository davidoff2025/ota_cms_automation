"""DASS 自动填报：支持 init（插入并填报）与 update（匹配已存在行并更新字段）两种模式。

用法：
    cd ota_cms_automation
    ./.venv/bin/python -m automation.dass_fill [OPTION] "/路径/更新后的文件.xlsx"
    # OPTION 与 Excel 路径都是位置参数，顺序不限、均可省略。
    # 不传 Excel 则取 ota_v4 输出目录里最新的 .xlsx（见 OTA_V4_OUTPUTS）。

OPTION 取值（直接写关键字，不加 --update 前缀）：
    init    插入 N 行并逐格填入（默认，等同于不传 OPTION）。
    ota50 / ota200 / ota1k / ota5k / ota20k / ota-all
            数据行已存在，不插入行：用【查询驱动】的『粗筛批量匹配 + 精选回退』两阶段定位。
            报表默认只显示有限的一页（显示行数 limit 10），所以不扫描默认显示的行，而是把
            条件填进顶部红色过滤框（方案号/目标版本/设备扩展信息/当前阶段）、点绿色『查询』，
            以查询返回的行作为匹配来源。显示行数默认设 'limit 100'。
            Stage1（粗筛查询）按 方案号 + 目标版本 +（递进选项再加）当前阶段==上一阶段
            (PREV_STAGE) 查询，再用 ruleid 确认是否属于本批 Excel（在集合内才更新，不在则
            不动）；Stage1 不填设备扩展信息。Stage2（精选回退）对 Stage1 未命中的 ruleid，
            用 方案号+目标版本+设备扩展信息 三字段查询并按 ruleid 定位（恰好一行才更新；
            零行报警跳过；多行歧义跳过不猜）。ota50 无上一阶段，Stage1 不过滤『当前阶段』。
            命中行更新『当前阶段』『当前阶段时间』（以及 ota50 另写『是否展示』=1；ota-all
            等其它选项不触碰『是否展示』）。每次查询后会话锚会变，实现会在查询后重读表头布局。

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

import re
import sys
import time
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

from . import config
from ._log import ok, warn, err, color
from .dass_login import login, navigate_to_fillin, _shot
from .dass_data import read_updated_excel
# 注意：dass_perf 复用本模块的常量/查询件，于 main() 内延迟导入以避免模块级循环导入。

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


def _strip_group_separators(s) -> str:
    """去掉千分位分组逗号（半角 ',' 与全角 '，'）。

    FineReport 的数字格式单元格会在【显示】文本里注入千分位逗号（如 '20,261,026'、
    'OTA,200'），而底层存储/我们键入的原值是无逗号的（见 probe-findings.md 的
    DISPLAY-ONLY 判定）。本函数只剥离分组逗号，不做其它改写（不转 int、不动空白），
    供回读校验与 ruleid 归一化共用。
    """
    return str(s or "").replace(",", "").replace("，", "")


def norm_ruleid(value) -> str:
    """归一化 ruleid 以便两边精确比较。

    DASS 页面把 ruleid 当数字展示，带千分位逗号（如 '65,197'），而 Excel 里是纯数字
    '65197'。直接字符串比较永远不相等，导致匹配全部落空。

    此外在【模糊查找】返回的结果里，ruleid 可能以逗号作为分隔符呈现（半角 '65,196'
    或全角 '65，196'），甚至夹带空白 / 全角空格 / 千分位点等。为了让两侧稳定对齐，
    这里把所有可能的分隔符（以及其它一切非数字字符）统统去掉，只保留数字后再比较，
    使 '65,197' / '65，197' / '65 197' 等都归一化为同一个键 '65197'。
    不转 int，避免前导零等语义丢失。
    """
    return re.sub(r"\D", "", str(value or ""))

# —— 查询(筛选) UI 选择器（来自只读探测 .agents/tasks/probe-findings.md）——
# 顶部红色标注的过滤框：把方案号/目标版本/设备扩展信息/当前阶段填进去，点绿色『查询』，
# 查询返回的行即为后续匹配/更新的数据来源（不再扫描默认显示的那十来行）。
FILTER_INPUTS = {
    "方案号": "input[name='PLANCODE']",
    "目标版本": "input[name='NEXTVERSION']",   # UI 标签显示为『标版本』
    "设备扩展信息": "input[name='DEVICEMSG']",
    "当前阶段": "input[name='CURRENTSTAGE']",
}
# 其余过滤框（清空用，避免上一次查询的残留条件干扰本次）。
FILTER_INPUTS_OTHERS = {
    "大区域": "input[name='REGION']",
    "国家": "input[name='COUNTRYCODE']",
    "机型": "input[name='DEVICECODE']",
    "源版本": "input[name='CURRENTVERSION']",
}
# 显示行数输入框：其内容被原样拼进数据集 SQL 的 LIMIT 子句，必须写完整片段
# 'limit N'（只写数字会报 SQL 语法错误）。见探测结论。
LIMIT_INPUT = "input[name='LIMIT']"
# 绿色『查询』按钮（FineReport 自由按钮 div id）。
QUERY_BUTTON = "#fr-btn-BTNSUBMIT"
# 默认显示行数（粗筛起始）设为 100；查询结果数 >= 此值视为可能被截断，升到 LIMIT_MAX 重查。
LIMIT_START = 100
LIMIT_MAX = 100


def set_display_limit(fr, page, n: int) -> bool:
    """设置『显示行数』过滤框为完整 SQL 片段 'limit N'（探测确认只写数字会报错）。"""
    loc = fr.locator(LIMIT_INPUT).first
    try:
        loc.scroll_into_view_if_needed(timeout=2000)
        loc.click()
        loc.fill("")
        loc.fill(f"limit {int(n)}")
        page.wait_for_timeout(150)
        print(f"  显示行数 已设为 'limit {int(n)}'")
        return True
    except Exception as e:
        print(color(f"  ✗ 设置显示行数失败（{LIMIT_INPUT}）：{e}", "red"))
        return False


def set_filter(fr, page, values: dict) -> None:
    """按 字段->值 填充顶部过滤框，并清空其它过滤框（避免残留条件干扰）。

    values 只含本次需要的字段（方案号/目标版本/设备扩展信息/当前阶段）；其余 FILTER_INPUTS
    与 FILTER_INPUTS_OTHERS 一律清空。值为空串的字段也会被清空。
    """
    def _set_one(sel: str, val: str) -> None:
        loc = fr.locator(sel).first
        try:
            loc.scroll_into_view_if_needed(timeout=2000)
            loc.click()
            loc.fill("")
            if val:
                loc.fill(str(val))
        except Exception as e:
            print(f"  ⚠ 过滤框 {sel} 设置失败：{e}")

    for field, sel in FILTER_INPUTS.items():
        _set_one(sel, values.get(field, ""))
    # 清空非本次使用的过滤框。
    for sel in FILTER_INPUTS_OTHERS.values():
        _set_one(sel, "")
    shown = {k: v for k, v in values.items() if v}
    print(f"  过滤条件：{shown}")


def click_query(fr, page) -> bool:
    """点击绿色『查询』按钮并等待网格重绘（显式等待表头重现，超时兜底 sleep）。

    探测确认：查询后网格 detach 重绘、会话锚会变；用 read_column_layout 轮询表头恢复。
    """
    loc = fr.locator(QUERY_BUTTON).first
    try:
        loc.scroll_into_view_if_needed(timeout=2000)
        loc.click()
    except Exception as e:
        print(f"  ✗ 点击『查询』失败（{QUERY_BUTTON}）：{e}")
        return False
    print("  ✓ 已点击『查询』，等待结果网格重绘 ...")
    waited, step, timeout_ms = 0, 800, 20000
    while waited < timeout_ms:
        page.wait_for_timeout(step)
        waited += step
        try:
            lay = read_column_layout(fr, debug=False)
            if lay.get("session") is not None and "方案号" in lay.get("letters", {}):
                page.wait_for_timeout(600)  # 略等数据行绘制完成
                return True
        except Exception:
            continue
    print(f"  ⚠ 等待 {timeout_ms}ms 后网格表头仍未就绪；已兜底继续。")
    return False


def read_query_results(fr) -> tuple[list[dict], dict, str]:
    """查询后重读表头布局并读取结果行。

    返回 (rows, letters, session)：rows 为 existing_row_info 读到的结果数据行；
    letters/session 来自查询后重读的 read_column_layout（探测确认查询后会话锚会变，
    必须重读，再用新的 letters/session 拼单元格 id）。
    """
    layout = read_column_layout(fr, debug=False)
    letters, session = layout["letters"], layout["session"]
    rows = existing_row_info(fr, letters) if letters.get("方案号") else []
    return rows, letters, session


def _run_query(fr, page, filters: dict) -> tuple[list[dict], dict, str]:
    """设 limit=LIMIT_START（默认 100）-> 填过滤框 -> 查询 -> 读结果；若未到上限且疑似被截断再升档重查。

    返回 (rows, letters, session)。截断警告在调用方按业务语义处理，这里只负责升档重查。
    """
    set_display_limit(fr, page, LIMIT_START)
    set_filter(fr, page, filters)
    click_query(fr, page)
    rows, letters, session = read_query_results(fr)
    print(f"  查询结果行数（limit {LIMIT_START}）：{len(rows)}")
    # 仅在起始 limit 小于上限时才升档重查；两者相等（默认 100）时已是上限，直接返回。
    if LIMIT_START < LIMIT_MAX and len(rows) >= LIMIT_START:
        print(f"  结果数 >= {LIMIT_START}，疑似被截断，升到 'limit {LIMIT_MAX}' 重查 ...")
        set_display_limit(fr, page, LIMIT_MAX)
        set_filter(fr, page, filters)
        click_query(fr, page)
        rows, letters, session = read_query_results(fr)
        print(f"  查询结果行数（limit {LIMIT_MAX}）：{len(rows)}")
    return rows, letters, session


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
    print(color(f"  ✗ 未找到可见的『{label}』按钮（{selector}）。", "red"))
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
        print(color("  ✗ 未找到『提交/Submit』按钮，未提交。请手动在页面提交。", "red"))
        return False
    _shot(page, "submitted")
    print(color("  ✅ 已提交。", "green"))
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


def format_init_detail(row_n: int, rec: dict, written: dict) -> str:
    """把一条 init 记录格式化成可检索的写入明细行（纯逻辑，供 run_init 打印与 selftest 断言）。

    含：行号、Rule ID、方案号、以及本行实际填入的各字段 `列名=值`（written 为实际被填的
    非空字段映射，含『当前阶段时间』等 init 写入字段）。
    """
    pairs = "  ".join(f"{col}={val!r}" for col, val in written.items())
    return (f"  [填报明细] 行号={row_n}  Rule ID={rec.get('ruleid', '')!r}  "
            f"方案号={rec.get('方案号', '')!r}  写入字段：{pairs}")


def format_update_detail(row_n: int, ruleid: str, scheme: str, tag: str, updates: dict) -> str:
    """把一行 update 写入格式化成可检索明细行（纯逻辑，供 _write_row 打印与 selftest 断言）。

    含：行号、Rule ID、方案号、来源阶段(Stage1/Stage2)、以及本次写入的 updates 各字段
    `列名=值`（『当前阶段』『当前阶段时间』，ota50 另含『是否展示=1』）。
    """
    pairs = "  ".join(f"{col}={val!r}" for col, val in updates.items())
    return (f"  [更新明细] 行号={row_n}  Rule ID={ruleid!r}  方案号={scheme!r}  "
            f"来源阶段={tag}  写入字段：{pairs}")


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
        print(color("✗ 未读到表头列布局，终止。", "red"))
        _shot(page, "fill_no_anchor")
        return False
    if missing0:
        print(color(f"⚠ 以下列在表头未找到，将跳过：{missing0}", "yellow"))

    rows_n = blank_row_numbers(fr, n)
    print(f"空白插入行（行号）：{rows_n}")
    if len(rows_n) < n:
        print(color(f"⚠ 只找到 {len(rows_n)} 个空行，少于 {n}；按可用的填。", "yellow"))

    ok_all = True
    for rec, row_n in zip(records, rows_n):
        print(f"\n填写行 {row_n}：")
        written: dict = {}  # 本行实际填入的非空字段（列名->值），供逐条明细打印
        for col in FILL_COLUMNS:
            letter, header_name = resolve_letter(col)
            if not letter:
                continue
            val = rec.get(col, "")
            if val == "":
                continue
            written[col] = val
            if not fill_cell(fr, page, row_n, letter, session, val):
                ok_all = False
        # C：逐条打印可检索的写入明细（行号 + Rule ID + 方案号 + 实际写入各字段值）。
        print(format_init_detail(row_n, rec, written))
    return ok_all


def _header_guard(letters: dict, option: str, phase: str) -> list[str]:
    """校验某阶段所需的列字母是否齐备；返回缺失列名列表（空列表=齐备）。

    phase='stage1' 需要 方案号/目标版本/ruleid（+递进选项的 当前阶段）与待写列；
    phase='stage2' 需要 方案号/目标版本/设备扩展信息/ruleid 与待写列。
    待写列：当前阶段/当前阶段时间（+ota50 的 是否展示）。两阶段都要能读 ruleid 做确认键。
    """
    if phase == "stage1":
        needed = ["方案号", "目标版本", "ruleid"]
        if option != "ota50":
            needed.append("当前阶段")
    else:  # stage2
        needed = ["方案号", "目标版本", "设备扩展信息", "ruleid"]
    write_cols = ["当前阶段", "当前阶段时间"]
    if option == "ota50":
        write_cols.append("是否展示")
    for c in write_cols:
        if c not in needed:
            needed.append(c)
    return [c for c in needed if c not in letters]


def run_update(fr, page, records, option: str) -> dict:
    """update 路径：查询驱动的两阶段『粗筛批量匹配 + 精准回退』更新已存在行。返回统计字典。

    不插入行、不应答任何 prompt。本批次的权威标识是 Excel 所有记录的 ruleid 集合。
    报表默认只显示有限的一页（显示行数 limit 10），所以不再扫描默认显示的行，而是把条件
    填进顶部红色过滤框、点绿色『查询』，以查询返回的行作为匹配/更新来源。

    前置（约定：一个 Excel = 同一个方案号、同一个目标版本）：
      · 方案号集合（正常应为 1 个；若 >1 则分组逐个处理，并告警）；
      · 目标版本集合（正常 1 个），作为 Stage1 的目标版本过滤值；
      · Excel ruleid 集合，以及按 ruleid 的记录索引（Stage2 取其 设备扩展信息）。

    Stage1（粗筛查询 + ruleid 确认）：
      查询条件 = 方案号 + 目标版本 +（递进选项再加 当前阶段=PREV_STAGE[option]；ota50 留空）。
      不填设备扩展信息（它只用于 Stage2 精选）。查询返回行中，ruleid ∈ Excel 集合且未更新过
      -> 更新（Stage1 命中）；ruleid ∉ Excel 集合 -> 不动。

    Stage2（精选查询，仅针对 Stage1 未满足的 ruleid）：
      对每个未满足 ruleid，用其 Excel 记录的 方案号 + 目标版本 + 设备扩展信息 三字段查询
      （当前阶段留空，避免手工改过的阶段把行藏起来）。在结果中按 ruleid 定位：
      恰好一行 -> 更新（Stage2 命中）；零行 -> WARNING 跳过；多行 -> WARNING 歧义跳过不猜。

    命中行写入：当前阶段=STAGE_LABELS[option]，当前阶段时间=今日 YYYYMMDD；
    ota50 额外写 是否展示=1（ota-all 及其它选项不触碰 是否展示）。
    每个 Excel ruleid 至多更新一次（Stage1 已满足者不进入 Stage2）。
    列字母/会话锚在【每次查询后】重读（探测确认查询会改变会话锚），用对应查询的
    letters/session 去拼单元格 id 并 fill_cell（双击进入编辑态 + 回读校验，保持不变）。
    所有比较均为对单元格显示文本的精确相等（仅沿用 existing_row_info 的 .trim()）。
    """
    stats = {
        "stage1_updated": 0,
        "stage2_updated": 0,
        "skipped_no_match": 0,
        "skipped_ambiguous": 0,
        "candidates_skipped_not_in_excel": 0,
        "truncation_warnings": 0,
    }

    stage_label = STAGE_LABELS[option]
    stage_time = datetime.now().strftime("%Y%m%d")

    # —— 构建 Excel 侧集合与按 ruleid 的记录索引（整批构建一次，不逐行重算）——
    excel_targets = {rec.get("目标版本", "") for rec in records if rec.get("目标版本", "")}
    excel_scheme_nos = {rec.get("方案号", "") for rec in records if rec.get("方案号", "")}
    # ruleid 统一用 norm_ruleid 归一化（去千分位逗号）后作为键，
    # 以便与 DASS 页面展示的『65,197』正确对上。
    excel_by_ruleid: dict[str, dict] = {}
    for rec in records:
        rid = norm_ruleid(rec.get("ruleid", ""))
        if not rid:
            print(f"⚠ 跳过 ruleid 为空的 Excel 记录（无法作为可靠匹配键）：方案号={rec.get('方案号', '')!r}")
            continue
        if rid not in excel_by_ruleid:  # 去重：同一 ruleid 取首条记录
            excel_by_ruleid[rid] = rec
    excel_ruleids = set(excel_by_ruleid)

    if not excel_ruleids:
        print(color("✗ Excel 中没有可用的 ruleid（全为空或无数据），没有可靠匹配键，终止更新。", "red"))
        return stats
    if not excel_scheme_nos:
        print(color("✗ Excel 中没有可用的方案号，无法作为查询条件，终止更新。", "red"))
        return stats
    if not excel_targets:
        print(color("⚠ Excel 的『目标版本』集合为空，Stage1 粗筛将不带目标版本条件，"
                    "所有 ruleid 可能直接走 Stage2 精选。", "yellow"))

    # 约定一个 Excel 对应一个方案号；若不止一个，告警并对每个方案号分组各跑一轮 Stage1。
    if len(excel_scheme_nos) > 1:
        print(color(f"⚠ 本 Excel 出现多个方案号 {sorted(excel_scheme_nos)!r}（约定应只有一个），"
                    "将对每个方案号分别执行 Stage1 粗筛。", "yellow"))
    # 目标版本正常只有一个；若多个则把它们都作为 Stage1 过滤值（逐个方案号×目标版本查询）。
    if len(excel_targets) > 1:
        print(color(f"⚠ 本 Excel 出现多个目标版本 {sorted(excel_targets)!r}（约定应只有一个），"
                    "将对每个目标版本分别查询。", "yellow"))

    satisfied_ruleids: set[str] = set()

    def _write_row(row_n: int, letters: dict, session: str, tag: str, ruleid: str) -> bool:
        """对一行写入本次更新字段；全部 fill_cell 成功才算更新成功。

        letters/session 必须来自【定位到该行的那次查询】重读的表头布局。
        """
        updates = {"当前阶段": stage_label, "当前阶段时间": stage_time}
        if option == "ota50":
            updates["是否展示"] = "1"
        print(f"\n更新行 {row_n}（{tag}，ruleid={ruleid!r}）：")
        ok_row = True
        for col, val in updates.items():
            if not fill_cell(fr, page, row_n, letters[col], session, val):
                ok_row = False
        # C：逐条打印可检索的更新明细（行号 + Rule ID + 方案号 + 来源阶段 + 写入字段值）。
        scheme = excel_by_ruleid.get(ruleid, {}).get("方案号", "")
        print(format_update_detail(row_n, ruleid, scheme, tag, updates))
        return ok_row

    # ———— Stage1：按 方案号×目标版本（+上一阶段）粗筛查询 ————
    # 目标版本为空时用单元素空串，表示 Stage1 不带目标版本条件（见上文告警）。
    target_values = sorted(excel_targets) if excel_targets else [""]
    for scheme in sorted(excel_scheme_nos):
        for target in target_values:
            print(f"\n====== Stage1 粗筛查询：方案号={scheme!r} 目标版本={target!r} ======")
            filters = {"方案号": scheme, "目标版本": target}
            if option != "ota50":
                filters["当前阶段"] = PREV_STAGE[option]  # 递进选项按上一阶段过滤
            rows, letters, session = _run_query(fr, page, filters)

            if session is None or "方案号" not in letters:
                print(color("✗ 查询后未读到表头列布局，跳过此次 Stage1 查询。", "red"))
                _shot(page, "update_no_anchor")
                continue
            missing = _header_guard(letters, option, "stage1")
            if missing:
                print(color(f"✗ 表头缺少 Stage1 所需列 {missing}，无法安全匹配，终止更新。", "red"))
                return stats
            if len(rows) >= LIMIT_MAX:
                print(color(f"⚠ 查询结果达到上限 {LIMIT_MAX}，结果可能被截断；按当前显示的行继续处理。", "yellow"))
                stats["truncation_warnings"] += 1

            for row in rows:
                rid = norm_ruleid(row.get("ruleid", ""))  # 归一化 DASS 侧『65,197』
                if rid in satisfied_ruleids:
                    print(f"  · 结果行 {row['n']} 的 ruleid={rid!r} 已更新过，跳过。")
                    continue
                if rid in excel_ruleids:
                    if _write_row(row["n"], letters, session, "Stage1", rid):
                        stats["stage1_updated"] += 1
                        satisfied_ruleids.add(rid)
                else:
                    stats["candidates_skipped_not_in_excel"] += 1
                    print(f"  · 结果行 {row['n']} 的 ruleid={rid!r} 不在 Excel 集合，跳过不动。")

    # ———— Stage2：对 Stage1 未命中的 ruleid 逐个精选查询 ————
    for rid in sorted(excel_ruleids - satisfied_ruleids):
        rec = excel_by_ruleid[rid]
        k_scheme = rec.get("方案号", "")
        k_target = rec.get("目标版本", "")
        k_ext = rec.get("设备扩展信息", "")
        print(f"\n====== Stage2 精选查询：ruleid={rid!r} 方案号={k_scheme!r} "
              f"目标版本={k_target!r} 设备扩展信息={k_ext!r} ======")
        # 精选三字段查询；当前阶段留空（手工改过的阶段不应把行藏起来）。
        rows, letters, session = _run_query(
            fr, page,
            {"方案号": k_scheme, "目标版本": k_target, "设备扩展信息": k_ext},
        )
        if session is None or "方案号" not in letters:
            print(color(f"✗ 查询后未读到表头列布局，Rule ID={rid!r} 跳过。", "red"))
            stats["skipped_no_match"] += 1
            continue
        missing = _header_guard(letters, option, "stage2")
        if missing:
            print(color(f"✗ 表头缺少 Stage2 所需列 {missing}，无法安全匹配，终止更新。", "red"))
            return stats
        if len(rows) >= LIMIT_MAX:
            print(color(f"⚠ 查询结果达到上限 {LIMIT_MAX}，结果可能被截断；按当前显示的行继续处理。", "yellow"))
            stats["truncation_warnings"] += 1

        matches = [r for r in rows if norm_ruleid(r.get("ruleid", "")) == rid]
        if len(matches) == 1:
            if _write_row(matches[0]["n"], letters, session, "Stage2", rid):
                stats["stage2_updated"] += 1
                satisfied_ruleids.add(rid)
        elif len(matches) == 0:
            print(color(f"⚠ Rule ID={rid!r} 精选查询结果中未找到该 ruleid，跳过。"
                        f"方案号={k_scheme!r} 目标版本={k_target!r} 设备扩展信息={k_ext!r}", "yellow"))
            stats["skipped_no_match"] += 1
        else:
            print(color(f"⚠ Rule ID={rid!r} 精选查询命中多行（歧义 {len(matches)} 行），跳过不猜。"
                        f"方案号={k_scheme!r} 目标版本={k_target!r} 设备扩展信息={k_ext!r}", "yellow"))
            stats["skipped_ambiguous"] += 1

    print("\n======== 更新汇总 ========")
    print(f"  Stage-1 批量更新：{stats['stage1_updated']}")
    print(f"  Stage-2 精选回退更新：{stats['stage2_updated']}")
    print(f"  未匹配跳过：{stats['skipped_no_match']}")
    print(f"  歧义跳过：{stats['skipped_ambiguous']}")
    print(f"  候选但不在 Excel（未动）：{stats['candidates_skipped_not_in_excel']}")
    if stats["truncation_warnings"]:
        print(color(f"  ⚠ 结果可能被截断的查询次数：{stats['truncation_warnings']}（已按显示行处理）", "yellow"))
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
        print(color("✗ Excel 中没有数据行。", "red"))
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

        # 延迟导入：dass_perf 在模块级依赖本模块的常量/查询件，放这里避免循环导入。
        from .dass_perf import GATE_OPTIONS, navigate_to_perf, check_success_rate

        try:
            login(page)

            # ———— 升级成功率闸门（仅 ota200 起的递进档位）————
            # 在『升级 DASS』(run_update) 之前，先去绩效页校验本批每个 Rule ID 的成功率：
            # 全部 >= 50% 才放行；任一 < 50% 或两阶段都查不到 -> abort（不 run_update、
            # 不进入人工提交环节）。ota50 / init 不校验。
            if option in GATE_OPTIONS:
                print("\n======== 升级成功率闸门（run_update 之前）========")
                navigate_to_perf(page)
                page.wait_for_timeout(2500)
                fr_perf = _report_frame(page)
                if fr_perf is None:
                    print(color("✗ 未找到绩效页 frame，无法校验成功率；安全失败，abort 升级。", "red"))
                    _shot(page, "perf_no_frame")
                    return 2
                if not check_success_rate(fr_perf, page, records, option):
                    print(color("\n⛔ 成功率闸门未通过：中止本次升级事务，"
                                "run_update 未执行，不进入人工提交环节。", "red"))
                    return 2
                print(color("\n✅ 成功率闸门通过，继续执行升级 DASS。", "green"))

            navigate_to_fillin(page)
            page.wait_for_timeout(2500)
            fr = _report_frame(page)
            if fr is None:
                print(color("✗ 未找到填报表 frame", "red"))
                return 1

            if option == "init":
                ok_all = run_init(fr, page, records)
                _shot(page, "fill_done")
                if ok_all:
                    print(color("\n✅ 填写完成（尚未验证、尚未提交）。", "green"))
                else:
                    print(color("\n⚠ 填写完成（尚未验证、尚未提交）。", "yellow"))
            else:
                run_update(fr, page, records, option)
                _shot(page, "update_done")
                print(color("\n✅ 更新完成（尚未验证、尚未提交）。", "green"))

            # 更新 DASS 数据后自动显示 OTA 部署汇总（数据取自本次更新用的 Excel）。
            # 仅当该 Excel 含 Rule ID 列时显示；没有则静默跳过（如 init 未回写 Rule ID）。
            try:
                from .summary import print_summary
                print("")
                print_summary(excel)
            except Exception as e:
                print(color(f"（部署汇总显示跳过：{e}）", "dim"))

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
                        print(color("\n✅ 记录已提交。可在浏览器确认结果。", "green"))
                    else:
                        print(color("\n⚠ 未能自动提交，请在浏览器手动完成。", "yellow"))
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
            print(color(f"✗ 出错：{e}", "red"))
            _shot(page, "fill_error")
            return 1
        finally:
            context.close()
            browser.close()


# —— 不依赖浏览器的纯函数自测（python -m automation.dass_fill --selftest）——
def _selftest() -> int:
    ok = True

    def check(name, cond):
        nonlocal ok
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        if not cond:
            ok = False

    # —— C：init 明细行格式 ——
    rec = {"ruleid": "65204", "方案号": "MT9655AndrGTVLas", "目标版本": "0100UQ1015"}
    written = {"方案号": "MT9655AndrGTVLas", "ruleid": "65204", "目标版本": "0100UQ1015",
               "当前阶段时间": "20260101"}
    init_line = format_init_detail(3, rec, written)
    check("init 明细含行号", "行号=3" in init_line)
    check("init 明细含 Rule ID", "Rule ID='65204'" in init_line)
    check("init 明细含方案号", "方案号='MT9655AndrGTVLas'" in init_line)
    check("init 明细含写入字段『当前阶段时间』", "当前阶段时间='20260101'" in init_line)

    # —— C：update 明细行格式（ota200，Stage1）——
    upd = {"当前阶段": "OTA200", "当前阶段时间": "20260101"}
    up_line = format_update_detail(5, "65204", "MT9655AndrGTVLas", "Stage1", upd)
    check("update 明细含行号", "行号=5" in up_line)
    check("update 明细含 Rule ID", "Rule ID='65204'" in up_line)
    check("update 明细含方案号", "方案号='MT9655AndrGTVLas'" in up_line)
    check("update 明细含来源阶段 Stage1", "来源阶段=Stage1" in up_line)
    check("update 明细含『当前阶段=OTA200』", "当前阶段='OTA200'" in up_line)
    check("update 明细含『当前阶段时间』", "当前阶段时间='20260101'" in up_line)

    # —— C：update 明细（ota50 变体，含是否展示=1；来源阶段 Stage2）——
    upd50 = {"当前阶段": "OTA50", "当前阶段时间": "20260101", "是否展示": "1"}
    up50_line = format_update_detail(7, "65210", "MT9655AndrGTVLas", "Stage2", upd50)
    check("ota50 update 明细含『是否展示=1』", "是否展示='1'" in up50_line)
    check("ota50 update 明细含来源阶段 Stage2", "来源阶段=Stage2" in up50_line)

    print("\n自测结果：", "全部通过 ✅" if ok else "存在失败 ❌")
    return 0 if ok else 1


if __name__ == "__main__":
    if "--selftest" in sys.argv[1:]:
        sys.exit(_selftest())
    sys.exit(main())
