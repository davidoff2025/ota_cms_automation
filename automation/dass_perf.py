"""升级成功率闸门：在『升级 DASS』(run_update) 之前，对本批 Excel 每个 Rule ID 的
升级成功率做校验。全部 >= 50% 才放行升级；任一 < 50%，或两阶段都查不到该 Rule ID，
则 abort 整个事务（run_update 不执行，也不进入人工提交环节）。

用法（通常由 dass_fill.main() 在 option != 'init' 分支里、run_update 之前调用）：
    from .dass_perf import navigate_to_perf, check_success_rate
    navigate_to_perf(page)
    if not check_success_rate(fr_perf, page, records, option):
        return 2  # abort

触发档位：仅当 option ∈ {ota200, ota1k, ota5k, ota20k, ota-all} 时校验。
ota50 / init 不校验（见 GATE_OPTIONS）。

入口页面：左侧菜单 Android New -> OTA -> 『安卓升级展示报表』（与 安卓升级状态填报 同级
的另一张 FineReport 报表）。其 frame 与填报页一样命中 URL 含 'entry_report' 的 iframe，
所以 dass_fill._report_frame(page) 可直接复用。

成功率来源：绩效页现成百分比列 **『升级成功率』**（列字母 Z），值形如 '28.26%'。
直接解析该百分比文本（去 % 转 float），不自行用设备数计算。

通过阈值：rate >= 50.0 判为通过（含等于，50.00% 放行）。

两阶段匹配（与 run_update 同构）：
  Stage1（模糊/粗筛）：按 方案号 + 目标版本 查询（limit 100），结果里按 norm_ruleid 对齐，
    命中则读其『升级成功率』。
  Stage2（精准回退）：对 Stage1 未查到的 Rule ID，用用户最初要求的四字段精准查询
    方案号 + 目标版本 + 设备扩展信息 + 当前阶段(上一阶段) 查询再定位（按 Rule ID 对齐）。

  设备扩展信息（见 .agents/tasks/perf-probe-findings.md 补探测实测）：
    绩效页顶部确实存在『设备扩展信息』过滤框，FineReport 参数名 = DEVICEMSG（与填报页一致）。
    但该框的 <input class="fr-trigger-texteditor"> 本身【没有】name 属性，参数名在其外层
    wrapper 的 widgetname 上，所以 input[name='DEVICEMSG'] 选不中；稳定选择器用
    div[widgetname='DEVICEMSG'] input.fr-trigger-texteditor（见 PERF_FILTER_INPUTS）。
    设备扩展信息过滤值取自该 Rule ID 对应 Excel 记录的『设备扩展信息』字段，与 run_update
    Stage2 的取值方式一致。『当前阶段』过滤值用该档位的上一阶段 PREV_STAGE[option]，与
    run_update Stage1 的递进逻辑一致（绩效页该列表头可能叫『当前策略』，但过滤框即 DEVICEMSG
    旁的 CURRENTSTAGE 框）。闸门对外可见行为（全部 >=50% 放行、否则 abort）不变。

安全失败：成功率/ruleid 列字母解析不到、或成功率文本无法解析（空值/缺列）时，一律按
『未找到/不通过』处理并 abort 升级，绝不默认放行。
"""
from __future__ import annotations

import re

from . import config
from ._log import color
from .dass_fill import (
    PREV_STAGE,
    norm_ruleid,
    click_query,
)
from .dass_login import _shot, _click_menu_in_any_frame, _dump_visible_menu

# 触发闸门校验的档位：从 ota200 起的递进档位。ota50 / init 不校验。
GATE_OPTIONS = {"ota200", "ota1k", "ota5k", "ota20k", "ota-all"}

# 通过阈值（含等于）。
PASS_THRESHOLD = 50.0

# —— 绩效页顶部过滤框选择器（实测，见 perf-probe-findings.md 补探测）——
# 方案号/目标版本/当前阶段 的 <input> 带 name 属性，直接用 input[name=...]。
# 『设备扩展信息』过滤框确实存在，FineReport 参数名 = DEVICEMSG（与填报页一致），但其
# <input class="fr-trigger-texteditor"> 本身【没有】name 属性，参数名在外层
# div[widgetname='DEVICEMSG'] 上，所以必须用 widgetname 锚定再取内层文本框；
# 若此选择器因页面结构调整而选不中、导致设备扩展信息无法过滤，需核对该选择器
# （仍以 perf-probe-findings.md 的实测 ancestors 为准）。
PERF_FILTER_INPUTS = {
    "方案号": "input[name='PLANCODE']",
    "目标版本": "input[name='NEXTVERSION']",
    "设备扩展信息": "div[widgetname='DEVICEMSG'] input.fr-trigger-texteditor",
    "当前阶段": "input[name='CURRENTSTAGE']",
}
# 清空用（避免上一次查询的残留条件干扰本次）。绩效页存在但本闸门不使用的过滤框。
PERF_FILTER_INPUTS_OTHERS = {
    "源版本": "input[name='CURRENTVERSION']",
    "大区域": "input[name='REGION']",
    "机型": "input[name='DEVICECODE']",
}

# —— 结果分页控件（实测，见 perf-probe-findings.md『结果分页机制』补探测）——
# 下一页按钮：div[widgetname='Next']（等价 #fr-btn-Next / button.x-emb-next）。
# 禁用判定：Next 的 wrapper div 带 class 'ui-state-disabled'（已翻到末页/本页即末页）。
PERF_NEXT_BUTTON = "#fr-btn-Next"
PERF_NEXT_WIDGET = "div[widgetname='Next']"
# 翻页安全上界：工具栏 '/N' 指示不等于『有数据的页数』（实测 /4 但数据只在第 1 页），
# 只拿它当『最多翻几次』的兜底上界，避免失控；真正判停靠『本页读到 0 行 / Next 禁用』。
PERF_PAGE_HARD_CAP = 200

# 绩效页表头【自己的】列名列表（不要改动 dass_fill.read_column_layout 的 known）。
# 实测该页成功率列表头文本是『升级成功率』(列字母 Z)、ruleid 列是『升级策略ruleid』(列字母 H)。
PERF_KNOWN_COLUMNS = [
    "大区域", "方案号", "方案国家", "机型", "设备扩展信息", "当前阶段",
    "当前阶段开放时间", "升级策略ruleid", "国家码", "国家", "品牌",
    "源版本", "目标版本", "目标版本类型", "升级成功率",
]

# 绩效页逻辑列名 -> 表头实际文本。
PERF_SUCCESS_RATE_HEADER = "升级成功率"
PERF_RULEID_HEADER = "升级策略ruleid"
PERF_PLAN_HEADER = "方案号"
PERF_TARGET_HEADER = "目标版本"

# 绩效页菜单叶子（实测以『安卓升级展示报表』为准；任务原称『安卓升级展示绩效』）。
PERF_MENU_LEAF_CANDIDATES = ["安卓升级展示报表", "安卓升级展示绩效"]


def navigate_to_perf(page) -> None:
    """自动导航 3 级左侧菜单打开绩效报表：
    『Android New』(展开) -> 『OTA』(展开) -> 『安卓升级展示报表』(打开报表)。

    参考 dass_login.navigate_to_fillin，仅最后一级菜单改点绩效叶子。叶子文案在不同
    环境可能是『安卓升级展示报表』或『安卓升级展示绩效』，按候选逐一尝试。
    """
    # 先展开前两级分组（状态感知：已展开就不再点，分组点击是展开/折叠切换）。
    from .dass_login import expand_menu_path
    expand_menu_path(page, ["Android New", "OTA"], PERF_MENU_LEAF_CANDIDATES)

    # 最后一级：绩效报表叶子（按候选逐一尝试，命中即停）。
    clicked = False
    for leaf in PERF_MENU_LEAF_CANDIDATES:
        print(f"\n点击菜单：{leaf} ...")
        if _click_menu_in_any_frame(page, leaf):
            print(color(f"  ✓ 已点击 {leaf}", "green"))
            clicked = True
            page.wait_for_timeout(9000)  # 报表打开较慢
            break
        else:
            print(color(f"  ⚠ 未找到菜单项：{leaf}", "yellow"))
    if not clicked:
        print(color("  ✗ 未能点中绩效报表菜单叶子，打印可见菜单供校准：", "red"))
        _dump_visible_menu(page)
        page.wait_for_timeout(2000)


def parse_percent(text) -> float | None:
    """解析形如 '28.26%' / '65.07%' / '50.00%' 的百分比文本为 float（去 %、去千分位
    逗号/空白）。无法解析（空值、无数字、缺列占位）返回 None —— 调用方按『未找到/不通过』处理。

    稳健性：去掉 '%'、半角/全角逗号、各类空白（含全角空格）后再 float()；允许前导/尾随
    杂字时用正则兜底抽取首个数字片段。解析不出数字 -> None（安全失败）。
    """
    if text is None:
        return None
    s = str(text).strip()
    if not s:
        return None
    # 去百分号、千分位分组逗号（半/全角）、所有空白（含全角空格）。
    cleaned = s.replace("%", "").replace(",", "").replace("，", "")
    cleaned = re.sub(r"\s", "", cleaned).replace("\u3000", "")
    try:
        return float(cleaned)
    except ValueError:
        pass
    # 兜底：抽取首个带小数的数字片段（如 '约 28.26 %' 之类）。
    m = re.search(r"-?\d+(?:\.\d+)?", cleaned)
    if m:
        try:
            return float(m.group(0))
        except ValueError:
            return None
    return None


def read_perf_layout(fr, debug: bool = False) -> dict:
    """解析绩效页 table.x-table 表头，返回 {letters: {列名: 列字母}, session: 会话锚}。

    与 dass_fill.read_column_layout 同构，但使用绩效页【自己的】列名列表
    PERF_KNOWN_COLUMNS（含『升级成功率』『升级策略ruleid』），不改动填报流程的 known。
    """
    result = fr.evaluate(
        """(known) => {
            const out = {letters: {}, session: null, dump: []};
            const xts = Array.from(document.querySelectorAll('table.x-table'));
            for (const t of xts) {
                for (let r = 0; r < Math.min(t.rows.length, 2); r++) {
                    for (const c of t.rows[r].cells) {
                        const txt = (c.innerText||'').trim();
                        if (!txt || !known.includes(txt)) continue;
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
        PERF_KNOWN_COLUMNS,
    )
    if debug:
        print("  绩效页表头单元格转储（文字 / id）：")
        for d in result["dump"]:
            print(f"    {d['txt']!r:20} id={d['id']!r}")
    return {"letters": result["letters"], "session": result["session"]}


def read_perf_rows(fr, letters: dict) -> list[dict]:
    """返回绩效页每个数据行的 {n, ruleid, rate_text}。

    ⚠ 以 **『升级策略ruleid』列（列字母 H）** 为行锚枚举数据行，而**不是**用『方案号』列。
    原因（实测，见 perf-probe-findings.md 第二轮补探测）：绩效页『方案号』(列字母 B) 是
    **跨行合并单元格**——同一方案号的多行结果里，方案号单元格只在首行渲染一个 `B<n>-...`，
    其余行没有 B 单元格。旧实现以方案号列非空筛『真实数据行』，会把 32 行结果误判成只有 1 行
    （甚至 0 行），这正是 HEADED 实跑读到 0 行的根因。而 ruleid 列每个结果行各有一个
    `H<行号>-0-<会话锚>` 单元格，成功率列同理有 `Z<行号>-...`，故以 ruleid 列行号为锚最可靠。

    做法：直接在整个 report frame 的 DOM 里按 id 前缀 `H<n>-` 收集所有 ruleid 数据单元格
    （n>=2 为数据行，n=1 是表头），对每个行号 n 再按 `Z<n>-` 读成功率文本。不再依赖
    『表头表 vs 数据表』的 table 结构假设（绩效页表头与数据同在一张 x-table）。
    若 ruleid/成功率列字母缺失，返回空列表（由调用方安全失败处理）。
    """
    rid_letter = letters.get(PERF_RULEID_HEADER)
    rate_letter = letters.get(PERF_SUCCESS_RATE_HEADER)
    if not rid_letter or not rate_letter:
        return []
    # 诊断用的两列（缺失时置空，不影响成功率/ruleid 的读取）。
    stage_letter = letters.get("当前阶段", "")
    open_letter = letters.get("当前阶段开放时间", "")
    rows = fr.evaluate(
        """({ridLetter, rateLetter, stageLetter, openLetter}) => {
            // 以 ruleid 列为行锚：收集所有 id 形如 '<ridLetter><n>-<x>-<session>' 且 n>=2 的单元格。
            const ridRe = new RegExp('^' + ridLetter + '(\\\\d+)-\\\\d+-');
            const seen = {};
            const res = [];
            document.querySelectorAll('[id]').forEach(e => {
                const m = (e.id || '').match(ridRe);
                if (!m) return;
                const n = parseInt(m[1]);
                if (n < 2 || seen[n]) return;        // n=1 表头；去重
                const ruleid = (e.innerText || '').trim();
                if (!ruleid) return;                 // 空 ruleid 跳过
                seen[n] = true;
                // 同行成功率单元格：id 以 '<rateLetter><n>-' 开头。
                const rc = document.querySelector('[id^="' + rateLetter + n + '-"]');
                const rate_text = rc ? (rc.innerText || '').trim() : '';
                const oc = openLetter ? document.querySelector('[id^="' + openLetter + n + '-"]') : null;
                res.push({ n: n, ruleid: ruleid, rate_text: rate_text,
                           open_text: oc ? (oc.innerText || '').trim() : '' });
            });
            res.sort((a, b) => a.n - b.n);
            // 『当前阶段』是跨行合并单元格：只在每组首行渲染，其余行沿用上方最近的那个。
            if (stageLetter) {
                const stageRe = new RegExp('^' + stageLetter + '(\\\\d+)-\\\\d+-');
                const stages = [];
                document.querySelectorAll('[id]').forEach(e => {
                    const m = (e.id || '').match(stageRe);
                    if (!m) return;
                    const k = parseInt(m[1]);
                    if (k >= 2) stages.push({ k: k, t: (e.innerText || '').trim() });
                });
                stages.sort((a, b) => a.k - b.k);
                res.forEach(r => {
                    let cur = '';
                    for (const st of stages) { if (st.k <= r.n) cur = st.t; else break; }
                    r.stage_text = cur;
                });
            }
            return res;
        }""",
        {"ridLetter": rid_letter, "rateLetter": rate_letter,
         "stageLetter": stage_letter, "openLetter": open_letter},
    )
    return rows


def _perf_set_filter(fr, page, values: dict) -> None:
    """按 字段->值 填绩效页顶部过滤框，并清空其它过滤框（避免残留条件干扰）。

    values 含本次需要的字段（方案号/目标版本/设备扩展信息/当前阶段）；PERF_FILTER_INPUTS 里
    未给值的字段与 PERF_FILTER_INPUTS_OTHERS 一律清空。设备扩展信息框的 <input> 无 name，
    用 div[widgetname='DEVICEMSG'] input.fr-trigger-texteditor 锚定（见 PERF_FILTER_INPUTS）。
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

    for field, sel in PERF_FILTER_INPUTS.items():
        _set_one(sel, values.get(field, ""))
    for sel in PERF_FILTER_INPUTS_OTHERS.values():
        _set_one(sel, "")
    shown = {k: v for k, v in values.items() if v}
    print(f"  过滤条件：{shown}")


def _perf_scroll_grid_horizontal(fr) -> None:
    """把结果网格的横向滚动容器来回滚一遍，促使 H(ruleid)/Z(成功率) 列单元格进入 DOM 后再读。

    实测（perf-probe-findings.md 第二轮）：本页 x-table 不做虚拟/懒渲染，H/Z 单元格按 id
    稳定存在、横向滚动前后计数相同。但用户明确要求『逐页横向滚动渲染 H/Z 列再读』，故这里
    做一次 best-effort 的横向滚动（滚到最右再滚回）——有懒渲染时能确保单元格入 DOM，无懒
    渲染时也无副作用（不改变单元格 id、不影响按 id 前缀读取）。任何异常都吞掉，不影响读取。
    """
    try:
        fr.evaluate(
            """() => {
                const sc = document.querySelector('div.content-container')
                    || document.querySelector('.x-table') && document.querySelector('.x-table').parentElement;
                if (!sc) return;
                const max = sc.scrollWidth - sc.clientWidth;
                if (max > 0) { sc.scrollLeft = max; sc.scrollLeft = 0; sc.scrollLeft = max; }
            }"""
        )
    except Exception:
        pass


def _perf_next_disabled(fr) -> bool:
    """『下一页』是否已禁用（已到末页）。实测禁用时 wrapper div 的 class 含 'ui-state-disabled'。"""
    try:
        cls = fr.evaluate(
            """(sel) => {
                const e = document.querySelector(sel);
                return e ? (e.className || '') : null;
            }""",
            PERF_NEXT_WIDGET,
        )
    except Exception:
        return False
    if cls is None:
        return True  # 找不到 Next 控件，当作不可继续翻页（保守停止）
    return "ui-state-disabled" in cls


def _perf_page_input_value(fr):
    """读当前页码输入框的值（实测该框无 name，用工具栏内 input.fr-texteditor 宽度小的那个）。

    返回字符串值或 None。仅用于『翻页是否完成』的变化信号，读不到不影响主流程。
    """
    try:
        return fr.evaluate(
            """() => {
                const ins = Array.from(document.querySelectorAll('input.fr-texteditor'));
                // 页码框实测 width:42px、text-align:center；取首个窄居中的文本框。
                for (const i of ins) {
                    const st = i.getAttribute('style') || '';
                    if (st.indexOf('text-align: center') >= 0) return i.value;
                }
                return ins.length ? ins[0].value : null;
            }"""
        )
    except Exception:
        return None


def _perf_loading_visible(fr) -> bool:
    """grid 是否正显示 Loading（翻页/重查中途）。读不到就当作不在 loading。"""
    try:
        return bool(
            fr.evaluate(
                """() => {
                    const els = Array.from(document.querySelectorAll('div, span, td'));
                    for (const e of els) {
                        const t = (e.innerText || '').trim();
                        if (t === 'Loading' || t === '正在加载' || t === 'Loading...') {
                            const r = e.getBoundingClientRect();
                            if (r.width > 0 && r.height > 0) return true;
                        }
                    }
                    return false;
                }"""
            )
        )
    except Exception:
        return False


def _perf_wait_page_turn(fr, page, prev_session, prev_page_val, timeout_ms: int = 15000) -> bool:
    """点『下一页』后等新页加载完成。

    判定信号（实测，见 perf-probe-findings.md『结果分页机制』第 3 节）：
      （会话锚变化 或 页码框值变化）且 Loading 不可见 -> 再 settle 500ms 视为稳定。
    返回 True 表示检测到翻页完成；超时返回 False（调用方据此停止，避免误读旧网格）。
    """
    waited, step = 0, 300
    while waited < timeout_ms:
        page.wait_for_timeout(step)
        waited += step
        if _perf_loading_visible(fr):
            continue
        cur_session = read_perf_layout(fr, debug=False)["session"]
        cur_page_val = _perf_page_input_value(fr)
        changed = (cur_session is not None and cur_session != prev_session) or (
            cur_page_val is not None and cur_page_val != prev_page_val
        )
        if changed:
            page.wait_for_timeout(500)  # settle，避免读到翻页中途的旧网格
            return True
    return False


def _perf_read_all_pages(
    fr,
    page,
    target_ruleids: set | None = None,
) -> tuple[list[dict], dict, str, int]:
    """从第 1 页起遍历所有结果页，逐页横向滚动 + 读取 + 按 norm_ruleid 去重累积。

    返回 (rows, letters, session, pages_traversed)：
      - rows：去重后的累积结果行（每行 {n, ruleid, rate_text}；n 为所在页内的行号，仅调试用）。
      - letters/session：**第 1 页** 的表头列字母映射与会话锚（供调用方校验列字母存在；后续
        页各自的 letters 内部读取时即时解析，不外泄——各页 H/Z 列字母一致）。
      - pages_traversed：实际翻过的页数（含第 1 页）。

    判停（见 findings『结果分页机制』第 5 节，不信任 '/N' 指示作数据页数）：
      1) 本页读到 0 个 ruleid(H) 单元格 -> 已越过数据末页，停；
      2) 『下一页』禁用（wrapper class 含 ui-state-disabled）-> 停；
      3) 翻页未检测到完成信号（超时）-> 停（保守，避免误读旧网格）；
      4) 命中 PERF_PAGE_HARD_CAP 安全上界 -> 停（防失控）。
    提前停止：若给了 target_ruleids 且已全部命中，则不再翻页（目标集齐即可返回）。
    去重：按 norm_ruleid 归一化后的 Rule ID 去重（防翻页重叠或合并单元格导致重复）。
    """
    # —— 第 1 页 ——
    first_layout = read_perf_layout(fr, debug=False)
    first_letters, first_session = first_layout["letters"], first_layout["session"]
    if not first_letters.get(PERF_RULEID_HEADER) or not first_letters.get(PERF_SUCCESS_RATE_HEADER):
        # 列字母缺失：直接返回空，交由调用方安全失败（与旧行为一致）。
        return [], first_letters, first_session, 1

    seen: set[str] = set()
    acc: list[dict] = []
    remaining = set(target_ruleids) if target_ruleids is not None else None

    def _accumulate(page_rows: list[dict]) -> int:
        added = 0
        for row in page_rows:
            rid = norm_ruleid(row.get("ruleid", ""))
            if not rid or rid in seen:  # 去重：空 ruleid 跳过，重复 ruleid 跳过
                continue
            seen.add(rid)
            acc.append(row)
            added += 1
            if remaining is not None:
                remaining.discard(rid)
        return added

    _perf_scroll_grid_horizontal(fr)
    page1_rows = read_perf_rows(fr, first_letters)
    _accumulate(page1_rows)
    print(f"  绩效第 1 页行数：{len(page1_rows)}（累计去重 {len(acc)}）")

    pages = 1
    # 提前停止：目标 Rule ID 已全部命中就不必再翻页。
    if remaining is not None and not remaining:
        print("  目标 Rule ID 已全部命中，提前停止翻页。")
        return acc, first_letters, first_session, pages

    prev_session = first_session
    prev_page_val = _perf_page_input_value(fr)

    while pages < PERF_PAGE_HARD_CAP:
        if _perf_next_disabled(fr):
            print(f"  『下一页』已禁用，停止翻页（共遍历 {pages} 页）。")
            break
        # 点『下一页』。
        try:
            loc = fr.locator(PERF_NEXT_BUTTON).first
            loc.scroll_into_view_if_needed(timeout=2000)
            loc.click()
        except Exception as e:
            print(f"  ⚠ 点击『下一页』失败（{PERF_NEXT_BUTTON}）：{e}；停止翻页（共 {pages} 页）。")
            break
        if not _perf_wait_page_turn(fr, page, prev_session, prev_page_val):
            print(f"  ⚠ 未检测到翻页完成信号（超时），停止翻页（共遍历 {pages} 页）。")
            break
        pages += 1

        cur_layout = read_perf_layout(fr, debug=False)
        cur_letters, cur_session = cur_layout["letters"], cur_layout["session"]
        _perf_scroll_grid_horizontal(fr)
        page_rows = read_perf_rows(fr, cur_letters) if cur_letters.get(PERF_RULEID_HEADER) else []
        added = _accumulate(page_rows)
        print(f"  绩效第 {pages} 页行数：{len(page_rows)}（本页新增去重 {added}，累计 {len(acc)}）")

        prev_session = cur_session
        prev_page_val = _perf_page_input_value(fr)

        # 判停①：本页读到 0 个 ruleid 单元格 -> 已越过数据末页。
        if len(page_rows) == 0:
            print(f"  本页无数据行，判定已越过数据末页，停止翻页（共遍历 {pages} 页）。")
            break
        # 提前停止：目标 Rule ID 已全部命中。
        if remaining is not None and not remaining:
            print(f"  目标 Rule ID 已全部命中，提前停止翻页（共遍历 {pages} 页）。")
            break
    else:
        print(f"  ⚠ 已达翻页安全上界 {PERF_PAGE_HARD_CAP} 页，停止翻页。")

    return acc, first_letters, first_session, pages


def _perf_query(
    fr,
    page,
    filters: dict,
    target_ruleids: set | None = None,
    stats_out: dict | None = None,
) -> tuple[list[dict], dict, str]:
    """填绩效页过滤框 -> 点『查询』-> 等重绘完成 -> 【跨所有页】遍历读取结果行。

    返回 (rows, letters, session)。rows 为**所有页**去重累积的结果（每行含 ruleid/rate_text），
    letters/session 为第 1 页的表头列字母映射与会话锚（供调用方校验列字母存在）。

    ⚠ 绩效页【没有『显示行数/LIMIT』过滤框】（实测，见 perf-probe-findings.md 第二轮补探测：
    顶部控件里没有 name='LIMIT' 的输入框，也没有 limit 50/100 选项）。因此这里**不设 limit**、
    也**不做**填报页那种『结果数 >= limit 疑似截断就升档重查』的逻辑——照搬填报页的
    set_display_limit(input[name='LIMIT']) 在绩效页只会超时失败且无意义。

    重绘完成信号（实测）：点『查询』后会话锚会变（如 1814->1815）、Loading 文本消失、
    方案号列切换为过滤后的方案；click_query 已显式轮询表头+会话锚恢复，这里再 settle 一下。

    翻页遍历（实测，见 perf-probe-findings.md『结果分页机制』）：查询生效后从第 1 页起，
    逐页横向滚动确保 H(ruleid)/Z(成功率) 列单元格入 DOM 再读，点『下一页』并用『会话锚/页码框
    变化且 Loading 消失』确认新页加载，逐页按 norm_ruleid 去重累积；直到本页 0 行 / Next 禁用 /
    目标 Rule ID 已集齐。**不信任工具栏 '/N' 指示作数据页数**（实测 /4 但数据仅在第 1 页）。
    target_ruleids 给定时用于提前停止（目标全部命中即不再翻页）。
    stats_out 给定时写入 {'pages': 实际遍历页数}，供调用方（如 CLI）在『未找到』提示里展示。
    """
    _perf_set_filter(fr, page, filters)
    click_query(fr, page)
    page.wait_for_timeout(1200)  # 等结果网格 settle（会话锚变化后数据单元格绘制）
    rows, letters, session, pages = _perf_read_all_pages(fr, page, target_ruleids)
    print(f"  绩效查询结果行数（已遍历 {pages} 页，去重后）：{len(rows)}")
    if stats_out is not None:
        stats_out["pages"] = pages
    return rows, letters, session


def build_gate_stage1_filters(scheme: str, target: str, option: str) -> dict:
    """构造闸门 Stage1 查询用的三字段过滤条件：方案号 + 目标版本 + 当前阶段(=上一阶段)。

    闸门只在 GATE_OPTIONS（ota200/ota1k/ota5k/ota20k/ota-all）内运行，均为递进档位，
    故 PREV_STAGE[option] 必存在（如 ota200 -> OTA50，ota-all -> OTA20000）。此函数为
    不依赖浏览器的纯逻辑，既供 check_success_rate 实际构造 Stage1 的『当前阶段』过滤值，
    也供 _selftest 离线断言『实际传给查询的三字段』，避免测试与实现漂移。
    """
    return {"方案号": scheme, "目标版本": target, "当前阶段": PREV_STAGE[option]}


def query_rates_stage1(
    fr,
    page,
    scheme: str,
    target: str,
    target_ruleids: set | None = None,
    stats_out: dict | None = None,
    current_stage: str = "",
) -> dict | None:
    """Stage1 共享查询件：方案号 + 目标版本（+可选 当前阶段）查询（跨所有页遍历），返回按
    norm_ruleid 归一化的成功率字典，供成功率闸门与独立查询 CLI 复用（不复制粘贴解析逻辑）。

    参数 current_stage：
      - 默认空串 ""：不按『当前阶段』过滤（仅方案号+目标版本两字段）。独立 CLI
        dass_perf_query 的两处调用不传该参数即保持原两字段行为，零破坏。
      - 非空：把该值并入过滤条件（filters['当前阶段']），构成『方案号+目标版本+当前阶段』
        三字段查询。闸门 Stage1 传 PREV_STAGE[option]（上一阶段）以缩小匹配范围。

    返回：
      - dict：{norm_ruleid: {"rate": float|None, "rate_text": str, "raw_ruleid": str}}，
        同一 norm_ruleid 多行时取首次出现。
      - None：绩效页未解析到『升级策略ruleid』/『升级成功率』列字母（安全失败信号），
        由调用方据此 abort / 报错。

    说明：这段逻辑原本内联在 check_success_rate 的 Stage1 循环里（_perf_query 查询 ->
    校验两列字母存在 -> 遍历结果行按 norm_ruleid 归一化 -> parse_percent 解析成功率），
    这里整体抽出为纯查询件，check_success_rate 改为调用它，行为逐行等价。

    翻页：_perf_query 现已跨所有结果页遍历并按 norm_ruleid 去重累积（见其 docstring）。
    target_ruleids 给定时作提前停止用（目标 Rule ID 全部命中即不再翻页）。
    stats_out 给定时写入 {'pages': 实际遍历页数}，供调用方展示『已遍历全部 N 页』。
    """
    filters = {"方案号": scheme, "目标版本": target}
    if current_stage:  # 仅非空才并入，空＝不按阶段过滤（CLI 不传即行为不变）
        filters["当前阶段"] = current_stage
    rows, letters, session = _perf_query(
        fr, page, filters, target_ruleids, stats_out
    )
    if not letters.get(PERF_RULEID_HEADER) or not letters.get(PERF_SUCCESS_RATE_HEADER):
        print(color(f"  ✗ 绩效页未解析到 "
                    f"『{PERF_RULEID_HEADER}』/『{PERF_SUCCESS_RATE_HEADER}』列字母"
                    f"（letters={letters}）；安全失败。", "red"))
        return None
    result: dict[str, dict] = {}
    for row in rows:
        rid = norm_ruleid(row.get("ruleid", ""))
        if not rid or rid in result:  # 同键取首次出现；空 ruleid 跳过
            continue
        rate_text = row.get("rate_text", "")
        result[rid] = {
            "rate": parse_percent(rate_text),
            "rate_text": rate_text,
            "raw_ruleid": row.get("ruleid", ""),
            "stage_text": row.get("stage_text", ""),   # 诊断用：报表里该规则的『当前阶段』
            "open_text": row.get("open_text", ""),     # 诊断用：『当前阶段开放时间』
        }
    return result


def read_report_date(fr) -> str:
    """读绩效页『日期』过滤框的值（YYYY-MM-DD；该框没有 name，按值的格式找）。读不到返回空串。"""
    try:
        return fr.evaluate("""() => {
            const el = [...document.querySelectorAll('input')]
                .find(i => /^\\d{4}-\\d{2}-\\d{2}$/.test((i.value || '').trim()));
            return el ? el.value.trim() : '';
        }""") or ""
    except Exception:
        return ""


def format_stage_diagnosis(missing: list[str], seen: dict, want_stage: str, report_date: str = "") -> list[str]:
    """把『闸门按阶段过滤查不到』的规则，对照『不按阶段查』时报表里的实际情况，生成诊断行（纯逻辑）。

    seen: {norm_ruleid: {"stage_text","open_text","rate_text"}} —— 不带阶段条件时报表里的行。
    只做说明，不改变闸门放行/中止的判定。
    """
    lines: list[str] = []
    if not missing:
        return lines
    in_report = [r for r in missing if r in seen]
    not_in = [r for r in missing if r not in seen]
    head = f"\n======== 未找到原因诊断（报表数据日期：{report_date or '未知'}）========"
    lines.append(color(head, "yellow"))
    if in_report:
        lines.append(color(
            f"  {len(in_report)} 个 Rule ID 在报表里【存在】，但『当前阶段』不是闸门要找的 {want_stage!r}：", "yellow"))
        for rid in sorted(in_report, key=lambda x: int(x) if x.isdigit() else 0):
            d = seen[rid]
            lines.append(f"    Rule ID={rid!r}  报表当前阶段={d.get('stage_text') or '?'!r}  "
                         f"开放时间={d.get('open_text') or '?'}  成功率={d.get('rate_text') or '?'}")
        lines.append(color(
            "  可能原因：展示报表的数据按日期快照、滞后于『安卓升级状态填报』——填报里已是上一阶段，"
            "报表还停在更早的阶段。此时报表里的成功率属于更早阶段的人群，不能代表上一阶段，"
            "所以仍按『未找到』中止；等报表同步（日期推进到阶段变更之后）再重试。", "yellow"))
    if not_in:
        lines.append(color(
            f"  {len(not_in)} 个 Rule ID 连不带阶段条件也查不到：{sorted(not_in)}。"
            "请核对 方案号 / 目标版本 / 设备扩展信息 是否与 DASS 一致。", "yellow"))
    return lines


def _diagnose_missing(fr, page, missing: list[str], excel_by_ruleid: dict, want_stage: str) -> dict:
    """只读诊断：对未找到的规则，按『方案号+目标版本』（不带阶段）再查一次，打印报表里的实际阶段。

    返回 {norm_ruleid: {stage_text, open_text, rate_text, rate}}（报表里实际查到的行，供逐条明细作参考值）。"""
    seen: dict[str, dict] = {}
    groups = sorted({(excel_by_ruleid[r].get("方案号", ""), excel_by_ruleid[r].get("目标版本", "")) for r in missing})
    print(color("\n  （只读诊断：不带『当前阶段』条件再查一次，看这些规则在报表里是什么情况 ...）", "dim"))
    for scheme, target in groups:
        res = query_rates_stage1(fr, page, scheme, target, set(missing), current_stage="")
        if res:
            seen.update(res)
    report_date = read_report_date(fr)
    for line in format_stage_diagnosis(missing, seen, want_stage, report_date):
        print(line)
    return seen


def check_success_rate(fr, page, records, option: str) -> bool:
    """成功率闸门：对本批每个 Rule ID 求成功率，全部 >= 50% 返回 True（放行升级），
    任一 < 50% 或两阶段都查不到 -> 打印逐条明细后返回 False（abort）。

    - option 不在 GATE_OPTIONS（即 ota50/init）时不应校验；此处若被误调用，直接放行 True
      并打印提示（真正的跳过在 dass_fill.main() 里判断，不进入本函数）。
    - 两阶段：Stage1 方案号+目标版本 模糊查询；Stage2 对未命中的 Rule ID 用四字段精准查询
      方案号+目标版本+设备扩展信息+当前阶段(上一阶段) 查询再本地按 norm_ruleid 定位。
    - 成功率/ruleid 列字母缺失、或成功率文本无法解析 -> 该 Rule ID 记为『未找到/不通过』并 abort。
    """
    if option not in GATE_OPTIONS:
        print(f"  （option={option!r} 不在成功率闸门触发集合 {sorted(GATE_OPTIONS)}，不校验，直接放行。）")
        return True

    # —— Excel 侧：按 norm_ruleid 归一化的 Rule ID 索引（取首条记录作查询条件）——
    excel_by_ruleid: dict[str, dict] = {}
    for rec in records:
        rid = norm_ruleid(rec.get("ruleid", ""))
        if not rid:
            print(color(f"  ⚠ 跳过 ruleid 为空的 Excel 记录：方案号={rec.get('方案号', '')!r}", "yellow"))
            continue
        excel_by_ruleid.setdefault(rid, rec)
    if not excel_by_ruleid:
        print(color("  ✗ Excel 中没有可用的 Rule ID（全为空），无法校验成功率，abort 升级。", "red"))
        return False

    excel_ruleids = set(excel_by_ruleid)
    excel_scheme_nos = {rec.get("方案号", "") for rec in records if rec.get("方案号", "")}
    excel_targets = {rec.get("目标版本", "") for rec in records if rec.get("目标版本", "")}

    # rid -> {"rate": float|None, "stage": "Stage1"/"Stage2"/"未找到", "rate_text": str}
    found_rate: dict[str, dict] = {}

    # ———— Stage1：方案号 × 目标版本 模糊查询 ————
    target_values = sorted(excel_targets) if excel_targets else [""]
    for scheme in sorted(excel_scheme_nos):
        for target in target_values:
            # A：闸门 Stage1 用三字段查询——方案号 + 目标版本 + 当前阶段(=上一阶段 PREV_STAGE)。
            gate_filters = build_gate_stage1_filters(scheme, target, option)
            prev_stage = gate_filters["当前阶段"]
            print(f"\n====== 成功率 Stage1 粗筛查询：方案号={scheme!r} 目标版本={target!r} "
                  f"当前阶段={prev_stage!r}（上一阶段）======")
            # 提前停止：本批目标 Rule ID 全部命中即可停翻页（跨页遍历在 _perf_query 内完成）。
            stage1 = query_rates_stage1(
                fr, page, scheme, target, excel_ruleids, current_stage=prev_stage
            )
            if stage1 is None:  # 列字母缺失 -> 安全失败 abort（行为同旧内联分支）
                _shot(page, "perf_missing_columns")
                return False
            for rid, info in stage1.items():
                if rid in excel_ruleids and rid not in found_rate:
                    found_rate[rid] = {
                        "rate": info["rate"],
                        "stage": "Stage1",
                        "rate_text": info["rate_text"],
                    }

    # ———— Stage2：对 Stage1 未命中的 Rule ID 逐个精准查询 ————
    for rid in sorted(excel_ruleids - set(found_rate)):
        rec = excel_by_ruleid[rid]
        k_scheme = rec.get("方案号", "")
        k_target = rec.get("目标版本", "")
        k_ext = rec.get("设备扩展信息", "")  # 取自该 Rule ID 的 Excel 记录，与 run_update Stage2 一致
        # 用户最初要求的四字段精准查询：方案号 + 目标版本 + 设备扩展信息 + 当前阶段(上一阶段)。
        filters = {"方案号": k_scheme, "目标版本": k_target, "设备扩展信息": k_ext}
        if option != "ota50":  # GATE_OPTIONS 恒为递进档位，必有上一阶段
            filters["当前阶段"] = PREV_STAGE[option]
        print(f"\n====== 成功率 Stage2 精准查询：Rule ID={rid!r} 方案号={k_scheme!r} "
              f"目标版本={k_target!r} 设备扩展信息={k_ext!r} 当前阶段={filters.get('当前阶段', '')!r} ======")
        # 跨页遍历（_perf_query 内）；只找这一个 rid，命中即提前停止翻页。
        rows, letters, session = _perf_query(fr, page, filters, {rid})
        if not letters.get(PERF_RULEID_HEADER) or not letters.get(PERF_SUCCESS_RATE_HEADER):
            print(color(f"  ✗ 绩效页未解析到 "
                        f"『{PERF_RULEID_HEADER}』/『{PERF_SUCCESS_RATE_HEADER}』列字母"
                        f"（letters={letters}）；安全失败，abort 升级。", "red"))
            _shot(page, "perf_missing_columns")
            return False
        matches = [r for r in rows if norm_ruleid(r.get("ruleid", "")) == rid]
        if matches:
            rate = parse_percent(matches[0].get("rate_text", ""))
            found_rate[rid] = {
                "rate": rate,
                "stage": "Stage2",
                "rate_text": matches[0].get("rate_text", ""),
            }
        else:
            found_rate[rid] = {"rate": None, "stage": "未找到", "rate_text": ""}

    _shot(page, "perf_queried")

    # ———— 诊断（只读，不影响判定）：按阶段查不到的规则，不带阶段条件再看一眼，说明原因 ————
    not_found = [r for r in sorted(excel_ruleids) if found_rate.get(r, {}).get("stage") == "未找到"]
    if not_found:
        try:
            seen = _diagnose_missing(fr, page, not_found, excel_by_ruleid, PREV_STAGE.get(option, ""))
            # 把报表里实际查到的数值挂到该规则上，逐条明细里作为『仅供参考』显示（不改变判定）。
            for rid, d in (seen or {}).items():
                if rid in found_rate and found_rate[rid].get("stage") == "未找到":
                    found_rate[rid]["fallback"] = d
        except Exception as e:
            print(color(f"  （未找到原因诊断跳过：{e}）", "dim"))

    # ———— 判定 + 逐条/汇总打印（B）————
    report, passed = format_gate_report(excel_ruleids, found_rate)
    print(report)
    if not passed:
        _shot(page, "perf_gate_abort")
    return passed


def format_gate_report(excel_ruleids, found_rate: dict) -> tuple[str, bool]:
    """把闸门判定结果格式化成『逐条 + 汇总 + 结论』文本，并返回 (文本, 是否放行)。

    为不依赖浏览器的纯逻辑（供 check_success_rate 打印，也供 _selftest 离线断言）。

    - 逐条：对 sorted(excel_ruleids) 每个 Rule ID 打印一行，无论通过与否都打印，含
      Rule ID、升级成功率（命中可解析 -> '49.09%'；未找到 -> '未找到'；命中但解析失败 ->
      '成功率解析失败（原文 ...）'）、是否 >=50%（'通过 ✅' / '不通过 ❌'）、来源阶段
      （Stage1 / Stage2 / 未找到）。
    - 汇总：总数 / 通过数 / 不通过数 / 未找到数。
    - 结论：无 failures -> 『放行 ✅』；有 failures -> 『abort ❌』。

    判定（与原行为一致）：rate 为 None（未找到/解析失败）或 rate < PASS_THRESHOLD 记为失败。
    """
    rids = sorted(excel_ruleids)
    lines: list[str] = []
    lines.append("\n======== 成功率闸门：逐条明细 ========")

    pass_n = 0
    fail_n = 0
    not_found_n = 0
    failures: list[str] = []

    for rid in rids:
        info = found_rate.get(rid, {"rate": None, "stage": "未找到", "rate_text": ""})
        rate = info.get("rate")
        stage = info.get("stage", "未找到")
        if rate is None:
            if stage == "未找到":
                fb = info.get("fallback")
                if fb and fb.get("rate_text"):
                    rate_desc = color(
                        f"按要求的阶段未找到；报表里该规则：阶段={fb.get('stage_text') or '?'}，"
                        f"成功率={fb.get('rate_text')}（仅供参考，不算通过）", "yellow")
                else:
                    rate_desc = color("未找到", "yellow")
                not_found_n += 1
            else:
                rate_desc = color(f"成功率解析失败（原文 {info.get('rate_text', '')!r}）", "red")
                fail_n += 1
            verdict = color("不通过 ❌", "red")
            failures.append(rid)
        else:
            rate_desc = f"升级成功率={rate:.2f}%"
            if rate >= PASS_THRESHOLD:
                verdict = color("通过 ✅", "green")
                pass_n += 1
            else:
                verdict = color("不通过 ❌", "red")
                failures.append(rid)
                fail_n += 1
        lines.append(f"  Rule ID={rid!r}  {rate_desc}  {verdict}  来源阶段={stage}")

    # 不通过数 = 命中但未达标/解析失败；未找到数单列。总失败 = 不通过 + 未找到。
    lines.append("\n======== 成功率闸门：汇总 ========")
    lines.append(f"  总数：{len(rids)}  通过：{pass_n}  不通过：{fail_n}  未找到：{not_found_n}")

    if not failures:
        lines.append(color("\n======== 最终结论：放行 ✅ ========", "green"))
        lines.append(color(
            f"  本批 {len(rids)} 个 Rule ID 成功率均 >= {PASS_THRESHOLD:.0f}%，放行升级。", "green"))
        return "\n".join(lines), True

    lines.append(color("\n======== 最终结论：abort ❌ 中止升级 ========", "red"))
    lines.append(color(
        f"  以下 {len(failures)} 个 Rule ID 未达标"
        f"（< {PASS_THRESHOLD:.0f}% 或两阶段都查不到）：{failures}", "red"))
    lines.append(color("  run_update 不执行，也不进入人工提交环节。", "red"))
    return "\n".join(lines), False


# —— 不依赖浏览器的纯函数自测（python -m automation.dass_perf --selftest）——
def _selftest_diag(check) -> None:
    seen = {"65235": {"stage_text": "OTA200", "open_text": "20261107", "rate_text": "0.00%"}}
    out = "\n".join(format_stage_diagnosis(["65235", "99999"], seen, "OTA1000", "2026-11-08"))
    check("诊断：报表阶段与闸门阶段不一致被点出", "OTA200" in out and "OTA1000" in out and "2026-11-08" in out)
    check("诊断：不带阶段也查不到的单独列出", "99999" in out and "连不带阶段条件也查不到" in out)
    check("诊断：无缺失时无输出", format_stage_diagnosis([], seen, "OTA1000") == [])
    rep, passed = format_gate_report({"65235", "65236"}, {
        "65235": {"rate": None, "stage": "未找到", "rate_text": "",
                  "fallback": {"stage_text": "OTA200", "rate_text": "38.43%"}},
        "65236": {"rate": 80.0, "stage": "Stage1", "rate_text": "80.00%"},
    })
    check("逐条明细：每个 ID 都显示成功率（含未找到时的报表参考值）",
          "38.43%" in rep and "仅供参考" in rep and "80.00%" in rep and passed is False)


def _selftest() -> int:
    ok = True

    def check(name, cond):
        nonlocal ok
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        if not cond:
            ok = False

    _selftest_diag(check)

    # parse_percent
    check("50.00% -> 50.0", parse_percent("50.00%") == 50.0)
    check("49.30% -> 49.3", parse_percent("49.30%") == 49.3)
    check("28.26% -> 28.26", parse_percent("28.26%") == 28.26)
    check("千分位 '1,234.50%' -> 1234.5", parse_percent("1,234.50%") == 1234.5)
    check("全角逗号 '1，234.50%' -> 1234.5", parse_percent("1，234.50%") == 1234.5)
    check("带空白 ' 65.07 % ' -> 65.07", parse_percent(" 65.07 % ") == 65.07)
    check("空值 '' -> None", parse_percent("") is None)
    check("None -> None", parse_percent(None) is None)
    check("非数字 '--' -> None", parse_percent("--") is None)

    # 阈值判定（含等于放行）
    check("50.00% 放行 (>=50)", parse_percent("50.00%") >= PASS_THRESHOLD)
    check("49.30% 不放行", not (parse_percent("49.30%") >= PASS_THRESHOLD))

    # norm_ruleid 对齐
    check("'65,196' 与 '65196' 对齐", norm_ruleid("65,196") == norm_ruleid("65196"))
    check("'65，196'(全角) 与 '65196' 对齐", norm_ruleid("65，196") == norm_ruleid("65196"))
    check("'65 196'(空格) 与 '65196' 对齐", norm_ruleid("65 196") == norm_ruleid("65196"))

    # 触发集合
    check("ota50 不触发闸门", "ota50" not in GATE_OPTIONS)
    check("init 不触发闸门", "init" not in GATE_OPTIONS)
    check("ota200 触发闸门", "ota200" in GATE_OPTIONS)
    check("ota-all 触发闸门", "ota-all" in GATE_OPTIONS)

    # 查不到的 Rule ID 触发 abort 的判定逻辑（rate=None => 失败）
    none_rate = None
    check("成功率 None 判为不通过", not (none_rate is not None and none_rate >= PASS_THRESHOLD))

    # —— A：闸门 Stage1 三字段 filters 构造 ——
    f200 = build_gate_stage1_filters("P", "T", "ota200")
    check("Stage1 filters 含方案号/目标版本/当前阶段三键",
          set(f200) == {"方案号", "目标版本", "当前阶段"})
    check("Stage1 filters 方案号/目标版本原样", f200["方案号"] == "P" and f200["目标版本"] == "T")
    check("ota200 当前阶段=上一阶段 OTA50", f200["当前阶段"] == PREV_STAGE["ota200"] == "OTA50")
    check("ota-all 当前阶段=上一阶段 OTA20000",
          build_gate_stage1_filters("P", "T", "ota-all")["当前阶段"] == "OTA20000")

    # —— A：query_rates_stage1 的 current_stage 默认空＝不改变 CLI 两字段行为 ——
    import inspect
    sig = inspect.signature(query_rates_stage1)
    check("query_rates_stage1 新增 current_stage 且默认空串",
          "current_stage" in sig.parameters and sig.parameters["current_stage"].default == "")

    # —— B：闸门逐条/汇总/结论日志格式（含通过/不通过/未找到/解析失败四类样本）——
    excel_ids = {"101", "102", "103", "104"}
    found = {
        "101": {"rate": 72.5, "stage": "Stage1", "rate_text": "72.50%"},   # 通过
        "102": {"rate": 49.09, "stage": "Stage2", "rate_text": "49.09%"},  # 不通过
        "103": {"rate": None, "stage": "未找到", "rate_text": ""},          # 未找到
        "104": {"rate": None, "stage": "Stage1", "rate_text": "--"},        # 解析失败
    }
    report, passed = format_gate_report(excel_ids, found)
    check("有 failures 时不放行", passed is False)
    check("逐条含每个 Rule ID", all(f"Rule ID={rid!r}" in report for rid in excel_ids))
    check("逐条含『通过 ✅』", "通过 ✅" in report)
    check("逐条含『不通过 ❌』", "不通过 ❌" in report)
    check("逐条含『未找到』", "未找到" in report)
    check("逐条含『成功率解析失败』", "成功率解析失败" in report)
    check("逐条含来源阶段 Stage1/Stage2", "来源阶段=Stage1" in report and "来源阶段=Stage2" in report)
    check("成功率按 49.09% 原样格式化", "49.09%" in report)
    # 样本：101 通过 / 102 <50%(不通过) / 104 解析失败(不通过) / 103 未找到
    check("汇总含总数/通过/不通过/未找到四个计数",
          "总数：4" in report and "通过：1" in report and "不通过：2" in report and "未找到：1" in report)
    check("结论行含 abort", "abort" in report)

    # 全部通过 -> 放行
    found_pass = {
        "201": {"rate": 50.0, "stage": "Stage1", "rate_text": "50.00%"},  # 等于阈值放行
        "202": {"rate": 88.0, "stage": "Stage2", "rate_text": "88.00%"},
    }
    report2, passed2 = format_gate_report({"201", "202"}, found_pass)
    check("全部 >=50% 放行", passed2 is True)
    check("放行结论含『放行』", "放行" in report2)
    check("放行汇总 通过：2 不通过：0 未找到：0",
          "通过：2" in report2 and "不通过：0" in report2 and "未找到：0" in report2)

    print("\n自测结果：", "全部通过 ✅" if ok else "存在失败 ❌")
    return 0 if ok else 1


if __name__ == "__main__":
    import sys
    if "--selftest" in sys.argv[1:]:
        sys.exit(_selftest())
    print("此模块通常由 dass_fill 在升级前调用。纯函数自测：python -m automation.dass_perf --selftest")
