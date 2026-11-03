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
    # 先展开前两级分组。
    for text, wait_ms in (("Android New", 1500), ("OTA", 1500)):
        print(f"\n点击菜单：{text} ...")
        if _click_menu_in_any_frame(page, text):
            print(f"  ✓ 已点击 {text}")
        else:
            print(f"  ⚠ 未找到菜单项：{text}")
            _dump_visible_menu(page)
        page.wait_for_timeout(wait_ms)

    # 最后一级：绩效报表叶子（按候选逐一尝试，命中即停）。
    clicked = False
    for leaf in PERF_MENU_LEAF_CANDIDATES:
        print(f"\n点击菜单：{leaf} ...")
        if _click_menu_in_any_frame(page, leaf):
            print(f"  ✓ 已点击 {leaf}")
            clicked = True
            page.wait_for_timeout(9000)  # 报表打开较慢
            break
        else:
            print(f"  ⚠ 未找到菜单项：{leaf}")
    if not clicked:
        print("  ✗ 未能点中绩效报表菜单叶子，打印可见菜单供校准：")
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
    rows = fr.evaluate(
        """({ridLetter, rateLetter}) => {
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
                res.push({ n: n, ruleid: ruleid, rate_text: rate_text });
            });
            res.sort((a, b) => a.n - b.n);
            return res;
        }""",
        {"ridLetter": rid_letter, "rateLetter": rate_letter},
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


def _perf_query(fr, page, filters: dict) -> tuple[list[dict], dict, str]:
    """填绩效页过滤框 -> 点『查询』-> 等重绘完成 -> 读结果行。

    返回 (rows, letters, session)。rows 来自 read_perf_rows（每行含 ruleid/rate_text）。

    ⚠ 绩效页【没有『显示行数/LIMIT』过滤框】（实测，见 perf-probe-findings.md 第二轮补探测：
    顶部控件里没有 name='LIMIT' 的输入框，也没有 limit 50/100 选项）。因此这里**不设 limit**、
    也**不做**填报页那种『结果数 >= limit 疑似截断就升档重查』的逻辑——照搬填报页的
    set_display_limit(input[name='LIMIT']) 在绩效页只会超时失败且无意义。查询流程就是：
    填方案号+目标版本 -> 点查询 -> 等重绘 -> 读结果。

    重绘完成信号（实测）：点『查询』后会话锚会变（如 1814->1815）、Loading 文本消失、
    方案号列切换为过滤后的方案；click_query 已显式轮询表头+会话锚恢复，这里再 settle 一下。
    成功率列(Z)/ruleid 列(H) 的单元格在查询后即渲染到 DOM（实测无需额外横向滚动即可读到
    H/Z 单元格——它们是独立的单元格 id，不受视口影响；本函数按 id 前缀直接读取）。
    """
    _perf_set_filter(fr, page, filters)
    click_query(fr, page)
    page.wait_for_timeout(1200)  # 等结果网格 settle（会话锚变化后数据单元格绘制）
    layout = read_perf_layout(fr, debug=False)
    letters, session = layout["letters"], layout["session"]
    rows = read_perf_rows(fr, letters)
    print(f"  绩效查询结果行数：{len(rows)}")
    return rows, letters, session


def query_rates_stage1(fr, page, scheme: str, target: str) -> dict | None:
    """Stage1 共享查询件：方案号 + 目标版本 模糊查询（limit 100），返回按 norm_ruleid
    归一化的成功率字典，供成功率闸门与独立查询 CLI 复用（不复制粘贴解析逻辑）。

    返回：
      - dict：{norm_ruleid: {"rate": float|None, "rate_text": str, "raw_ruleid": str}}，
        同一 norm_ruleid 多行时取首次出现。
      - None：绩效页未解析到『升级策略ruleid』/『升级成功率』列字母（安全失败信号），
        由调用方据此 abort / 报错。

    说明：这段逻辑原本内联在 check_success_rate 的 Stage1 循环里（_perf_query 查询 ->
    校验两列字母存在 -> 遍历结果行按 norm_ruleid 归一化 -> parse_percent 解析成功率），
    这里整体抽出为纯查询件，check_success_rate 改为调用它，行为逐行等价。
    """
    rows, letters, session = _perf_query(fr, page, {"方案号": scheme, "目标版本": target})
    if not letters.get(PERF_RULEID_HEADER) or not letters.get(PERF_SUCCESS_RATE_HEADER):
        print(f"  ✗ 绩效页未解析到 "
              f"『{PERF_RULEID_HEADER}』/『{PERF_SUCCESS_RATE_HEADER}』列字母"
              f"（letters={letters}）；安全失败。")
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
        }
    return result


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
            print(f"  ⚠ 跳过 ruleid 为空的 Excel 记录：方案号={rec.get('方案号', '')!r}")
            continue
        excel_by_ruleid.setdefault(rid, rec)
    if not excel_by_ruleid:
        print("  ✗ Excel 中没有可用的 Rule ID（全为空），无法校验成功率，abort 升级。")
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
            print(f"\n====== 成功率 Stage1 粗筛查询：方案号={scheme!r} 目标版本={target!r} ======")
            stage1 = query_rates_stage1(fr, page, scheme, target)
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
        rows, letters, session = _perf_query(fr, page, filters)
        if not letters.get(PERF_RULEID_HEADER) or not letters.get(PERF_SUCCESS_RATE_HEADER):
            print(f"  ✗ 绩效页未解析到 "
                  f"『{PERF_RULEID_HEADER}』/『{PERF_SUCCESS_RATE_HEADER}』列字母"
                  f"（letters={letters}）；安全失败，abort 升级。")
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

    # ———— 判定：全部 >= 50% 才放行 ————
    failures: list[tuple[str, dict]] = []
    for rid in sorted(excel_ruleids):
        info = found_rate.get(rid, {"rate": None, "stage": "未找到", "rate_text": ""})
        rate = info["rate"]
        if rate is None or rate < PASS_THRESHOLD:
            failures.append((rid, info))

    if not failures:
        print("\n======== 成功率闸门：通过 ✅ ========")
        print(f"  本批 {len(excel_ruleids)} 个 Rule ID 成功率均 >= {PASS_THRESHOLD:.0f}%，放行升级。")
        for rid in sorted(excel_ruleids):
            info = found_rate[rid]
            print(f"    Rule ID={rid!r}  成功率={info['rate']:.2f}%  （{info['stage']}）")
        return True

    print("\n======== 成功率闸门：不通过 ❌ abort 升级 ========")
    print(f"  以下 {len(failures)} 个 Rule ID 未达标（< {PASS_THRESHOLD:.0f}% 或两阶段都查不到），"
          "run_update 不执行，也不进入人工提交环节：")
    for rid, info in failures:
        if info["rate"] is None:
            detail = "未找到" if info["stage"] == "未找到" else f"成功率解析失败（原文 {info['rate_text']!r}）"
            print(f"    Rule ID={rid!r}  {detail}  （{info['stage']}）")
        else:
            print(f"    Rule ID={rid!r}  成功率={info['rate']:.2f}% < {PASS_THRESHOLD:.0f}%  （{info['stage']}）")
    _shot(page, "perf_gate_abort")
    return False


# —— 不依赖浏览器的纯函数自测（python -m automation.dass_perf --selftest）——
def _selftest() -> int:
    ok = True

    def check(name, cond):
        nonlocal ok
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        if not cond:
            ok = False

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

    print("\n自测结果：", "全部通过 ✅" if ok else "存在失败 ❌")
    return 0 if ok else 1


if __name__ == "__main__":
    import sys
    if "--selftest" in sys.argv[1:]:
        sys.exit(_selftest())
    print("此模块通常由 dass_fill 在升级前调用。纯函数自测：python -m automation.dass_perf --selftest")
