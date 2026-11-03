"""只读探测：『安卓升级展示绩效』页（升级成功率闸门的数据来源）。

运行（项目根目录）：
    cd ota_cms_automation
    ./.venv/bin/python -m automation.dass_probe_perf

背景（见 .agents/tasks/probe-findings.md、dass_login.py、dass_fill.py）：
- 『安卓升级展示绩效』与『安卓升级状态填报』同级（左侧菜单 Android New -> OTA 下），
  都是 FineReport 报表，结构类似：报表在 iframe 内、顶部红色过滤框、table.x-table 表头。
- 后续要新增『升级成功率闸门』：在升级 DASS 之前，用 方案号+目标版本（模糊）查询绩效页，
  按 ruleid 对比成功率是否 >=50%，全部满足才放行，否则 abort。
- 本步【只读探测】：确认绩效页的 frame、过滤框 name、表头列名->列字母映射、成功率列与
  ruleid 列的实际表头文本与单元格 id 格式、一两行示例数据文本。不实现闸门逻辑、不改业务代码。

只读安全：本脚本只登录、导航、点一次『查询』、读 DOM。绝不插入行、绝不填任何单元格、
绝不点提交/校验。所有 JS prompt 一律 dismiss。
"""
from __future__ import annotations

import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

from . import config
from .dass_login import login, _shot, _click_menu_in_any_frame, _dump_visible_menu
from .dass_fill import (
    _report_frame,
    read_column_layout,
    set_display_limit,
    set_filter,
    click_query,
    LIMIT_START,
    FILTER_INPUTS,
    FILTER_INPUTS_OTHERS,
    LIMIT_INPUT,
)

FINDINGS_PATH = Path(__file__).resolve().parents[1] / ".agents" / "tasks" / "perf-probe-findings.md"

# 绩效页叶子菜单名（与『安卓升级状态填报』同级）。
# 任务指定名为『安卓升级展示绩效』，但实测左侧 OTA 菜单里**没有**这个确切名字，
# 与填报页同级、紧挨其下的是『安卓升级展示报表』（见探测截图）。因此按任务名优先、
# 随后用实测候选名回退逐个尝试，并记录“真正打开了报表的那个叶子名”。
# 可用环境变量 PERF_MENU 覆盖候选首选（不改代码即可校准）。
PERF_MENU_CANDIDATES = [
    os.getenv("PERF_MENU", "").strip() or "安卓升级展示绩效",
    "安卓升级展示报表",
]
# 去重保序。
PERF_MENU_CANDIDATES = list(dict.fromkeys([c for c in PERF_MENU_CANDIDATES if c]))

# 成功率列的候选表头文本（探测时用来在表头里找它；实际文本以 dump 为准）。
# 实测该列表头为『升级成功率』（列字母 Z）；任务原假设的『升级成功率（11）』其实是
# 另一列『升级成功设备数（11）』（设备数，列字母 Y），不是成功率列。保留几个写法做兜底。
SUCCESS_RATE_CANDIDATES = [
    "升级成功率", "升级成功率（11）", "升级成功率(11)", "成功率",
]
# ruleid 列：实测表头为『升级策略ruleid』（列字母 H），不是填报页的裸 'ruleid'。
RULEID_CANDIDATES = ["升级策略ruleid", "ruleid"]


def _report_frame_loaded(page) -> bool:
    """报表是否真的打开了：存在 URL 含 entry_report 的 frame。"""
    return any("entry_report" in (fr.url or "") for fr in page.frames)


def navigate_to_perf(page) -> str | None:
    """自动导航左侧 3 级菜单：Android New -> OTA -> <绩效/展示报表叶子>。

    照搬 dass_login.navigate_to_fillin 的逐级点击（_click_menu_in_any_frame），
    先展开 Android New -> OTA，再对 PERF_MENU_CANDIDATES 逐个尝试点击叶子，
    点中且报表 frame 真正加载（URL 含 entry_report）即停止。

    返回真正打开报表的叶子名；都没打开则返回 None（调用方据此判为未实跑）。"""
    for text, wait_ms in (("Android New", 1500), ("OTA", 1500)):
        print(f"\n点击菜单：{text} ...")
        if _click_menu_in_any_frame(page, text):
            print(f"  ✓ 已点击 {text}")
        else:
            print(f"  ⚠ 未找到菜单项：{text}")
        page.wait_for_timeout(wait_ms)

    for leaf in PERF_MENU_CANDIDATES:
        print(f"\n尝试打开叶子报表：{leaf} ...")
        if not _click_menu_in_any_frame(page, leaf):
            print(f"  ⚠ 未找到菜单项：{leaf}")
            continue
        print(f"  ✓ 已点击 {leaf}，等待报表加载 ...")
        page.wait_for_timeout(9000)
        if _report_frame_loaded(page):
            print(f"  ✓ 报表已加载（entry_report frame 出现），采用叶子：{leaf!r}")
            return leaf
        print(f"  ⚠ 点击 {leaf!r} 后未见 entry_report 报表 frame，尝试下一个候选。")
    print("  ✗ 所有候选叶子都未能打开报表。")
    _dump_visible_menu(page)
    return None


def dump_filter_inputs(fr) -> list[dict]:
    """列出绩效页顶部所有过滤框的 name/type/placeholder/value + 边界框坐标（供确认选择器）。

    重点确认：当前阶段过滤框 name 是否仍是 CURRENTSTAGE，以及 PLANCODE / NEXTVERSION /
    DEVICEMSG 是否与填报页一致。额外返回每个 input 的 top/left/right 坐标，供标签就近匹配。"""
    return fr.evaluate(
        """() => {
            const out = [];
            const els = document.querySelectorAll("input, textarea, select");
            for (const e of els) {
                const r = e.getBoundingClientRect();
                out.push({
                    tag: e.tagName.toLowerCase(),
                    name: e.getAttribute('name') || '',
                    type: e.getAttribute('type') || '',
                    placeholder: e.getAttribute('placeholder') || '',
                    value: (e.value || '').slice(0, 20),
                    visible: r.width > 0 && r.height > 0,
                    top: Math.round(r.top),
                    left: Math.round(r.left),
                    right: Math.round(r.right),
                    bottom: Math.round(r.bottom),
                    width: Math.round(r.width),
                    height: Math.round(r.height),
                });
            }
            return out;
        }"""
    )


def dump_filter_label_map(fr) -> dict:
    """把顶部每个可见过滤 input 与其对应的『标签文字』就近配对。

    FineReport 的过滤栏通常是：一段标签文本（如『设备扩展信息』），其右侧紧跟一个 input。
    本函数几何配对：对每个 input，在同一水平带（纵向重叠）里找它【左侧最近】的文本节点作为
    标签；若左侧没有，再找【正上方最近】的文本节点兜底。这样即便 input 的 name 不直观
    （如 DEVICEMSG / 某个随机 name），也能据实知道它屏幕上对应哪个中文标签。

    返回：
      {
        "pairs": [ {name, type, value, left, right, top, label, label_left, label_top,
                    match: 'left'|'above'|'none', dist}... ],   # 每个可见 input 一条
        "device_msg": {label, name, ...} | None,   # 标签含『设备扩展信息』的那个 input
        "all_label_texts": [ ... ],                 # 过滤栏里出现过的候选标签文本（去重）
      }
    """
    return fr.evaluate(
        r"""() => {
            // 1) 收集所有可见、可输入的过滤 input（排除隐藏/0 尺寸）。
            const inputs = [];
            for (const e of document.querySelectorAll("input, textarea, select")) {
                const r = e.getBoundingClientRect();
                if (r.width <= 0 || r.height <= 0) continue;
                inputs.push({
                    el: e,
                    name: e.getAttribute('name') || '',
                    type: e.getAttribute('type') || '',
                    value: (e.value || '').slice(0, 20),
                    left: Math.round(r.left), right: Math.round(r.right),
                    top: Math.round(r.top), bottom: Math.round(r.bottom),
                    cx: (r.left + r.right) / 2, cy: (r.top + r.bottom) / 2,
                });
            }

            // 2) 收集所有“叶子文本节点”的屏幕位置（含中文标签）。只取可见、非纯空白、
            //    长度 <= 20 的短文本（标签都很短），排除 input 自身。
            const labels = [];
            const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null);
            let node;
            while ((node = walker.nextNode())) {
                const txt = (node.nodeValue || '').trim();
                if (!txt || txt.length > 20) continue;
                // 必须含中文或可见字符，排除纯符号。
                if (!/[\u4e00-\u9fa5A-Za-z]/.test(txt)) continue;
                const range = document.createRange();
                range.selectNodeContents(node);
                const r = range.getBoundingClientRect();
                if (r.width <= 0 || r.height <= 0) continue;
                labels.push({
                    text: txt,
                    left: Math.round(r.left), right: Math.round(r.right),
                    top: Math.round(r.top), bottom: Math.round(r.bottom),
                    cx: (r.left + r.right) / 2, cy: (r.top + r.bottom) / 2,
                });
            }

            const vOverlap = (a, b) => Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top);

            const pairs = [];
            for (const inp of inputs) {
                // (a) 左侧同一行最近的标签：纵向有重叠且在 input 左边。
                let bestLeft = null, bestLeftDist = Infinity;
                for (const lb of labels) {
                    if (vOverlap(inp, lb) <= 0) continue;      // 不在同一水平带
                    if (lb.right > inp.left + 4) continue;     // 必须在 input 左侧
                    const d = inp.left - lb.right;
                    if (d >= 0 && d < bestLeftDist) { bestLeftDist = d; bestLeft = lb; }
                }
                // (b) 正上方最近的标签（兜底）：横向有重叠且在 input 上方。
                let bestAbove = null, bestAboveDist = Infinity;
                for (const lb of labels) {
                    const hOverlap = Math.min(inp.right, lb.right) - Math.max(inp.left, lb.left);
                    if (hOverlap <= 0) continue;
                    if (lb.bottom > inp.top + 4) continue;     // 必须在上方
                    const d = inp.top - lb.bottom;
                    if (d >= 0 && d < bestAboveDist) { bestAboveDist = d; bestAbove = lb; }
                }
                let label = '', match = 'none', dist = null, lleft = null, ltop = null;
                if (bestLeft && bestLeftDist <= 200) {
                    label = bestLeft.text; match = 'left'; dist = bestLeftDist;
                    lleft = bestLeft.left; ltop = bestLeft.top;
                } else if (bestAbove && bestAboveDist <= 60) {
                    label = bestAbove.text; match = 'above'; dist = bestAboveDist;
                    lleft = bestAbove.left; ltop = bestAbove.top;
                }
                pairs.push({
                    name: inp.name, type: inp.type, value: inp.value,
                    left: inp.left, right: inp.right, top: inp.top,
                    label: label, label_left: lleft, label_top: ltop,
                    match: match, dist: dist,
                });
            }

            // 3) 直接按文本『设备扩展信息』定位它对应的 input（右侧或下方最近）。
            let deviceMsg = null;
            const targets = labels.filter(l => l.text.indexOf('设备扩展信息') >= 0);
            for (const lb of targets) {
                let best = null, bestDist = Infinity;
                for (const inp of inputs) {
                    // 右侧同一行优先
                    if (vOverlap(inp, lb) > 0 && inp.left >= lb.right - 4) {
                        const d = inp.left - lb.right;
                        if (d >= 0 && d < bestDist) { bestDist = d; best = inp; }
                    }
                }
                if (!best) {
                    // 兜底：下方横向重叠最近
                    for (const inp of inputs) {
                        const hOverlap = Math.min(inp.right, lb.right) - Math.max(inp.left, lb.left);
                        if (hOverlap > 0 && inp.top >= lb.bottom - 4) {
                            const d = inp.top - lb.bottom;
                            if (d >= 0 && d < bestDist) { bestDist = d; best = inp; }
                        }
                    }
                }
                if (best) {
                    // FineReport 过滤控件的 input 常常没有 name 属性，真实参数名可能在
                    // input 自身的其它属性（widgetname/id/data-*）或祖先 div 上。把这些都抓出来
                    // 供据实判断真实 name / 稳定选择器。
                    const el = best.el;
                    const attrs = {};
                    for (const a of el.attributes) attrs[a.name] = a.value;
                    const ancestors = [];
                    let p = el.parentElement;
                    for (let d = 0; d < 6 && p; d++) {
                        const pa = {};
                        for (const a of p.attributes) pa[a.name] = a.value;
                        ancestors.push({ tag: p.tagName.toLowerCase(), attrs: pa });
                        p = p.parentElement;
                    }
                    deviceMsg = {
                        label: lb.text, name: best.name, type: best.type,
                        value: best.value, dist: bestDist,
                        label_left: lb.left, label_top: lb.top,
                        input_left: best.left, input_top: best.top,
                        input_attrs: attrs, ancestors: ancestors,
                    };
                    break;
                }
            }

            // 4) 过滤栏候选标签文本（去重，仅短中文），便于人工核对布局顺序。
            const labelTexts = [];
            const seen = {};
            labels.slice().sort((a,b)=> (a.top-b.top) || (a.left-b.left)).forEach(l => {
                if (!seen[l.text]) { seen[l.text] = true; labelTexts.push(l.text); }
            });

            return { pairs, device_msg: deviceMsg, all_label_texts: labelTexts };
        }"""
    )


def dump_filter_inputs_all_frames(page) -> list[dict]:
    """遍历【所有】frame，转储每个 frame 里的过滤 input（name/label 配对）。

    上一轮只在 _report_frame(page) 返回的那个 iframe 里 dump，若『设备扩展信息』框实际
    渲染在另一个 frame（或 _report_frame 命中的 frame 不含过滤栏），就会漏掉。这里对每个
    frame 都跑 dump_filter_label_map，据实记录它各自命中的 input/标签，彻底排除漏看。"""
    results = []
    for i, fr in enumerate(page.frames):
        try:
            data = dump_filter_label_map(fr)
        except Exception as e:
            results.append({"frame": i, "name": fr.name, "url": (fr.url or "")[:100],
                            "error": str(e)})
            continue
        results.append({
            "frame": i,
            "name": fr.name,
            "url": (fr.url or "")[:100],
            "pairs": data.get("pairs", []),
            "device_msg": data.get("device_msg"),
            "all_label_texts": data.get("all_label_texts", []),
        })
    return results


def dump_header_cells(fr) -> list[dict]:
    """转储 table.x-table 前两行（表头）的每个单元格：文字 + id（含列字母/会话锚）。

    不依赖 read_column_layout 内写死的 known 列名，直接抓所有表头单元格，
    这样即便绩效页有新列（如『升级成功率』『升级策略ruleid』），也能被看到。"""
    return fr.evaluate(
        """() => {
            const out = [];
            const xts = Array.from(document.querySelectorAll('table.x-table'));
            for (let ti = 0; ti < xts.length; ti++) {
                const t = xts[ti];
                for (let r = 0; r < Math.min(t.rows.length, 2); r++) {
                    for (const c of t.rows[r].cells) {
                        const txt = (c.innerText || '').trim();
                        const id = c.id || '';
                        const m = id.match(/^([A-Z]+)(\\d+)-(\\d+)-(.+)$/);
                        out.push({
                            table: ti,
                            row: r,
                            text: txt,
                            id: id,
                            letter: m ? m[1] : null,
                            rownum: m ? m[2] : null,
                            session: m ? m[4] : null,
                        });
                    }
                }
            }
            return out;
        }"""
    )


def dump_sample_rows(fr, letters: dict, success_col: str | None,
                     ruleid_col: str | None, max_rows: int = 2) -> list[dict]:
    """读前 max_rows 个数据行里 方案号/目标版本/ruleid/成功率 列的文本 + 对应单元格 id。

    用『方案号』列 id 解析行号 N，再按各列字母拼出单元格 id 并读 innerText。
    success_col / ruleid_col 为探测到的实际表头文本（可能为 None）。"""
    cols = {}
    for name in ("方案号", "目标版本"):
        if letters.get(name):
            cols[name] = letters[name]
    if ruleid_col and letters.get(ruleid_col):
        cols["ruleid"] = letters[ruleid_col]
    if success_col and letters.get(success_col):
        cols["成功率"] = letters[success_col]
    if not letters.get("方案号"):
        return []
    # 绩效页的表头行与数据行处在【同一张】table.x-table（表头=行号1，数据=行号>=2），
    # 不像填报页那样拆成两张表。所以这里不按“表头表 vs 数据表”区分，而是直接扫描
    # 所有 x-table 里 方案号列(letter) 且行号>=2 的单元格，由其 id 解析行号 N，
    # 再按各列字母 + N 拼出同行其它单元格读文本（按 id 前缀 letter+N+'-' 定位）。
    return fr.evaluate(
        """({planLetter, cols, maxRows}) => {
            const xts = Array.from(document.querySelectorAll('table.x-table'));
            const seen = {};
            const res = [];
            const re = new RegExp('^'+planLetter+'(\\\\d+)-');
            for (const t of xts) {
                for (const row of t.rows) {
                    for (const c of row.cells) {
                        const m = (c.id||'').match(re);
                        if (!m) continue;
                        const n = parseInt(m[1]);
                        if (n < 2 || seen[n]) continue;   // 1=表头行；跳过
                        seen[n] = true;
                        const rec = {n:n, cells:{}};
                        for (const col in cols) {
                            const letter = cols[col];
                            const target = document.querySelector(
                                '[id^="'+letter+n+'-"]');
                            rec.cells[col] = {
                                id: target ? (target.id||'') : null,
                                text: target ? (target.innerText||'').trim() : null,
                            };
                        }
                        res.push(rec);
                    }
                }
            }
            res.sort((a,b)=>a.n-b.n);
            return res.slice(0, maxRows);
        }""",
        {"planLetter": letters["方案号"], "cols": cols, "maxRows": max_rows},
    )


def main() -> int:
    reachable = False
    obs: dict = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not config.HEADED)
        context = browser.new_context()
        page = context.new_page()

        # 以防万一出现任何 prompt，一律 dismiss，绝不确认任何对话框。
        def _on_dialog(dialog):
            print(f"  [对话框] type={dialog.type} message={dialog.message!r} -> dismiss")
            try:
                dialog.dismiss()
            except Exception:
                pass
        page.on("dialog", _on_dialog)

        try:
            login(page)
            reachable = True
            opened_leaf = navigate_to_perf(page)
            obs["opened_leaf"] = opened_leaf
            obs["menu_candidates"] = PERF_MENU_CANDIDATES
            page.wait_for_timeout(2000)
            _shot(page, "perf_opened")

            if opened_leaf is None:
                print("✗ 未能通过任何候选叶子打开绩效报表（entry_report frame 未出现）")
                _shot(page, "perf_no_report")
                obs["error"] = "no report frame opened for any candidate leaf"
                raise SystemExit(0)

            fr = _report_frame(page)
            if fr is None:
                print("✗ 未找到绩效页报表 frame")
                _shot(page, "perf_no_frame")
                obs["error"] = "report frame not found"
                raise SystemExit(0)
            obs["frame_url"] = fr.url
            obs["frame_name"] = fr.name
            obs["frame_has_entry_report"] = "entry_report" in (fr.url or "")
            print(f"绩效页 frame：name={fr.name!r} url={fr.url[:100]!r} "
                  f"(含 entry_report: {obs['frame_has_entry_report']})")

            # —— 过滤框 name 转储（查询前，过滤框此时已渲染）——
            filt = dump_filter_inputs(fr)
            obs["filter_inputs"] = filt
            print("\n顶部过滤框（可见 input/textarea/select）：")
            for f in filt:
                if f["visible"] and (f["name"] or f["placeholder"]):
                    print(f"  name={f['name']!r:16} type={f['type']!r:10} "
                          f"placeholder={f['placeholder']!r}  "
                          f"left={f.get('left')} top={f.get('top')}")

            # —— 过滤框 ↔ 标签就近配对（本轮补探测的核心：据实查出『设备扩展信息』的 name）——
            label_map = dump_filter_label_map(fr)
            obs["filter_label_map"] = label_map
            print("\n过滤框 ↔ 标签 就近配对（report frame）：")
            for p in label_map.get("pairs", []):
                if p.get("label") or p.get("name"):
                    print(f"  标签={p.get('label')!r:14} -> name={p.get('name')!r:16} "
                          f"type={p.get('type')!r:8} (match={p.get('match')}, dist={p.get('dist')})")
            dm = label_map.get("device_msg")
            print(f"\n★ 『设备扩展信息』框就近配对结果：{dm}")
            if dm:
                print(f"  input 全属性：{dm.get('input_attrs')}")
                print(f"  祖先链属性：{dm.get('ancestors')}")
            print(f"过滤栏候选标签文本（按位置排序）：{label_map.get('all_label_texts')}")

            # —— 保险起见，遍历所有 frame 再查一遍（防止框在别的 frame）——
            all_frames = dump_filter_inputs_all_frames(page)
            obs["filter_inputs_all_frames"] = all_frames
            print("\n所有 frame 的『设备扩展信息』命中情况：")
            device_msg_hits = []
            for fobj in all_frames:
                dmf = fobj.get("device_msg")
                if dmf:
                    device_msg_hits.append({"frame": fobj["frame"], **dmf})
                    print(f"  frame[{fobj['frame']}] name={fobj.get('name')!r}: {dmf}")
            obs["device_msg_hits"] = device_msg_hits
            # 汇总最终结论：『设备扩展信息』框的真实 name 与稳定选择器。
            # FineReport 过滤 input 常无 name 属性；真实参数名优先取 name，其次 widgetname，
            # 再看 id（形如 'DEVICEMSG' 或 'xxx-DEVICEMSG' 等）。据实记录，缺失则置 None。
            device_msg_name = None
            device_msg_selector = None
            attrs = (dm or {}).get("input_attrs", {}) or {}
            ancestors = (dm or {}).get("ancestors", []) or []
            # FineReport 参数名所在：优先 input 自身的 name；其次 input 自身 widgetname/id；
            # 实测该框参数名其实在祖先 div 的 widgetname 上（class 含 fr-trigger-editor 的那个
            # 控件 wrapper，widgetname='DEVICEMSG'）。跳过纯结构容器（PARA / __layout*__）。
            STRUCT_WIDGETS = {"PARA", "__layout4align__"}
            if dm and dm.get("name"):
                device_msg_name = dm["name"]
                device_msg_selector = f"input[name='{device_msg_name}']"
            elif attrs.get("widgetname"):
                device_msg_name = attrs["widgetname"]
                device_msg_selector = f"input[widgetname='{device_msg_name}']"
            elif attrs.get("id"):
                device_msg_name = attrs["id"]
                device_msg_selector = f"input[id='{device_msg_name}']"
            else:
                # 在祖先链里找最近的有意义 widgetname（跳过结构容器）。
                for anc in ancestors:
                    wn = (anc.get("attrs") or {}).get("widgetname", "").strip()
                    if wn and wn not in STRUCT_WIDGETS:
                        device_msg_name = wn
                        # 该 widgetname 在 wrapper div 上，稳定选择器用 div[widgetname=...]
                        # 内的文本 input 定位。
                        device_msg_selector = (
                            f"div[widgetname='{wn}'] input.fr-trigger-texteditor"
                        )
                        break
                if device_msg_name is None and device_msg_hits and device_msg_hits[0].get("name"):
                    device_msg_name = device_msg_hits[0]["name"]
                    device_msg_selector = f"input[name='{device_msg_name}']"
            obs["device_msg_name"] = device_msg_name
            obs["device_msg_selector"] = device_msg_selector
            obs["device_msg_input_attrs"] = attrs
            print(f"\n★★ 实测『设备扩展信息』过滤框 name = {device_msg_name!r}  "
                  f"选择器 = {device_msg_selector!r}")

            # —— 点一次查询（不填条件）让结果行显示出来 ——
            try:
                set_display_limit(fr, page, LIMIT_START)
            except Exception as e:
                print(f"  ⚠ 设置显示行数失败（绩效页可能无 LIMIT 框）：{e}")
                obs["limit_set_failed"] = str(e)
            try:
                set_filter(fr, page, {})  # 清空所有填报页已知过滤框（若存在）
            except Exception as e:
                print(f"  ⚠ 清空过滤框失败（绩效页过滤框可能不同）：{e}")
                obs["set_filter_failed"] = str(e)
            try:
                click_query(fr, page)
            except Exception as e:
                print(f"  ⚠ 点击查询失败：{e}")
                obs["query_failed"] = str(e)
            page.wait_for_timeout(1500)
            _shot(page, "perf_after_query")

            # —— 表头转储（不依赖写死 known 列名）——
            header_cells = dump_header_cells(fr)
            obs["header_cells"] = header_cells
            print("\n表头单元格（table/row 文字 -> 列字母/会话锚/id）：")
            for h in header_cells:
                if h["text"]:
                    print(f"  t{h['table']}r{h['row']} {h['text']!r:20} "
                          f"letter={h['letter']} session={h['session']} id={h['id']!r}")

            # 从表头抽出 列名->列字母 映射与会话锚。
            letters: dict = {}
            session = None
            for h in header_cells:
                if h["text"] and h["letter"] and h["text"] not in letters:
                    letters[h["text"]] = h["letter"]
                if session is None and h["session"]:
                    session = h["session"]
            obs["letters"] = letters
            obs["session"] = session
            print(f"\n解析得到 列名->列字母：{letters}")
            print(f"会话锚：{session}")

            # read_column_layout 能否解析到成功率新列（它内部 known 不含成功率列）。
            rcl = read_column_layout(fr, debug=False)
            obs["read_column_layout_letters"] = rcl.get("letters", {})
            print(f"\nread_column_layout(known 写死) 解析到：{rcl.get('letters', {})}")

            # —— 定位成功率列与 ruleid 列的实际表头文本 ——
            success_col = None
            for cand in SUCCESS_RATE_CANDIDATES:
                if cand in letters:
                    success_col = cand
                    break
            if success_col is None:
                # 兜底：表头文字里含『成功率』的任意列。
                for txt in letters:
                    if "成功率" in txt:
                        success_col = txt
                        break
            ruleid_col = None
            for cand in RULEID_CANDIDATES:
                if cand in letters:
                    ruleid_col = cand
                    break
            if ruleid_col is None:
                for txt in letters:
                    if "ruleid" in txt.lower() or "rule" in txt.lower():
                        ruleid_col = txt
                        break
            obs["success_col_header"] = success_col
            obs["success_col_letter"] = letters.get(success_col) if success_col else None
            obs["ruleid_col_header"] = ruleid_col
            obs["ruleid_col_letter"] = letters.get(ruleid_col) if ruleid_col else None
            print(f"\n成功率列表头文本：{success_col!r}  列字母：{obs['success_col_letter']!r}")
            print(f"ruleid 列表头文本：{ruleid_col!r}  列字母：{obs['ruleid_col_letter']!r}")
            print(f"read_column_layout 能否解析到成功率列："
                  f"{success_col in rcl.get('letters', {}) if success_col else False}")

            # —— 示例数据行 ——
            samples = dump_sample_rows(fr, letters, success_col, ruleid_col, max_rows=2)
            obs["sample_rows"] = samples
            print("\n示例数据行：")
            for s in samples:
                print(f"  行 {s['n']}：{s['cells']}")

            if config.HEADED:
                try:
                    input("\n浏览器保持打开（只读，未改动）。查看完毕按回车关闭... ")
                except EOFError:
                    time.sleep(10)
        except SystemExit:
            pass
        except Exception as e:
            print(f"✗ 探测出错：{e}")
            obs["exception"] = str(e)
            try:
                _shot(page, "perf_error")
            except Exception:
                pass
        finally:
            context.close()
            browser.close()

    _write_findings(reachable, obs)
    print(f"\nfindings 已写入：{FINDINGS_PATH}")
    print(f"站点可达：{reachable}")
    return 0


def _write_findings(reachable: bool, obs: dict) -> None:
    FINDINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ran = (
        reachable
        and not obs.get("error")
        and not obs.get("exception")
        and obs.get("opened_leaf")
        and obs.get("letters")
    )

    lines: list[str] = []
    lines.append("# DASS 『安卓升级展示绩效』页 只读探测结论（升级成功率闸门 第一步）")
    lines.append("")
    lines.append(f"- 生成时间：{stamp}")
    lines.append(f"- 站点可达：{'是' if reachable else '否'}")
    lines.append(f"- 实跑探测：{'是' if ran else '否（未能实跑，见下方默认假设 + TODO）'}")
    lines.append("")

    if not ran:
        lines.append("## ⚠ 未能实跑补探测（设备扩展信息过滤框 name）")
        lines.append("")
        reason = (
            obs.get("exception")
            or obs.get("error")
            or ("站点不可达 / 登录失败" if not reachable else "未读到绩效页表头布局")
        )
        lines.append(f"- 原因：{reason}")
        lines.append(f"- 尝试过的叶子菜单候选：{obs.get('menu_candidates')}")
        lines.append(f"- 真正打开报表的叶子：{obs.get('opened_leaf')!r}")
        lines.append("")
        lines.append("> 本轮补探测目标：据实查出绩效页顶部『设备扩展信息』过滤框的真实 "
                     "input name。但本次环境未能实跑，下面是**默认假设**，不是实测值。"
                     "站点可用后请重跑 `./.venv/bin/python -m automation.dass_probe_perf` 校准。")
        lines.append("")
        lines.append("### ✅ 本轮采用的默认假设（供下游 Stage2 修改据此实现并加注释）")
        lines.append("")
        lines.append("- **设备扩展信息过滤框 name = `DEVICEMSG`**（填报页 "
                     "`dass_fill.FILTER_INPUTS` 已知 name；绩效页与填报页同为 FineReport 报表，"
                     "最可能一致）。")
        lines.append("- 下游 Stage2 修改应采用的稳定选择器：**`input[name='DEVICEMSG']`**。")
        lines.append("- ⚠ 这是**假设而非实测**；下游实现处必须加注释标明"
                     "『绩效页 DEVICEMSG 框 name 未经实测，沿用填报页已知 name，"
                     "站点可用时以 dass_probe_perf 实测为准』。")
        lines.append("")
        lines.append("### 其余默认假设（未经实测，后续务必核对）")
        lines.append("")
        lines.append("- 绩效页 frame：与填报页一致，取 URL 含 `entry_report` 的 iframe"
                     "（`_report_frame(page)` 已按此逻辑兜底到 `<td>` 最多的 frame）。")
        lines.append("- 其余过滤框 name 假设与填报页相同：")
        lines.append("  - 方案号 `PLANCODE`、目标版本 `NEXTVERSION`、当前阶段 `CURRENTSTAGE`。")
        lines.append("- 成功率列表头文本（上一轮已实测为 `升级成功率`，列字母 Z）。")
        lines.append("- Rule ID 列表头文本（上一轮已实测为 `升级策略ruleid`，列字母 H）。")
        lines.append("- 单元格 id 格式：`{列字母}{行号}-0-{会话锚}`，"
                     "成功率显示形如 `28.26%` / `65.07%` / `50.00%`。")
        lines.append("")
        lines.append("### 后续实现的硬性要求（安全失败，勿误放行）")
        lines.append("")
        lines.append("- **列缺失必须 abort，绝不默认放行**：若查询后表头解析不到成功率列或 "
                     "ruleid 列（列字母为空），闸门必须**终止升级（abort）**并报警，"
                     "而不是当作『全部满足』放行。")
        lines.append("- 成功率解析要去掉 `%` 再转 float，判定用 **>= 50**（见原始需求澄清：>=50%）。")
        lines.append("- ruleid 对比沿用 `norm_ruleid`（去掉逗号等所有分隔符，只留数字）。")
        lines.append("- 绩效页若有新列名（成功率列），`read_column_layout` 内写死的 `known` "
                     "列表可能解析不到它；**让绩效页解析用自己的列名列表**，不要改动 "
                     "`dass_fill.read_column_layout` 的 `known`（避免影响填报流程）。")
        lines.append("")
        FINDINGS_PATH.write_text("\n".join(lines), encoding="utf-8")
        return

    # —— 实跑成功：写实测结论 ——
    letters = obs.get("letters", {})
    rcl_letters = obs.get("read_column_layout_letters", {})
    success_col = obs.get("success_col_header")
    ruleid_col = obs.get("ruleid_col_header")

    lines.append("## 绩效页 菜单叶子与 frame 判定")
    lines.append("")
    lines.append(f"- 真正打开报表的左侧菜单叶子：`{obs.get('opened_leaf')!r}`"
                 "（任务原称『安卓升级展示绩效』，实测以此叶子为准）")
    lines.append(f"- 尝试过的候选：{obs.get('menu_candidates')}")
    lines.append(f"- frame.name：`{obs.get('frame_name')!r}`")
    lines.append(f"- frame.url：`{obs.get('frame_url')}`")
    lines.append(f"- URL 是否含 `entry_report`：{obs.get('frame_has_entry_report')}")
    lines.append("- 判定方式：`_report_frame(page)` 优先取 URL 含 `entry_report` 的 iframe，"
                 "否则回退到 `<td>` 最多的 frame。实测该页" +
                 ("命中 `entry_report`。" if obs.get("frame_has_entry_report")
                  else "**不含 `entry_report`**，靠 `<td>` 最多兜底命中；"
                       "后续实现需确认此兜底对绩效页稳定。"))
    lines.append("")

    lines.append("## 顶部过滤框 name")
    lines.append("")
    lines.append("| name | type | placeholder | 可见 |")
    lines.append("| --- | --- | --- | --- |")
    for f in obs.get("filter_inputs", []):
        if f.get("name") or f.get("placeholder"):
            lines.append(f"| `{f.get('name','')}` | {f.get('type','')} | "
                         f"{f.get('placeholder','')} | {f.get('visible')} |")
    lines.append("")
    names = {f.get("name", "") for f in obs.get("filter_inputs", [])}
    def _chk(n: str) -> str:
        return "✓ 存在" if n in names else "✗ 未发现（与填报页不同，实现需调整选择器）"
    lines.append("与填报页过滤框选择器对照：")
    lines.append("")
    lines.append(f"- 方案号 `PLANCODE`：{_chk('PLANCODE')}")
    lines.append(f"- 目标版本 `NEXTVERSION`：{_chk('NEXTVERSION')}")
    lines.append(f"- 设备扩展信息 `DEVICEMSG`：{_chk('DEVICEMSG')}"
                 " —— ⚠ **这里的 ✗ 仅指 `<input>` 的 `name` 属性**；该框真实存在，"
                 "参数名在外层 `div[widgetname='DEVICEMSG']` 上，详见下一节『🔎 本轮补探测』。")
    lines.append(f"- 当前阶段 `CURRENTSTAGE`：{_chk('CURRENTSTAGE')}")
    lines.append("")
    lines.append("> 注：上表只列出带 `name` 属性的 `<input>`。绩效页『设备扩展信息』过滤框的 "
                 "`<input>` 没有 `name` 属性（参数名在祖先 wrapper 的 `widgetname` 上），"
                 "因此不出现在上表里——这不代表它不存在。")
    lines.append("")

    # —— 本轮补探测核心：『设备扩展信息』过滤框 ↔ 标签 就近配对实测结论 ——
    label_map = obs.get("filter_label_map", {}) or {}
    dm = label_map.get("device_msg")
    dm_name = obs.get("device_msg_name")
    dm_selector = obs.get("device_msg_selector")
    dm_attrs = obs.get("device_msg_input_attrs", {}) or {}
    lines.append("## 🔎 本轮补探测：『设备扩展信息』过滤框真实 name（推翻上一轮结论）")
    lines.append("")
    lines.append("> 上一轮错误结论：『绩效页没有设备扩展信息(DEVICEMSG)过滤框，"
                 "设备扩展信息只作为结果列存在』。用户已用页面截图确认：绩效页顶部"
                 "确实存在一个『设备扩展信息』文本输入框（位于『国家』右侧、『当前策略』左侧）。"
                 "本轮用【标签文字 ↔ input 就近几何配对】据实重新探测，**结论：该框确实存在**。")
    lines.append("")
    lines.append(f"- **『设备扩展信息:』标签右侧确实存在一个文本输入框**（实测标签在 "
                 f"left={ (dm or {}).get('label_left') }, top={ (dm or {}).get('label_top') }，"
                 f"input 紧邻其右侧 left={ (dm or {}).get('input_left') }, "
                 f"top={ (dm or {}).get('input_top') }，几何距离 { (dm or {}).get('dist') }px）。"
                 "推翻上一轮『无此框』的结论。")
    lines.append(f"- 该 `<input>` 本身的 `name` 属性实测为：`{(dm or {}).get('name')!r}`"
                 "（空）。FineReport 过滤控件的 `<input class=\"fr-trigger-texteditor\">` "
                 "**没有 name 属性**——这正是上一轮只查 `input[name=...]` 而误判『无此框』的原因。")
    lines.append(f"- 该 input 的全部属性（实测）：`{dm_attrs}`")
    lines.append("- **真实参数名在祖先控件 wrapper 的 `widgetname` 上**：该 input 外层紧邻一个 "
                 "`<div class=\"fr-trigger-editor ...\" widgetname=\"DEVICEMSG\">`（实测，见下方 "
                 "ancestors）。即该『设备扩展信息』过滤框的 FineReport 参数名就是 "
                 "**`DEVICEMSG`**，与填报页 `dass_fill.FILTER_INPUTS['设备扩展信息']` 的 "
                 "`DEVICEMSG` **一致**。")
    if dm_name:
        lines.append(f"- **结论：『设备扩展信息』过滤框参数名 = `{dm_name}`**。")
        lines.append(f"- 下游 Stage2 修改应采用的稳定选择器：**`{dm_selector}`**")
        lines.append("  （用 wrapper 的 `widgetname='DEVICEMSG'` 锚定再取其内的文本 input，"
                     "比只写 `input[name='DEVICEMSG']` 更稳——因为该 input 没有 name 属性，"
                     "绩效页上 `input[name='DEVICEMSG']` **选不中**）。")
    else:
        lines.append("- ⚠ 未能据实读出参数名（既无 input name，祖先也无 widgetname）；"
                     "下游请用几何定位（『设备扩展信息:』标签右侧最近的 "
                     "`input.fr-trigger-texteditor`），兜底沿用 `DEVICEMSG` 并加注释。")
    lines.append("")
    lines.append("就近配对原始命中（含 input 全属性与祖先链，供下游选稳定选择器）：")
    lines.append("")
    lines.append("```json")
    lines.append(json.dumps(dm, ensure_ascii=False, indent=2))
    lines.append("```")
    lines.append(f"- 所有 frame device_msg 命中：`{obs.get('device_msg_hits')}`")
    lines.append("")
    lines.append("report frame 内 input ↔ 标签 配对表：")
    lines.append("")
    lines.append("| 标签 | input name | type | 配对方式 | 距离(px) |")
    lines.append("| --- | --- | --- | --- | --- |")
    for p in label_map.get("pairs", []):
        if p.get("label") or p.get("name"):
            lines.append(f"| {p.get('label','')} | `{p.get('name','')}` | "
                         f"{p.get('type','')} | {p.get('match','')} | {p.get('dist')} |")
    lines.append("")
    lines.append(f"- 过滤栏候选标签文本（按屏幕位置排序）：`{label_map.get('all_label_texts')}`")
    lines.append("")

    lines.append("## 表头 列名 -> 列字母 映射")
    lines.append("")
    lines.append(f"- 会话锚（session）：`{obs.get('session')}`")
    lines.append("")
    lines.append("| 列名 | 列字母 |")
    lines.append("| --- | --- |")
    for name, letter in letters.items():
        lines.append(f"| {name} | {letter} |")
    lines.append("")

    lines.append("## 成功率列 与 ruleid 列")
    lines.append("")
    lines.append(f"- 成功率列表头文本：`{success_col!r}`  列字母：`{obs.get('success_col_letter')!r}`")
    lines.append(f"- ruleid 列表头文本：`{ruleid_col!r}`  列字母：`{obs.get('ruleid_col_letter')!r}`")
    lines.append(f"- 单元格 id 格式：`{{列字母}}{{行号}}-0-{{会话锚}}`，"
                 f"例如成功率格 `{obs.get('success_col_letter')}<行号>-0-{obs.get('session')}`。")
    lines.append("")
    in_rcl = success_col in rcl_letters if success_col else False
    lines.append(f"- `read_column_layout`（dass_fill 内写死 known）能否解析到成功率列："
                 f"**{'能' if in_rcl else '不能'}**。")
    if not in_rcl:
        lines.append("  - 因此后续绩效页解析**必须用自己的列名列表**（把成功率列表头文本加入），"
                     "**不要改动 `dass_fill.read_column_layout` 的 `known`**，以免影响填报流程。")
    lines.append(f"- `read_column_layout` 实测解析到的列：`{rcl_letters}`")
    lines.append("")

    lines.append("## 示例数据行")
    lines.append("")
    lines.append("```json")
    lines.append(json.dumps(obs.get("sample_rows", []), ensure_ascii=False, indent=2))
    lines.append("```")
    lines.append("")

    lines.append("## ⚠ 与任务原假设的关键差异（实现必须据此修正）")
    lines.append("")
    lines.append(f"1. 成功率列表头是 **`升级成功率`（列字母 {letters.get('升级成功率')}）**，"
                 "不是任务假设的 `升级成功率（11）`。"
                 f"`升级成功率（11）` 不存在；带『（11）』的是另一列 "
                 f"**`升级成功设备数（11）`（设备数，列字母 {letters.get('升级成功设备数（11）')}）**，"
                 "是设备数不是比率，切勿当成功率用。")
    lines.append(f"2. ruleid 列表头是 **`升级策略ruleid`（列字母 {letters.get('升级策略ruleid')}）**，"
                 "不是裸 `ruleid`。")
    _fallback_sel = "div[widgetname='DEVICEMSG'] input.fr-trigger-texteditor"
    lines.append("3. **（更正上一轮错误结论）** 绩效页顶部**确实存在『设备扩展信息』过滤框**，"
                 f"FineReport 参数名 = `{dm_name or 'DEVICEMSG'}`（与填报页一致）。"
                 "该框的 `<input>` 本身没有 name 属性，参数名在其外层 "
                 "`div[widgetname='DEVICEMSG']` 上，所以 `input[name='DEVICEMSG']` **选不中**；"
                 f"下游稳定选择器用 `{dm_selector or _fallback_sel}`。")
    lines.append("   - ⚠ 上一轮『绩效页没有 DEVICEMSG 过滤框、设备扩展信息仅作为结果列存在』的"
                 "结论**错误，已推翻**（上一轮只查 `input[name=...]`，漏看了 wrapper 的 "
                 "widgetname）。设备扩展信息既是结果【列】（列字母 "
                 f"{letters.get('设备扩展信息')}），也是顶部可用的【过滤框】。")
    lines.append("   - 影响：任务澄清的『查不到 ruleid 就按 方案号+目标版本+设备扩展码+当前阶段 "
                 "精准查询』在绩效页**可以用该设备扩展信息过滤框实现**，Stage2 不必再退化成三字段。")
    lines.append(f"4. 升级成功率示例值形如 `28.26%`（带 `%`，需去 `%` 转 float 再比 >= 50）。")
    lines.append("")

    lines.append("## 后续实现要点（供下一步，不在本步实施）")
    lines.append("")
    lines.append("- 闸门在升级 DASS 之前执行：按 方案号+目标版本（模糊）查询绩效页，"
                 "按 ruleid 对比成功率 **>= 50%**，全部满足才放行，否则 abort。")
    lines.append("- 成功率解析：去掉 `%` 后转 float，判定 `>= 50.0`。")
    lines.append("- ruleid 对比沿用 `norm_ruleid`（去掉逗号等分隔符只留数字）。")
    lines.append("- 查不到某 ruleid 时，按澄清用 方案号+目标版本+设备扩展码+当前阶段 精准查询。")
    lines.append("- **列缺失安全失败**：解析不到成功率/ruleid 列字母时必须 abort 升级，"
                 "绝不默认放行。")
    lines.append("")
    FINDINGS_PATH.write_text("\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    sys.exit(main())
