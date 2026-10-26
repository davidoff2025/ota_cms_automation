"""READ-ONLY probe of the FineReport 查询 (Query) UI above the fill-in grid.

Run:
    cd ota_cms_automation
    ./.venv/bin/python -m automation.dass_probe_query

This probe is strictly READ-ONLY with respect to data. The ONLY live actions it
performs are:
  - fill the top filter boxes (方案号 etc.),
  - set the 显示行数 'limit' input,
  - click the green 查询 (Query) button,
  - read back the returned rows.
It NEVER clicks 提交/Submit, 数据校验/Data validation, 删除行/Delete, never writes a
cell value, and never touches the 推荐1/推荐2 tabs.

It discovers and prints robust selectors for every control in the filter form,
sets the limit to 50 and (optionally) 100, runs ONE safe query cycle, reads the
result rows via the existing dass_fill.existing_row_info reader, and determines
empirically:
  (a) whether read_column_layout must be re-read after each 查询 (session anchor /
      column letters change on grid refresh?),
  (b) whether a single vs double click enters a cell's edit mode (observed by
      entering edit mode and pressing Escape WITHOUT writing — never committing).

Findings are also written to .agents/tasks/probe-findings.md.
"""
from __future__ import annotations

import sys
import time
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

from . import config
from .dass_login import login, navigate_to_fillin, _shot
from .dass_fill import _report_frame, read_column_layout, existing_row_info

# A harmless/representative 方案号 to type into the filter box. We do NOT rely on
# it matching anything; the goal is only to observe the query cycle. Leaving the
# boxes empty and just clicking 查询 is also acceptable (and we fall back to that
# if the 方案号 box cannot be located).
SAMPLE_SCHEME_NO = ""  # empty => just click 查询 and observe (safest)

# Confirmed FineReport filter input `name` attributes (verified live + against the
# screenshot by matching each input's x-position to its label). These `name`
# anchors are stable and are the preferred selectors for the implementer.
FILTER_NAME_BY_FIELD = {
    "方案号": "PLANCODE",
    "目标版本": "NEXTVERSION",   # UI label '目标版本' (shown '标版本' when truncated)
    "设备扩展信息": "DEVICEMSG",
    "当前阶段": "CURRENTSTAGE",
    "大区域": "REGION",
    "国家": "COUNTRYCODE",
    "机型": "DEVICECODE",
    "源版本": "CURRENTVERSION",
}
LIMIT_INPUT_NAME = "LIMIT"       # value shows 'limit 10'
QUERY_BUTTON_SELECTOR = "#fr-btn-BTNSUBMIT"  # the green 查询 button (div id)

FINDINGS_PATH = (
    Path(__file__).resolve().parents[1] / ".agents" / "tasks" / "probe-findings.md"
)


def discover_filter_controls(fr) -> dict:
    """Return a structured description of every <input>/button in the report frame.

    For each input we capture id / name / type / placeholder / value / className,
    its bounding box, and the nearest label text we can find (a preceding/adjacent
    cell or widget whose text names the field, e.g. '方案号', '标版本', '显示行数').
    FineReport renders filter widgets as absolutely-positioned divs with an input
    inside; the label is a separate text cell to the left. We pair each input with
    the closest text to its left on (roughly) the same row.
    """
    return fr.evaluate(
        r"""() => {
            const vis = (el) => {
                const r = el.getBoundingClientRect();
                return r.width > 0 && r.height > 0;
            };
            const rectOf = (el) => {
                const r = el.getBoundingClientRect();
                return {x: Math.round(r.x), y: Math.round(r.y),
                        w: Math.round(r.width), h: Math.round(r.height)};
            };

            // Collect candidate label texts: any small element with short text.
            const textNodes = [];
            const walkEls = document.querySelectorAll('div, td, span, label');
            for (const el of walkEls) {
                if (!vis(el)) continue;
                // only direct text (avoid huge containers)
                const own = Array.from(el.childNodes)
                    .filter(n => n.nodeType === 3)
                    .map(n => n.textContent.trim())
                    .join('');
                const txt = (own || '').trim();
                if (!txt || txt.length > 12) continue;
                const r = el.getBoundingClientRect();
                textNodes.push({txt, x: Math.round(r.x), y: Math.round(r.y),
                                w: Math.round(r.width), h: Math.round(r.height)});
            }

            const nearestLabel = (ir) => {
                // nearest text whose vertical center overlaps the input row and
                // sits to the LEFT of the input (typical label placement).
                const iyc = ir.y + ir.h / 2;
                let best = null, bestDist = 1e9;
                for (const t of textNodes) {
                    const tyc = t.y + t.h / 2;
                    if (Math.abs(tyc - iyc) > 16) continue;   // same row
                    if (t.x > ir.x + 4) continue;              // to the left
                    const dist = ir.x - (t.x + t.w);
                    if (dist < -4) continue;
                    if (dist < bestDist) { bestDist = dist; best = t.txt; }
                }
                return best;
            };

            const inputs = [];
            for (const el of document.querySelectorAll('input, textarea')) {
                if (!vis(el)) continue;
                const r = rectOf(el);
                inputs.push({
                    tag: el.tagName.toLowerCase(),
                    type: el.getAttribute('type') || '',
                    id: el.id || '',
                    name: el.getAttribute('name') || '',
                    placeholder: el.getAttribute('placeholder') || '',
                    value: el.value || '',
                    cls: el.className || '',
                    widgetName: el.getAttribute('widgetname')
                                || el.getAttribute('data-id') || '',
                    rect: r,
                    label: nearestLabel(r),
                });
            }
            inputs.sort((a, b) => (a.rect.y - b.rect.y) || (a.rect.x - b.rect.x));

            const buttons = [];
            const bsel = "button, input[type=button], input[type=submit], "
                       + "a, div[role='button'], .x-button, [class*='fr-btn'], [class*='x-emb']";
            for (const el of document.querySelectorAll(bsel)) {
                if (!vis(el)) continue;
                const txt = (el.innerText || el.value || '').trim();
                if (!txt && !/fr-btn|x-emb|x-button/.test(el.className || '')) continue;
                buttons.push({
                    tag: el.tagName.toLowerCase(),
                    text: txt,
                    id: el.id || '',
                    cls: el.className || '',
                    rect: rectOf(el),
                });
            }
            return {inputs, buttons};
        }"""
    )


def find_query_button(fr):
    """Return a Playwright locator for the green 查询 button (visible), or None.

    The robust anchor is the FineReport free-button div id `#fr-btn-BTNSUBMIT`.
    We prefer that; fall back to text matching only if the id is missing."""
    primary = fr.locator(QUERY_BUTTON_SELECTOR)
    try:
        if primary.count() and primary.first.is_visible():
            return primary.first, QUERY_BUTTON_SELECTOR
    except Exception:
        pass
    # FineReport free buttons render as a div/a with text '查询'.
    candidates = [
        "a:has-text('查询')",
        "div[role='button']:has-text('查询')",
        "button:has-text('查询')",
        "[class*='x-button']:has-text('查询')",
        "[class*='fr-btn']:has-text('查询')",
        "*:has-text('查询')",
    ]
    for sel in candidates:
        loc = fr.locator(sel)
        for i in range(loc.count()):
            el = loc.nth(i)
            try:
                if not el.is_visible():
                    continue
                txt = (el.inner_text() or "").strip()
                # Leaf element whose text is exactly / almost '查询'.
                if txt == "查询" or (len(txt) <= 4 and "查询" in txt):
                    return el, sel
            except Exception:
                continue
    return None, None


def find_input_by_label(fr, label_substrings: list[str]):
    """Return (locator, matched_label, descriptor) for the filter input whose
    nearest label text contains any of label_substrings. Uses geometry discovered
    by discover_filter_controls to build a stable #id / [name] / nth selector."""
    data = discover_filter_controls(fr)
    for inp in data["inputs"]:
        lab = (inp.get("label") or "")
        if any(s in lab for s in label_substrings):
            sel = _selector_for_input(inp)
            if sel is None:
                continue
            return fr.locator(sel).first, lab, inp
    return None, None, None


def _selector_for_input(inp: dict) -> str | None:
    """Build the most robust CSS selector we can for a discovered input dict."""
    if inp.get("id"):
        return f"#{_css_escape(inp['id'])}"
    if inp.get("name"):
        return f"input[name='{inp['name']}']"
    if inp.get("widgetName"):
        return f"input[widgetname='{inp['widgetName']}']"
    if inp.get("cls"):
        first = inp["cls"].split()[0]
        return f"input.{_css_escape(first)}"
    return None


def _css_escape(s: str) -> str:
    # ids in FineReport can contain odd chars; escape for querySelector/#id use.
    out = []
    for ch in s:
        if ch.isalnum() or ch in "-_":
            out.append(ch)
        else:
            out.append("\\" + ch)
    return "".join(out)


def set_limit_input(fr, page, value: str) -> tuple[bool, str | None]:
    """Find the 显示行数 limit input and set it to the full SQL fragment 'limit N'.

    IMPORTANT (confirmed live): the box content is interpolated VERBATIM into the
    dataset SQL as a LIMIT clause. It ships showing 'limit 10'; filling it with a
    bare number ('50') yields invalid SQL ('... 50') and the query errors with
    『数据集配置错误 ... error in your SQL syntax near '50'』. So we must write the
    full fragment, e.g. 'limit 50' / 'limit 100'. `value` may be '50' or
    'limit 50'; we normalize to 'limit N'. Returns (ok, selector_used)."""
    num = str(value).strip().lower().replace("limit", "").strip()
    payload = f"limit {num}"
    # Primary: the confirmed name anchor (value shows 'limit 10').
    sel = f"input[name='{LIMIT_INPUT_NAME}']"
    if not fr.locator(sel).count():
        # Fallbacks if the name is missing: label text, then value 'limit ...'.
        data = discover_filter_controls(fr)
        target = None
        for inp in data["inputs"]:
            lab = (inp.get("label") or "")
            val = (inp.get("value") or "")
            if ("显示行数" in lab or "limit" in lab.lower()
                    or "limit" in val.lower() or "limit" in (inp.get("placeholder") or "").lower()):
                target = inp
                break
        if target is None:
            return False, None
        sel = _selector_for_input(target)
        if sel is None:
            return False, None
    try:
        loc = fr.locator(sel).first
        loc.scroll_into_view_if_needed(timeout=2000)
        loc.click()
        loc.fill("")
        loc.fill(payload)
        page.wait_for_timeout(200)
        print(f"  显示行数 已填为 {payload!r}")
        return True, sel
    except Exception as e:
        print(f"  ✗ 设置显示行数失败（{sel}）：{e}")
        return False, sel


def probe_click_mode(fr, page, letters: dict, session: str) -> str:
    """Observe (do NOT commit) whether single vs double click enters edit mode on
    a data cell. Enters edit mode, then presses Escape WITHOUT typing anything, so
    no value is ever written. Returns a human-readable finding string."""
    plan_letter = letters.get("方案号")
    if not plan_letter or session is None:
        return "无法测试：缺少方案号列字母或会话锚。"
    existing = existing_row_info(fr, letters)
    if not existing:
        return "无法测试：当前网格无已存在数据行（查询结果为空）。"
    row_n = existing[0]["n"]
    cell_id = f"{plan_letter}{row_n}-0-{session}"
    cell = fr.locator(f"#{cell_id}")
    try:
        cell.scroll_into_view_if_needed(timeout=2000)
    except Exception:
        pass
    cb = cell.bounding_box()

    def aligned_editor_present() -> bool:
        ed = fr.locator("input.fr-texteditor, textarea.fr-texteditor")
        for i in range(ed.count()):
            try:
                e = ed.nth(i)
                if not e.is_visible():
                    continue
                eb = e.bounding_box()
                if cb and eb and abs(eb["y"] - cb["y"]) < 30 and abs(eb["x"] - cb["x"]) < 260:
                    return True
            except Exception:
                continue
        return False

    # --- single click ---
    single_ok = False
    try:
        cell.click()
        page.wait_for_timeout(600)
        single_ok = aligned_editor_present()
    except Exception as e:
        print(f"  单击测试异常：{e}")
    # exit edit mode without writing
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(200)
    except Exception:
        pass

    # --- double click ---
    double_ok = False
    try:
        cell.dblclick()
        page.wait_for_timeout(600)
        double_ok = aligned_editor_present()
    except Exception as e:
        print(f"  双击测试异常：{e}")
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(200)
    except Exception:
        pass

    print(f"  单击进入编辑态={single_ok}  双击进入编辑态={double_ok}  （测试单元格 {cell_id}，已 Escape，未写入）")
    if single_ok and double_ok:
        return f"单击即可进入编辑态（单击={single_ok}，双击={double_ok}）。实现可用单击。"
    if single_ok:
        return f"单击即可进入编辑态（单击={single_ok}，双击={double_ok}）。"
    if double_ok:
        return f"需双击进入编辑态（单击={single_ok}，双击={double_ok}）。dass_fill.fill_cell 的双击做法成立。"
    return f"单击与双击均未检出对位编辑框（单击={single_ok}，双击={double_ok}）；需进一步排查。"


def _wait_for_grid(fr, page, timeout_ms: int = 20000) -> bool:
    """After clicking 查询 the grid detaches and re-renders with a NEW session
    anchor; poll read_column_layout until the header reappears (方案号 column +
    session present). Returns True once the grid is ready."""
    waited = 0
    step = 800
    while waited < timeout_ms:
        page.wait_for_timeout(step)
        waited += step
        try:
            lay = read_column_layout(fr, debug=False)
            if lay.get("session") is not None and "方案号" in lay.get("letters", {}):
                # small extra settle so data rows are painted too
                page.wait_for_timeout(600)
                print(f"  （网格已重绘，用时约 {waited}ms；新会话锚={lay['session']!r}）")
                return True
        except Exception:
            continue
    print(f"  ⚠ 等待 {timeout_ms}ms 后网格表头仍未就绪。")
    return False


def _fmt_input(inp: dict) -> str:
    r = inp.get("rect", {})
    return (
        f"label={inp.get('label')!r} id={inp.get('id')!r} name={inp.get('name')!r} "
        f"placeholder={inp.get('placeholder')!r} value={inp.get('value')!r} "
        f"cls={inp.get('cls')!r} widget={inp.get('widgetName')!r} "
        f"@({r.get('x')},{r.get('y')} {r.get('w')}x{r.get('h')})"
    )


def main() -> int:
    findings: dict = {
        "reachable": False,
        "filters": {},          # label -> selector info
        "limit": {},            # selector + mechanism
        "query_button": None,
        "reader": None,
        "reread_layout": None,
        "click_mode": None,
        "sample_rows": [],
        "raw_inputs": [],
        "raw_buttons": [],
        "error": None,
    }

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not config.HEADED)
        context = browser.new_context()
        page = context.new_page()

        # READ-ONLY probe: never answer an insert-row prompt. Dismiss any dialog
        # (we never trigger 插入行, but stay safe if one appears).
        def _on_dialog(dialog):
            print(f"  [对话框] type={dialog.type} message={dialog.message!r} -> dismiss")
            try:
                dialog.dismiss()
            except Exception:
                pass
        page.on("dialog", _on_dialog)

        try:
            login(page)
            navigate_to_fillin(page)
            page.wait_for_timeout(3000)

            fr = _report_frame(page)
            if fr is None:
                print("✗ 未找到填报表 frame —— 站点可能不可达或页面结构变化。")
                findings["error"] = "report frame not found"
                _write_findings(findings)
                return 1
            findings["reachable"] = True
            print(f"填报表 frame：{(fr.url or '')[:90]!r}")
            _shot(page, "query_probe_initial")

            # -------- 1) Discover every filter control & button --------
            print("\n==================== 过滤控件探测 ====================")
            disc = discover_filter_controls(fr)
            findings["raw_inputs"] = disc["inputs"]
            findings["raw_buttons"] = disc["buttons"]
            print(f"可见 input/textarea 数量：{len(disc['inputs'])}")
            for inp in disc["inputs"]:
                print("  INPUT " + _fmt_input(inp))
            print(f"\n可见按钮数量：{len(disc['buttons'])}")
            for b in disc["buttons"]:
                r = b["rect"]
                print(f"  BTN {b['text']!r:10} <{b['tag']}> id={b['id']!r} cls={b['cls']!r} "
                      f"@({r['x']},{r['y']})")

            # -------- 2) Map the labeled filter inputs we care about --------
            print("\n==================== 关键过滤框定位 ====================")
            label_specs = {
                "方案号": ["方案号"],
                "目标版本": ["标版本", "目标版本"],  # UI label 目标版本 (可能被截断为 标版本)
                "设备扩展信息": ["设备扩展信息", "扩展信息"],
                "当前阶段": ["当前阶段", "阶段"],
                "大区域": ["大区域", "区域"],
                "国家": ["国家"],
                "机型": ["机型"],
                "源版本": ["源版本"],
            }
            # Index discovered inputs by name so we can confirm each name anchor
            # actually exists in the live DOM (and capture its rect/value).
            disc_by_name = {inp.get("name"): inp for inp in disc["inputs"] if inp.get("name")}
            for key, name_attr in FILTER_NAME_BY_FIELD.items():
                sel = f"input[name='{name_attr}']"
                present = fr.locator(sel).count() > 0
                inp = disc_by_name.get(name_attr, {})
                if present:
                    # best-effort label from geometry (informational only)
                    _, lab, _ = find_input_by_label(fr, label_specs.get(key, [key]))
                    findings["filters"][key] = {
                        "selector": sel,
                        "name": name_attr,
                        "matched_label": lab,
                        "value": inp.get("value"),
                        "cls": inp.get("cls"),
                        "rect": inp.get("rect"),
                    }
                    print(f"  ✓ {key}: selector={sel!r} (value={inp.get('value')!r} rect={inp.get('rect')})")
                else:
                    findings["filters"][key] = None
                    print(f"  ✗ {key}: name={name_attr!r} 的输入框不存在（DOM 可能已变化）。")

            # -------- 3) Query button --------
            qbtn, qsel = find_query_button(fr)
            if qbtn is not None:
                try:
                    qbox = qbtn.bounding_box()
                except Exception:
                    qbox = None
                findings["query_button"] = {"selector": qsel, "box": qbox}
                print(f"\n  ✓ 查询按钮：selector={qsel!r} box={qbox}")
            else:
                findings["query_button"] = None
                print("\n  ✗ 未定位到『查询』按钮。")

            # -------- 4) Baseline layout BEFORE query --------
            print("\n==================== 查询前表头布局 ====================")
            layout_before = read_column_layout(fr, debug=False)
            letters_before = layout_before["letters"]
            session_before = layout_before["session"]
            print(f"  会话锚(before)={session_before!r}")
            print(f"  列字母(before)={letters_before}")

            # -------- 5) Set limit = 50 --------
            print("\n==================== 设置显示行数=50 ====================")
            ok50, lim_sel = set_limit_input(fr, page, "50")
            findings["limit"] = {
                "selector": lim_sel,
                "set_50": ok50,
                "mechanism": (
                    "该框内容被原样拼进数据集 SQL 的 LIMIT 子句，必须写完整片段 "
                    "'limit N'（不能只写数字）。设 50 => fill('limit 50')；设 100 => "
                    "fill('limit 100')。只写 '50' 会报 11300001 数据集配置错误"
                    "（SQL 语法错误 near '50'）。"
                ),
            }
            print(f"  set limit=50: ok={ok50} selector={lim_sel!r}")

            # -------- 6) Optionally type a sample 方案号 (empty = skip) --------
            if SAMPLE_SCHEME_NO:
                sloc, slab, sinp = find_input_by_label(fr, ["方案号"])
                if sloc is not None:
                    try:
                        sloc.click()
                        sloc.fill("")
                        sloc.fill(SAMPLE_SCHEME_NO)
                        print(f"  已向方案号过滤框输入 {SAMPLE_SCHEME_NO!r}")
                    except Exception as e:
                        print(f"  ⚠ 输入方案号失败：{e}")
            else:
                print("  （方案号过滤框留空，直接点查询观察行为）")

            # -------- 7) Click 查询 and wait --------
            print("\n==================== 点击查询并等待结果 ====================")
            if qbtn is not None:
                try:
                    qbtn.scroll_into_view_if_needed(timeout=2000)
                    qbtn.click()
                    print("  ✓ 已点击『查询』，等待结果加载 ...")
                    _wait_for_grid(fr, page, timeout_ms=20000)
                    _shot(page, "query_probe_after_query_50")
                except Exception as e:
                    print(f"  ✗ 点击查询失败：{e}")
            else:
                print("  ✗ 无查询按钮，跳过查询点击。")

            # -------- 8) Layout AFTER query: did session / letters change? --------
            print("\n==================== 查询后表头布局 ====================")
            layout_after = read_column_layout(fr, debug=False)
            letters_after = layout_after["letters"]
            session_after = layout_after["session"]
            print(f"  会话锚(after)={session_after!r}")
            print(f"  列字母(after)={letters_after}")
            session_changed = session_before != session_after
            letters_changed = letters_before != letters_after
            findings["reread_layout"] = {
                "session_before": session_before,
                "session_after": session_after,
                "letters_before": letters_before,
                "letters_after": letters_after,
                "session_changed": session_changed,
                "letters_changed": letters_changed,
            }
            print(f"  会话锚是否变化={session_changed}  列字母是否变化={letters_changed}")

            # -------- 9) Read result rows via existing_row_info --------
            print("\n==================== 读取查询结果行 ====================")
            rows = existing_row_info(fr, letters_after)
            findings["reader"] = "existing_row_info(fr, letters) 读取查询结果（方案号非空的数据行）"
            findings["result_count_50"] = len(rows)
            print(f"  结果数据行数（existing_row_info）：{len(rows)}")
            sample = []
            for r in rows[:8]:
                sample.append({
                    "n": r.get("n"),
                    "方案号": r.get("方案号", ""),
                    "ruleid": r.get("ruleid", ""),
                    "目标版本": r.get("目标版本", ""),
                    "当前阶段": r.get("当前阶段", ""),
                })
                print(f"    行{r.get('n')}: 方案号={r.get('方案号','')!r} ruleid={r.get('ruleid','')!r} "
                      f"目标版本={r.get('目标版本','')!r} 当前阶段={r.get('当前阶段','')!r}")
            findings["sample_rows"] = sample

            # If results == 50 we would bump the limit to 100 (per user rule). We
            # demonstrate the limit=100 mechanism when appropriate.
            if len(rows) >= 50:
                print("\n  结果行数 >= 50，演示把显示行数改为 100 并重查 ...")
                ok100, lim_sel2 = set_limit_input(fr, page, "100")
                findings["limit"]["set_100"] = ok100
                if ok100 and qbtn is not None:
                    try:
                        qbtn.click()
                        _wait_for_grid(fr, page, timeout_ms=20000)
                        _shot(page, "query_probe_after_query_100")
                        layout_100 = read_column_layout(fr, debug=False)
                        rows100 = existing_row_info(fr, layout_100["letters"])
                        print(f"  limit=100 后结果行数：{len(rows100)}")
                        findings["sample_rows_100_count"] = len(rows100)
                    except Exception as e:
                        print(f"  ⚠ limit=100 重查失败：{e}")
            else:
                print(f"\n  结果行数 {len(rows)} < 50，无需升到 100（但机制同 fill('100')）。")

            # -------- 10) single vs double click (observe only, Escape, no write) --------
            print("\n==================== 单击/双击进入编辑态（只观察，不写入）====================")
            layout_now = read_column_layout(fr, debug=False)
            click_finding = probe_click_mode(
                fr, page, layout_now["letters"], layout_now["session"]
            )
            findings["click_mode"] = click_finding
            print(f"  结论：{click_finding}")

            _shot(page, "query_probe_done")
            _write_findings(findings)
            print("\n✅ 查询 UI 探测完成（全程只读，未提交/未写入）。")
            return 0
        except Exception as e:
            print(f"✗ 出错：{e}")
            findings["error"] = str(e)
            _shot(page, "query_probe_error")
            _write_findings(findings)
            return 1
        finally:
            context.close()
            browser.close()


def _write_findings(f: dict) -> None:
    FINDINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def sel_of(key):
        v = f["filters"].get(key)
        if not v:
            return "（未定位）"
        return f"`{v['selector']}`  (name={v.get('name')!r}, 当前value={v.get('value')!r})"

    lines = []
    lines.append("# DASS 查询 UI 探测结果（实现者可直接复制）")
    lines.append("")
    lines.append(f"- 生成时间：{stamp}")
    lines.append(f"- 站点可达：{'是' if f.get('reachable') else '否'}")
    if f.get("error"):
        lines.append(f"- 错误：`{f['error']}`")
    lines.append("")
    lines.append("## 过滤框选择器")
    lines.append("")
    lines.append("| 字段 | 选择器 |")
    lines.append("| --- | --- |")
    for key in ["方案号", "目标版本", "设备扩展信息", "当前阶段", "大区域", "国家", "机型", "源版本"]:
        note = ""
        if key == "目标版本":
            note = "（UI 标签显示为『标版本』）"
        lines.append(f"| {key}{note} | {sel_of(key)} |")
    lines.append("")
    lines.append("## 显示行数（limit）输入框")
    lim = f.get("limit") or {}
    lines.append(f"- 选择器：`{lim.get('selector')}`")
    lines.append(f"- 设为 50：{lim.get('set_50')}；设为 100：{lim.get('set_100', '未演示')}")
    lines.append(f"- 机制：{lim.get('mechanism')}")
    lines.append("")
    lines.append("## 查询按钮")
    qb = f.get("query_button")
    if qb:
        lines.append(f"- 选择器：`{qb.get('selector')}`（取可见且文本=『查询』的叶子元素；box={qb.get('box')}）")
    else:
        lines.append("- 未定位到查询按钮。")
    lines.append("")
    lines.append("## 结果读取方式")
    lines.append(f"- {f.get('reader') or '（未验证）'}")
    lines.append("- 结论：查询后网格显示的行即查询结果，`dass_fill.existing_row_info(fr, letters)` 可直接读取（返回方案号非空的数据行，带 n/方案号/设备扩展信息/机型/目标版本/ruleid/当前阶段）。")
    lines.append("")
    lines.append("## 查询后是否需重读表头布局（read_column_layout）")
    rl = f.get("reread_layout")
    if rl:
        lines.append(f"- 会话锚：before={rl['session_before']!r} after={rl['session_after']!r} -> 变化={rl['session_changed']}")
        lines.append(f"- 列字母是否变化：{rl['letters_changed']}")
        if rl["session_changed"] or rl["letters_changed"]:
            lines.append("- 结论：**每次查询后必须重新调用 `read_column_layout(fr)`**（会话锚/列字母会随网格刷新改变），再用新的 letters/session 去读行与拼单元格 id。")
        else:
            lines.append("- 结论：本次查询前后会话锚与列字母**未变化**；但稳妥起见，实现仍建议在每次查询后重读一次（成本低、可防刷新改变）。")
    else:
        lines.append("- 未取得前后对比数据。")
    lines.append("")
    lines.append("## 单击 vs 双击进入单元格编辑态")
    lines.append(f"- {f.get('click_mode') or '（未验证）'}")
    lines.append("- 说明：测试时仅进入编辑态后按 Escape 退出，**未写入任何值**。")
    lines.append("")
    lines.append("## 样例查询结果")
    sr = f.get("sample_rows") or []
    cnt50 = f.get("result_count_50", len(sr))
    lines.append(f"- 结果行数（limit=50）：{cnt50}" + (f"；limit=100 行数：{f['sample_rows_100_count']}" if 'sample_rows_100_count' in f else ""))
    lines.append(f"- （下表为前 {len(sr)} 行样例）")
    if sr:
        lines.append("")
        lines.append("| 行号n | 方案号 | ruleid | 目标版本 | 当前阶段 |")
        lines.append("| --- | --- | --- | --- | --- |")
        for r in sr:
            lines.append(f"| {r['n']} | {r['方案号']} | {r['ruleid']} | {r['目标版本']} | {r['当前阶段']} |")
    lines.append("")
    lines.append("## 原始探测转储（供排错）")
    lines.append("")
    lines.append("### inputs")
    for inp in (f.get("raw_inputs") or []):
        lines.append(f"- {_fmt_input(inp)}")
    lines.append("")
    lines.append("### buttons")
    for b in (f.get("raw_buttons") or []):
        r = b["rect"]
        lines.append(f"- {b['text']!r} <{b['tag']}> id={b['id']!r} cls={b['cls']!r} @({r['x']},{r['y']})")
    lines.append("")

    FINDINGS_PATH.write_text("\n".join(lines), encoding="utf-8")
    print(f"  已写出探测结论：{FINDINGS_PATH}")


if __name__ == "__main__":
    sys.exit(main())
