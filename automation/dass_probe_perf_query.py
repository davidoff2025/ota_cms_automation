"""只读探测（补）：绩效页『安卓升级展示报表』的【查询交互机制】。

运行（项目根目录）：
    cd ota_cms_automation
    CMS_HEADED=true ./.venv/bin/python -m automation.dass_probe_perf_query

背景：HEADED 实跑发现，照搬填报页的 `input[name='LIMIT']` 与 `#fr-btn-BTNSUBMIT` 在绩效页
失效——LIMIT 框找不到、点查询后网格仍停在默认未过滤数据（1/190、A311D2）且一直 Loading，
read_perf_rows 读到 0 行。本脚本【只读】据实探测以下三件事并把结论追加写进
.agents/tasks/perf-probe-findings.md：
  1）『显示行数/LIMIT』在绩效页的真实控件（name / widgetname / 是否根本不存在、用别的分页）。
  2）绿色『查询』按钮在绩效页的真实选择器（id/class/文本/widgetname）。
  3）点查询后『重查真正完成』的可靠信号（会话锚变化？网格重建？分页从 1/190 变？行数变？）。

只读安全：只登录、导航、填两个过滤框(方案号/目标版本)、点一次查询、读 DOM。绝不插入行、
绝不填任何单元格、绝不点提交/校验/升级。所有 JS prompt 一律 dismiss。
"""
from __future__ import annotations

import json
import time
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

from . import config
from .dass_login import login, _shot
from .dass_fill import _report_frame
from .dass_perf import (
    navigate_to_perf,
    read_perf_layout,
    read_perf_rows,
    PERF_RULEID_HEADER,
    PERF_SUCCESS_RATE_HEADER,
    PERF_PLAN_HEADER,
)

FINDINGS_PATH = Path(__file__).resolve().parents[1] / ".agents" / "tasks" / "perf-probe-findings.md"

PROBE_PLAN = "MT9603AndrGTV4K"
PROBE_TARGET = "0100AQ0725"


def dump_clickables(fr) -> list[dict]:
    """转储报表 frame 内所有『可点』控件：button / [onclick] / FineReport 自由按钮
    (div.fr-btn / [widgetname]) / 含『查询』文字的元素。记录 id/class/文本/widgetname/坐标。"""
    return fr.evaluate(
        r"""() => {
            const out = [];
            const seen = new Set();
            const push = (e, why) => {
                if (seen.has(e)) return;
                seen.add(e);
                const r = e.getBoundingClientRect();
                out.push({
                    why,
                    tag: e.tagName.toLowerCase(),
                    id: e.id || '',
                    cls: e.getAttribute('class') || '',
                    widgetname: e.getAttribute('widgetname') || '',
                    name: e.getAttribute('name') || '',
                    text: (e.innerText || e.value || '').trim().slice(0, 24),
                    title: e.getAttribute('title') || '',
                    visible: r.width > 0 && r.height > 0,
                    left: Math.round(r.left), top: Math.round(r.top),
                    w: Math.round(r.width), h: Math.round(r.height),
                });
            };
            document.querySelectorAll('button, input[type=button], input[type=submit]').forEach(e => push(e, 'button'));
            document.querySelectorAll('[onclick]').forEach(e => push(e, 'onclick'));
            document.querySelectorAll("[id^='fr-btn-']").forEach(e => push(e, 'fr-btn-id'));
            document.querySelectorAll('[widgetname]').forEach(e => {
                const wn = (e.getAttribute('widgetname') || '');
                if (/BTN|SUBMIT|QUERY|SEARCH/i.test(wn)) push(e, 'widgetname-btn');
            });
            // 任何直接文本是『查询』的元素（FineReport 自由按钮常是 div，文字『查询』）。
            document.querySelectorAll('div, span, a, td').forEach(e => {
                const t = (e.innerText || '').trim();
                if (t === '查询' || t === '查 询') push(e, 'text-查询');
            });
            return out;
        }"""
    )


def dump_all_controls(fr) -> list[dict]:
    """转储顶部所有 widgetname 控件 + 所有 input（找 LIMIT/显示行数/分页控件）。"""
    return fr.evaluate(
        r"""() => {
            const out = [];
            document.querySelectorAll('[widgetname]').forEach(e => {
                const r = e.getBoundingClientRect();
                out.push({
                    kind: 'widget',
                    tag: e.tagName.toLowerCase(),
                    widgetname: e.getAttribute('widgetname') || '',
                    id: e.id || '',
                    cls: (e.getAttribute('class') || '').slice(0, 60),
                    text: (e.innerText || '').trim().slice(0, 24),
                    left: Math.round(r.left), top: Math.round(r.top),
                });
            });
            document.querySelectorAll('input, select, textarea').forEach(e => {
                const r = e.getBoundingClientRect();
                out.push({
                    kind: 'input',
                    tag: e.tagName.toLowerCase(),
                    name: e.getAttribute('name') || '',
                    cls: (e.getAttribute('class') || '').slice(0, 60),
                    value: (e.value || '').slice(0, 24),
                    left: Math.round(r.left), top: Math.round(r.top),
                    visible: r.width > 0 && r.height > 0,
                });
            });
            return out;
        }"""
    )


def dump_pager(fr) -> dict:
    """抓工具栏分页信息：当前页/总页数文本（形如 '1 / 190'）、x-table 数量、Loading 遮罩是否可见。"""
    return fr.evaluate(
        r"""() => {
            const bodyText = (document.body.innerText || '');
            // 分页：匹配形如 '1/190' 或 '1 / 190' 的片段。
            let pager = null;
            const m = bodyText.match(/(\d+)\s*\/\s*(\d+)/);
            if (m) pager = m[0].replace(/\s+/g, '');
            // Loading 遮罩：FineReport 常用 .fr-core-loading / 文本『Loading』。
            const loadingEl = document.querySelector(
                '.fr-core-loading, .loading, [class*="loading"]');
            let loadingVisible = false;
            if (loadingEl) {
                const r = loadingEl.getBoundingClientRect();
                loadingVisible = r.width > 0 && r.height > 0;
            }
            const hasLoadingText = /Loading/i.test(bodyText);
            const xtables = document.querySelectorAll('table.x-table').length;
            return { pager, loadingVisible, hasLoadingText, xtables };
        }"""
    )


def grid_state(fr) -> dict:
    """综合快照：会话锚、perf 行数、前 3 行方案号、分页/Loading。用于比较查询前后变化。"""
    layout = read_perf_layout(fr, debug=False)
    letters, session = layout["letters"], layout["session"]
    rows = read_perf_rows(fr, letters) if letters.get(PERF_PLAN_HEADER) else []
    pager = dump_pager(fr)
    plans = []
    try:
        plans = fr.evaluate(
            r"""(planLetter) => {
                const out = [];
                const xts = Array.from(document.querySelectorAll('table.x-table'));
                for (const t of xts) {
                    for (const row of t.rows) {
                        for (const c of row.cells) {
                            const m = (c.id||'').match(new RegExp('^'+planLetter+'(\\d+)-'));
                            if (m && parseInt(m[1]) >= 2) {
                                const txt = (c.innerText||'').trim();
                                if (txt) out.push(txt);
                            }
                        }
                    }
                }
                return out.slice(0, 3);
            }""",
            letters.get(PERF_PLAN_HEADER) or "B",
        )
    except Exception:
        pass
    return {
        "session": session,
        "row_count": len(rows),
        "ruleid_letter": letters.get(PERF_RULEID_HEADER),
        "rate_letter": letters.get(PERF_SUCCESS_RATE_HEADER),
        "first_plans": plans,
        "pager": pager.get("pager"),
        "loading_text": pager.get("hasLoadingText"),
        "xtables": pager.get("xtables"),
    }


def dump_col_letters_in_dom(fr) -> dict:
    """直接扫 DOM，看 H（升级策略ruleid）/ Z（升级成功率）两列的【数据行单元格】是否已渲染。

    不依赖表头 known 列名，纯按单元格 id 前缀 '<letter><rownum>-' 统计（rownum>=2 为数据行），
    用于验证『横向滚动前 vs 后』H/Z 两列单元格是否出现在 DOM（FineReport 懒渲染验证）。"""
    return fr.evaluate(
        r"""() => {
            const countLetter = (letter) => {
                let n = 0; const sample = [];
                const re = new RegExp('^'+letter+'(\\d+)-\\d+-');
                document.querySelectorAll('[id]').forEach(e => {
                    const m = (e.id||'').match(re);
                    if (m && parseInt(m[1]) >= 2) {
                        n++;
                        if (sample.length < 3) sample.push({id: e.id, text: (e.innerText||'').trim().slice(0,12)});
                    }
                });
                return {count: n, sample};
            };
            return { H: countLetter('H'), Z: countLetter('Z') };
        }"""
    )


def scroll_grid_horizontally(fr, page) -> str:
    """把结果网格横向滚动到最右，让右侧的 H/Z 列进入视口以触发懒渲染。

    FineReport 报表网格通常有一个水平滚动容器（.x-table 外层带 overflow 的 div）。
    策略：对所有可滚动容器设 scrollLeft = scrollWidth（滚到最右），分几次递增滚动以触发
    虚拟渲染，每步稍等。返回滚动前后 scrollLeft 摘要。"""
    before = fr.evaluate(
        r"""() => {
            const els = Array.from(document.querySelectorAll('div')).filter(e => {
                return e.scrollWidth > e.clientWidth + 20 && e.clientWidth > 100;
            });
            return els.map(e => ({cls: (e.className||'').slice(0,40), sl: e.scrollLeft,
                                  sw: e.scrollWidth, cw: e.clientWidth})).slice(0, 10);
        }"""
    )
    # 分段横向滚动（0% -> 25% -> ... -> 100%），每步触发懒渲染。
    for frac in (0.25, 0.5, 0.75, 1.0):
        fr.evaluate(
            r"""(frac) => {
                document.querySelectorAll('div').forEach(e => {
                    if (e.scrollWidth > e.clientWidth + 20 && e.clientWidth > 100) {
                        e.scrollLeft = Math.round((e.scrollWidth - e.clientWidth) * frac);
                    }
                });
            }""",
            frac,
        )
        page.wait_for_timeout(500)
    after = fr.evaluate(
        r"""() => {
            const els = Array.from(document.querySelectorAll('div')).filter(e => {
                return e.scrollWidth > e.clientWidth + 20 && e.clientWidth > 100;
            });
            return els.map(e => ({cls: (e.className||'').slice(0,40), sl: e.scrollLeft,
                                  sw: e.scrollWidth, cw: e.clientWidth})).slice(0, 10);
        }"""
    )
    return json.dumps({"before": before, "after": after}, ensure_ascii=False)


def try_fill(fr, selector: str, value: str) -> str:
    loc = fr.locator(selector).first
    try:
        loc.scroll_into_view_if_needed(timeout=2000)
        loc.click()
        loc.fill("")
        loc.fill(value)
        return "ok"
    except Exception as e:
        return f"FAIL: {e}"


def try_click(fr, selector: str) -> str:
    loc = fr.locator(selector).first
    try:
        loc.scroll_into_view_if_needed(timeout=2000)
        loc.click()
        return "ok"
    except Exception as e:
        return f"FAIL: {e}"


def main() -> int:
    obs: dict = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not config.HEADED)
        context = browser.new_context()
        page = context.new_page()

        def _on_dialog(dialog):
            print(f"  [对话框] {dialog.type} {dialog.message!r} -> dismiss")
            try:
                dialog.dismiss()
            except Exception:
                pass
        page.on("dialog", _on_dialog)

        try:
            login(page)
            navigate_to_perf(page)
            page.wait_for_timeout(2500)
            fr = _report_frame(page)
            if fr is None:
                print("✗ 未找到报表 frame")
                obs["error"] = "no frame"
                raise SystemExit(0)
            obs["frame_url"] = (fr.url or "")[:120]

            # —— 1) 所有可点控件（找查询按钮）——
            clickables = dump_clickables(fr)
            obs["clickables"] = clickables
            print("\n==== 可点控件（button / fr-btn / 文字=查询 / widgetname-btn）====")
            for c in clickables:
                if c["visible"]:
                    print(f"  [{c['why']:14}] tag={c['tag']:6} id={c['id']!r:24} "
                          f"widgetname={c['widgetname']!r:16} text={c['text']!r:10} "
                          f"cls={c['cls'][:40]!r} @({c['left']},{c['top']})")

            # —— 2) 所有控件（找 LIMIT/显示行数/分页）——
            controls = dump_all_controls(fr)
            obs["controls"] = controls
            print("\n==== 所有 widgetname 控件 ====")
            for c in controls:
                if c["kind"] == "widget" and c.get("widgetname"):
                    print(f"  widgetname={c['widgetname']!r:20} tag={c['tag']:6} "
                          f"text={c['text']!r:12} @({c['left']},{c['top']})")
            print("\n==== 所有 input ====")
            for c in controls:
                if c["kind"] == "input" and c.get("visible"):
                    print(f"  name={c['name']!r:18} cls={c['cls'][:40]!r:42} "
                          f"value={c['value']!r} @({c['left']},{c['top']})")

            # —— 3) 查询前快照 ——
            before = grid_state(fr)
            obs["state_before"] = before
            print(f"\n==== 查询前快照 ====\n  {before}")
            _shot(page, "probeq_before")

            # —— 填方案号+目标版本 ——
            r1 = try_fill(fr, "input[name='PLANCODE']", PROBE_PLAN)
            r2 = try_fill(fr, "input[name='NEXTVERSION']", PROBE_TARGET)
            obs["fill_plancode"] = r1
            obs["fill_nextversion"] = r2
            print(f"\n填 PLANCODE={PROBE_PLAN!r} -> {r1}")
            print(f"填 NEXTVERSION={PROBE_TARGET!r} -> {r2}")

            # —— 逐一尝试几个候选查询按钮选择器，记录哪个点得中 ——
            query_candidates = [
                "#fr-btn-BTNSUBMIT",
                "div[widgetname='BTNSUBMIT']",
                "text=查询",
                "div.fr-btn:has-text('查询')",
            ]
            # 真点查询：按屏幕上文字『查询』的元素点（最贴近用户操作）。
            clicked_sel = None
            click_results = {}
            # 先只探测哪些选择器“存在且可见”，但只真正点一次（点文字=查询）。
            for sel in query_candidates:
                try:
                    cnt = fr.locator(sel).count()
                except Exception as e:
                    cnt = f"err:{e}"
                click_results[sel] = {"count": cnt}
            obs["query_selector_counts"] = click_results
            print("\n==== 查询按钮候选选择器命中数 ====")
            for sel, info in click_results.items():
                print(f"  {sel!r:40} -> count={info['count']}")

            # 真正点一次：优先用文字『查询』。
            print("\n点击『查询』（文字选择器） ...")
            clicked = try_click(fr, "text=查询")
            if clicked != "ok":
                clicked = try_click(fr, "#fr-btn-BTNSUBMIT")
                clicked_sel = "#fr-btn-BTNSUBMIT" if clicked == "ok" else None
            else:
                clicked_sel = "text=查询"
            obs["query_clicked_via"] = clicked_sel
            obs["query_click_result"] = clicked
            print(f"  点击结果：{clicked}（选择器={clicked_sel!r}）")

            # —— 查询后：按时间轮询快照，观察“何时真正重查完成” ——
            timeline = []
            for i in range(16):  # 最多 ~16*800ms = 12.8s
                page.wait_for_timeout(800)
                snap = grid_state(fr)
                timeline.append({"t_ms": (i + 1) * 800, **snap})
                print(f"  t={snap and (i+1)*800:>5}ms session={snap['session']!r} "
                      f"rows={snap['row_count']} pager={snap['pager']!r} "
                      f"loading={snap['loading_text']} plans={snap['first_plans']}")
                # 完成信号候选：会话锚变了，且结果行里方案号命中我们的过滤方案（或行数明显变小、Loading 消失）。
                if (snap["session"] != before["session"]
                        and not snap["loading_text"]
                        and snap["row_count"] > 0):
                    print("  ✓ 检测到疑似重查完成（session 变化 + 无 Loading + 有行）。")
                    break
            obs["state_after_timeline"] = timeline
            after = timeline[-1] if timeline else None
            obs["state_after"] = after

            # —— 关键验证：横向滚动前后，H/Z 两列数据单元格是否出现在 DOM（懒渲染）——
            hz_before_scroll = dump_col_letters_in_dom(fr)
            obs["hz_before_scroll"] = hz_before_scroll
            print(f"\n==== 横向滚动【前】H/Z 列数据单元格计数 ====\n  {hz_before_scroll}")

            print("\n横向滚动结果网格（触发右侧 H/Z 列懒渲染） ...")
            scroll_info = scroll_grid_horizontally(fr, page)
            obs["scroll_info"] = scroll_info
            print(f"  滚动容器 scrollLeft 变化：{scroll_info}")

            hz_after_scroll = dump_col_letters_in_dom(fr)
            obs["hz_after_scroll"] = hz_after_scroll
            print(f"\n==== 横向滚动【后】H/Z 列数据单元格计数 ====\n  {hz_after_scroll}")

            # 滚动后再读一次 perf 布局/行，确认能否真正读到 ruleid/rate。
            after2 = grid_state(fr)
            obs["state_after_scroll"] = after2
            print(f"\n==== 横向滚动后 grid_state ====\n  {after2}")
            _shot(page, "probeq_after_scroll")

            print(f"\n==== 查询后最终快照 ====\n  {after}")
            if after:
                print("\n==== 前后对比 ====")
                print(f"  session: {before['session']!r} -> {after['session']!r} "
                      f"（{'变化' if before['session'] != after['session'] else '未变'}）")
                print(f"  row_count: {before['row_count']} -> {after['row_count']}")
                print(f"  pager: {before['pager']!r} -> {after['pager']!r}")
                print(f"  first_plans: {before['first_plans']} -> {after['first_plans']}")

            if config.HEADED:
                try:
                    input("\n浏览器保持打开（只读）。回车关闭... ")
                except EOFError:
                    time.sleep(5)
        except SystemExit:
            pass
        except Exception as e:
            print(f"✗ 探测出错：{e}")
            obs["exception"] = str(e)
            try:
                _shot(page, "probeq_error")
            except Exception:
                pass
        finally:
            context.close()
            browser.close()

    _append_findings(obs)
    print(f"\n补探测结论已追加写入：{FINDINGS_PATH}")
    return 0


def _append_findings(obs: dict) -> None:
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines: list[str] = []
    lines.append("")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append("## 🔎 第二轮补探测：绩效页【查询交互机制】（LIMIT / 查询按钮 / 完成信号）")
    lines.append("")
    lines.append(f"- 生成时间：{stamp}")
    lines.append(f"- 探测方案号+目标版本：`{PROBE_PLAN}` + `{PROBE_TARGET}`")
    lines.append(f"- 报表 frame url：`{obs.get('frame_url')}`")
    if obs.get("error") or obs.get("exception"):
        lines.append(f"- ⚠ 未能完整实跑：{obs.get('exception') or obs.get('error')}")
    lines.append("")

    lines.append("### 原始转储（JSON，供下游据实选选择器）")
    lines.append("")
    lines.append("```json")
    dump = {
        "clickables_visible": [c for c in obs.get("clickables", []) if c.get("visible")],
        "widgets": [c for c in obs.get("controls", []) if c.get("kind") == "widget" and c.get("widgetname")],
        "inputs_visible": [c for c in obs.get("controls", []) if c.get("kind") == "input" and c.get("visible")],
        "query_selector_counts": obs.get("query_selector_counts"),
        "query_clicked_via": obs.get("query_clicked_via"),
        "query_click_result": obs.get("query_click_result"),
        "fill_plancode": obs.get("fill_plancode"),
        "fill_nextversion": obs.get("fill_nextversion"),
        "state_before": obs.get("state_before"),
        "state_after": obs.get("state_after"),
        "state_after_timeline": obs.get("state_after_timeline"),
        "hz_before_scroll": obs.get("hz_before_scroll"),
        "hz_after_scroll": obs.get("hz_after_scroll"),
        "scroll_info": obs.get("scroll_info"),
        "state_after_scroll": obs.get("state_after_scroll"),
    }
    lines.append(json.dumps(dump, ensure_ascii=False, indent=2))
    lines.append("```")
    lines.append("")
    FINDINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
    with FINDINGS_PATH.open("a", encoding="utf-8") as f:
        f.write("\n".join(lines))


if __name__ == "__main__":
    import sys
    sys.exit(main())
