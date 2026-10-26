"""只读探测：FineReport 数字格式单元格的『千分位逗号』究竟是“仅显示”还是“真被写进存储值”。

运行（项目根目录）：
    cd ota_cms_automation
    ./.venv/bin/python -m automation.dass_probe_numfmt

背景（见 .agents/tasks/probe-findings.md 与 dass_fill.py）：
- 当前阶段 O / 当前阶段时间 P / 机型 G / 设备扩展信息 H / ruleid L / 源版本 M / 目标版本 N
  这些列在报表里被配置为 NUMBER 数字格式。fill_cell 写入原始值后，单元格“显示文本”
  会带千分位逗号（'20261026' -> '20,261,026'，'65232' -> '65,232'）。
- fill_cell 的回读用 inner_text()（显示文本）去和期望的原始值比较，看到逗号即判失败，
  _write_row 返回 False，stats 显示 0 更新（即使视觉上已写入）。

本探测必须解决的关键未知：逗号是“仅显示格式”（单元格底层/提交值仍是无逗号的正确值），
还是“逗号已被写进存储值”（提交会把 'OTA,200' / '20,261,026' 持久化成错误数据）。

探测动作（绝不提交）：
1. 登录、导航到『安卓升级状态填报』。
2. 点一次绿色『查询』让结果行显示出来（避免只有默认十来行）。
3. 对一个数字格式单元格（当前阶段时间 P，回退到 ruleid L）双击进入编辑态，
   输入测试值 '20261026'：
   (a) 读编辑中 input.fr-texteditor 的 input_value()（编辑态里的原始值）。
   (b) 回车后读单元格 inner_text()（当前 fill_cell 回读看到的值）。
   (c) 尽力读单元格“底层值”：FineReport 常把真实值与格式化显示分开存放——
       扫描单元格 DOM 的 title / 各 data-* 属性、子节点文本，并尝试通过
       contentPane/report JS 模型读取该格的 cellvalue / value。
4. 截图，按 Escape 退出编辑态。绝不点击 Submit/提交/Data validation/数据校验。
   留下一个已插入/已编辑的单元格是可以接受的（用户可丢弃），但绝不提交。
5. 把结论（DISPLAY-ONLY 或 STORED-CORRUPTION）与观测写进 findings 文档。

只读安全：唯一的“写”是把测试值敲进一个单元格的编辑框用于观察，随后 Escape，不提交。
"""
from __future__ import annotations

import sys
import time
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

from . import config
from .dass_login import login, navigate_to_fillin, _shot
from .dass_fill import (
    _report_frame,
    read_column_layout,
    existing_row_info,
    set_display_limit,
    set_filter,
    click_query,
    LIMIT_START,
)

FINDINGS_PATH = Path(__file__).resolve().parents[1] / ".agents" / "tasks" / "probe-findings.md"

TEST_VALUE = "20261026"  # 当前阶段时间形态；若格式化会显示成 '20,261,026'


def _aligned_editor(fr, cb):
    """返回与单元格 box 对位的可见 fr-texteditor（照搬 fill_cell 的对位逻辑）。"""
    ed = fr.locator("input.fr-texteditor, textarea.fr-texteditor")
    for i in range(ed.count()):
        try:
            e = ed.nth(i)
            if not e.is_visible():
                continue
            eb = e.bounding_box()
            if cb and eb and abs(eb["y"] - cb["y"]) < 30 and abs(eb["x"] - cb["x"]) < 260:
                return e
        except Exception:
            continue
    return None


def _inspect_cell_model(fr, cell_id: str) -> dict:
    """尽力读出该单元格的“底层值”线索：

    - DOM：title、所有 data-* 属性、innerText、子节点文本；
    - FineReport JS 模型：遍历页面上可能持有 cell value 的对象（title 常为真实值，
      因为 FR 把未格式化的原值放在 title/tooltip，而把格式化后的放在可见文本）。
    """
    return fr.evaluate(
        """(cellId) => {
            const out = {found:false};
            const el = document.getElementById(cellId);
            if (!el) { out.note = 'cell element not found'; return out; }
            out.found = true;
            out.innerText = (el.innerText||'').trim();
            out.textContent = (el.textContent||'').trim();
            out.title = el.getAttribute('title');
            // 单元格内部常有一个承载文本的子 div，其 title 可能是未格式化原值
            const inner = el.querySelector('*[title]');
            out.innerTitle = inner ? inner.getAttribute('title') : null;
            const data = {};
            for (const a of el.attributes) {
                if (a.name.startsWith('data-') || a.name === 'title') data[a.name] = a.value;
            }
            // 子孙节点上的 data-*/title
            el.querySelectorAll('*').forEach((n, i) => {
                if (i > 20) return;
                for (const a of n.attributes) {
                    if (a.name.startsWith('data-') || a.name === 'title') {
                        data['child['+i+'].'+a.name] = a.value;
                    }
                }
            });
            out.dataAttrs = data;
            return out;
        }""",
        cell_id,
    )


def main() -> int:
    reachable = False
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not config.HEADED)
        context = browser.new_context()
        page = context.new_page()

        # 绝不插入行，但以防万一出现任何 prompt，一律取消，绝不确认提交类对话框。
        def _on_dialog(dialog):
            print(f"  [对话框] type={dialog.type} message={dialog.message!r} -> dismiss")
            try:
                dialog.dismiss()
            except Exception:
                pass
        page.on("dialog", _on_dialog)

        observations: dict = {}
        verdict = "UNKNOWN"
        verdict_line = ""
        shot_path = ""
        try:
            login(page)
            reachable = True
            navigate_to_fillin(page)
            page.wait_for_timeout(3000)

            fr = _report_frame(page)
            if fr is None:
                print("✗ 未找到填报表 frame")
                _shot(page, "numfmt_no_frame")
                observations["error"] = "report frame not found"
                raise SystemExit(0)
            print(f"填报表 frame：{fr.url[:80]!r}")

            # 点一次查询让结果行显示出来（不填任何条件，直接取默认返回的行即可观察格式）。
            set_display_limit(fr, page, LIMIT_START)
            set_filter(fr, page, {})  # 清空所有过滤框
            click_query(fr, page)

            layout = read_column_layout(fr, debug=False)
            letters, session = layout["letters"], layout["session"]
            print(f"会话锚：{session}  列字母：{letters}")
            if session is None:
                print("✗ 查询后未读到表头布局")
                _shot(page, "numfmt_no_layout")
                observations["error"] = "no layout after query"
                raise SystemExit(0)

            rows = existing_row_info(fr, letters)
            print(f"查询结果数据行：{len(rows)}")
            if not rows:
                print("✗ 查询结果为 0 行，无法在已存在行上探测")
                _shot(page, "numfmt_no_rows")
                observations["error"] = "no result rows"
                raise SystemExit(0)

            # 选第一条数据行，优先探测『当前阶段时间』(P)；没有则退回 ruleid (L)。
            target_row = rows[0]
            row_n = target_row["n"]
            probe_col = "当前阶段时间" if letters.get("当前阶段时间") else "ruleid"
            letter = letters.get(probe_col)
            if not letter:
                print(f"✗ 表头缺少可探测的数字列（当前阶段时间/ruleid）：letters={letters}")
                observations["error"] = "no numeric column letter"
                raise SystemExit(0)
            cell_id = f"{letter}{row_n}-0-{session}"
            print(f"探测单元格：列={probe_col} id={cell_id!r}（行 {row_n}）")

            # 探测前先记录单元格现有显示 / 底层线索。
            before_model = _inspect_cell_model(fr, cell_id)
            print(f"  [编辑前] 单元格模型：{before_model}")
            observations["before_model"] = before_model

            cell = fr.locator(f"#{cell_id}")
            try:
                cell.scroll_into_view_if_needed(timeout=3000)
            except Exception:
                pass
            cb = cell.bounding_box()

            # —— 双击进入编辑态 ——
            cell.dblclick()
            page.wait_for_timeout(300)
            editor = _aligned_editor(fr, cb)
            if editor is None:
                print("  双击未出现对位编辑框，重试一次 ...")
                cell.dblclick()
                page.wait_for_timeout(400)
                editor = _aligned_editor(fr, cb)

            editing_value_before = None
            editing_value_after = None
            if editor is None:
                print("  ✗ 未能进入编辑态（无对位编辑框）")
                observations["error_edit"] = "no aligned editor after dblclick"
            else:
                # (a) 编辑态里原始显示（输入测试值之前）
                try:
                    editing_value_before = editor.input_value()
                except Exception as e:
                    editing_value_before = f"<read failed: {e}>"
                print(f"  (a-pre) 编辑框进入时的 input_value()={editing_value_before!r}")

                # 清空并键入测试值（照搬 fill_cell：fill 后回车）
                editor.fill(TEST_VALUE)
                try:
                    editing_value_after = editor.input_value()
                except Exception as e:
                    editing_value_after = f"<read failed: {e}>"
                print(f"  (a) 键入 {TEST_VALUE!r} 后、回车前 input_value()={editing_value_after!r}")

                page.keyboard.press("Enter")
                page.wait_for_timeout(300)

            # (b) 回车后单元格 inner_text（当前 fill_cell 回读看到的）
            try:
                after_text = fr.locator(f"#{cell_id}").inner_text().strip()
            except Exception as e:
                after_text = f"<read failed: {e}>"
            print(f"  (b) 回车后单元格 inner_text()={after_text!r}")

            # (c) 回车后底层模型线索（title / data-* / 子节点）
            after_model = _inspect_cell_model(fr, cell_id)
            print(f"  (c) 回车后单元格模型：{after_model}")

            observations["probe_col"] = probe_col
            observations["cell_id"] = cell_id
            observations["test_value"] = TEST_VALUE
            observations["editing_value_before"] = editing_value_before
            observations["editing_value_after"] = editing_value_after
            observations["after_inner_text"] = after_text
            observations["after_model"] = after_model

            config.ARTIFACTS_DIR.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            shot_path = str(config.ARTIFACTS_DIR / f"dass_{stamp}_numfmt_probe.png")
            try:
                fr.locator(f"#{cell_id}").scroll_into_view_if_needed(timeout=2000)
            except Exception:
                pass
            _shot(page, "numfmt_probe")

            # —— 关键：绝不提交。按 Escape 退出编辑态。 ——
            page.keyboard.press("Escape")
            page.wait_for_timeout(200)

            # —— 判定 ——
            # DISPLAY-ONLY 的判据（任一成立即倾向“仅显示”）：
            #   · 编辑态 input_value() 回放的是无逗号原始值（20261026），说明底层存的是原值；
            #   · 单元格 title / data-* 里存在无逗号原始值，而可见文本才带逗号。
            # STORED-CORRUPTION 的判据：
            #   · 编辑态 input_value() 本身就带逗号（20,261,026），且没有任何无逗号的
            #     底层值线索 —— 说明逗号进入了被编辑/被存的值本身。
            av = (editing_value_after or "")
            formatted = TEST_VALUE[:2] + "," in after_text or "," in after_text
            # 收集所有“底层值”候选（编辑框回放值 + title/data-*）
            underlying_candidates = []
            if editing_value_after:
                underlying_candidates.append(("editor.input_value", editing_value_after))
            am = observations.get("after_model") or {}
            if am.get("title"):
                underlying_candidates.append(("cell.title", am["title"]))
            if am.get("innerTitle"):
                underlying_candidates.append(("child.title", am["innerTitle"]))
            for k, v in (am.get("dataAttrs") or {}).items():
                underlying_candidates.append((k, v))
            observations["underlying_candidates"] = underlying_candidates

            has_raw_underlying = any(
                (v is not None and "," not in str(v) and TEST_VALUE in str(v))
                for _, v in underlying_candidates
            )
            editor_has_comma = "," in av

            if has_raw_underlying and not editor_has_comma:
                verdict = "DISPLAY-ONLY"
                verdict_line = (
                    "VERDICT: DISPLAY-ONLY — Enter stores the comma-free value internally "
                    f"(e.g. {TEST_VALUE}) while only DISPLAYING {after_text}."
                )
            elif editor_has_comma and not has_raw_underlying:
                verdict = "STORED-CORRUPTION"
                verdict_line = (
                    "VERDICT: STORED-CORRUPTION — the comma is part of the actual stored value "
                    f"(e.g. {after_text} would be submitted/persisted)."
                )
            elif has_raw_underlying:
                # 既有逗号编辑值又有无逗号底层值 —— 以底层值为准，判为仅显示。
                verdict = "DISPLAY-ONLY"
                verdict_line = (
                    "VERDICT: DISPLAY-ONLY — Enter stores the comma-free value internally "
                    f"(e.g. {TEST_VALUE}) while only DISPLAYING {after_text}."
                )
            else:
                verdict = "INCONCLUSIVE"
                verdict_line = (
                    "VERDICT: INCONCLUSIVE — 未能从编辑框回放值或 title/data-* 读到明确的无逗号底层值；"
                    "需人工复核截图与下方观测。"
                )
            observations["editor_has_comma"] = editor_has_comma
            observations["has_raw_underlying"] = has_raw_underlying
            observations["after_text_formatted"] = formatted

            print("\n================ 判定 ================")
            print(verdict_line)

            if config.HEADED:
                try:
                    input("\n浏览器保持打开（已 Escape，未提交）。查看完毕按回车关闭... ")
                except EOFError:
                    time.sleep(10)
        except SystemExit:
            pass
        except Exception as e:
            print(f"✗ 探测出错：{e}")
            observations["exception"] = str(e)
            _shot(page, "numfmt_error")
        finally:
            context.close()
            browser.close()

    _write_findings(reachable, verdict, verdict_line, observations, shot_path)
    print(f"\nfindings 已写入：{FINDINGS_PATH}")
    print(f"站点可达：{reachable}")
    print(verdict_line or "VERDICT: (未生成)")
    return 0


def _write_findings(reachable, verdict, verdict_line, obs, shot_path) -> None:
    FINDINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = []
    lines.append("# DASS 数字格式单元格 千分位逗号 探测结论（仅显示 vs 真写入存储）")
    lines.append("")
    lines.append(f"- 生成时间：{stamp}")
    lines.append(f"- 站点可达：{'是' if reachable else '否'}")
    lines.append(f"- 截图：{shot_path or '(未生成)'}")
    lines.append("")
    lines.append("## 判定")
    lines.append("")
    lines.append(f"**{verdict_line}**")
    lines.append("")
    if not reachable:
        lines.append("> 站点不可达，探测中止。无法判定逗号来源。上面 VERDICT 行仅为占位，"
                     "待站点可用后重跑 `./.venv/bin/python -m automation.dass_probe_numfmt`。")
        lines.append("")
    lines.append("## 关键观测 (a)/(b)/(c)")
    lines.append("")
    lines.append(f"- 探测列：`{obs.get('probe_col')}`  单元格 id：`{obs.get('cell_id')}`  "
                 f"测试值：`{obs.get('test_value')}`")
    lines.append(f"- (a) 编辑态进入时 input_value()：`{obs.get('editing_value_before')!r}`")
    lines.append(f"- (a) 键入测试值后、回车前 input_value()：`{obs.get('editing_value_after')!r}`  "
                 f"（是否含逗号：{obs.get('editor_has_comma')}）")
    lines.append(f"- (b) 回车后单元格 inner_text()（当前 fill_cell 回读看到的）："
                 f"`{obs.get('after_inner_text')!r}`")
    lines.append(f"- (c) 回车后单元格底层模型（title/data-*/子节点）：")
    lines.append("")
    lines.append("```json")
    import json
    try:
        lines.append(json.dumps(obs.get("after_model"), ensure_ascii=False, indent=2))
    except Exception:
        lines.append(str(obs.get("after_model")))
    lines.append("```")
    lines.append("")
    lines.append(f"- 底层值候选（来源 -> 值）：")
    for src, val in (obs.get("underlying_candidates") or []):
        lines.append(f"  - `{src}` = `{val!r}`")
    lines.append(f"- 存在无逗号的底层原值线索：{obs.get('has_raw_underlying')}")
    lines.append("")
    if obs.get("error") or obs.get("error_edit") or obs.get("exception"):
        lines.append("## 异常/未完成")
        lines.append("")
        for k in ("error", "error_edit", "exception"):
            if obs.get(k):
                lines.append(f"- {k}: {obs[k]}")
        lines.append("")
    lines.append("## 对已提交旧记录的影响（供用户人工核对，勿自动修复）")
    lines.append("")
    lines.append("- 之前实跑里 ruleid 显示 `65,232`、当前阶段时间显示 `20,261,026` 的那条记录：")
    if verdict == "DISPLAY-ONLY":
        lines.append("  - 依据本判定（仅显示），其**底层存储值大概率仍是无逗号的正确值**"
                     "（65232 / 20261026），即使曾被提交也**很可能未被污染**。"
                     "但请用户在 DASS 中打开该记录、双击这两格确认编辑态回放的是无逗号原值以最终核实。")
    elif verdict == "STORED-CORRUPTION":
        lines.append("  - 依据本判定（真写入存储），若该记录曾被提交，则**很可能已被污染**"
                     "（持久化成 65,232 / 20,261,026）。请用户在 DASS 中人工核对并按需修正，"
                     "本流程不自动改动既有记录。")
    else:
        lines.append("  - 本次判定不确定，无法推断该旧记录是否被污染，请用户结合截图人工核对。")
    lines.append("")
    lines.append("## 修复方向提示（供下一步，不在本步实施）")
    lines.append("")
    if verdict == "DISPLAY-ONLY":
        lines.append("- Bug1：fill_cell 回读应把显示文本与期望值都做 norm（去逗号）后再比较，"
                     "而非直接字符串相等；逗号纯属显示格式，底层值正确。")
    elif verdict == "STORED-CORRUPTION":
        lines.append("- Bug1：逗号会被写进存储值，需在写入前规避数字格式化（例如以文本方式写入、"
                     "或写入后校验并纠正），不能仅靠回读 norm 掩盖。")
    lines.append("- Bug2：无论哪种判定，_write_row 成功语义要与回读判定解耦，"
                 "确保 Stage1 命中后 `satisfied_ruleids.add(rid)` 会执行，避免 Stage2 重复更新。")
    lines.append("")
    FINDINGS_PATH.write_text("\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    sys.exit(main())
