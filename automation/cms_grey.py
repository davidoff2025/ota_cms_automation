"""CMS 灰度放开（`deployota update` 的 CMS 部分）：在「整机升级管理 → 在线规则」里把 Rule ID 放开到目标阶段。

页面结构（2026-11-11 只读实测）：
  · 页面：/basicservice/otamanage/uprule（左侧菜单『在线规则』）。
  · 过滤：设备类型（antd v3 下拉）、规则ID范围（两个框）、批量规则ID（placeholder『多个用英文逗号隔开』）、
    『搜索』按钮。用批量规则ID搜，结果就是这些 Rule ID。
  · 表头：规则ID / 设备类型 / 设备扩展信息 / 源版本 / 目标版本 / 阶段名称 / 累计检测数 / 升级方式 / 备注 / 状态。
  · 批量操作类型下拉：批量灰度 / 批量全量 / 批量正常 / 批量暂停 —— ★ 默认是『批量暂停』。
    只有选了『批量灰度』才出现『配置最大升级检测数』按钮。
  · 『配置最大升级检测数』弹窗：开始时间 / 结束时间（已预填）/ 当前阶段（下拉：测试阶段(5)、50阶段(50)、
    200阶段(200)、1000阶段(1000)、5000阶段(5000)、20000阶段(20000)、全量(-1)）/ 累计检测数（选阶段后自动填）。
  · 『执行批量操作』：对当前搜索结果执行所选批量操作。

流程（每个设备类型代码一组，每组最多 CHUNK 个 Rule ID）：
  1) 选设备类型 → 批量规则ID 填这组 Rule ID → 搜索；核对结果行恰好是这些 Rule ID。
  2) 看『累计检测数』判断当前阶段：等于上一阶段的检测数（ota50 → 5）才放开；
     已等于目标（ota50 → 50）记为『已是目标阶段』；其它值（跳档、更高）不动并列出。
  3) 只用可放开的 Rule ID 重新搜索，再核对一遍结果行。
  4) 批量操作类型选『批量灰度』（并回读确认不是批量暂停）→『配置最大升级检测数』→ 当前阶段选目标阶段 →
     核对累计检测数已自动变成目标值 → 点弹窗『确 定』。
  5) 你输入 yes（或 --yes）后，再次确认批量操作类型仍是『批量灰度』，点『执行批量操作』；
     若弹确认框则在确认框里点确定。
  6) 回读：重新搜索，『累计检测数』都等于目标值才算成功。
"""
from __future__ import annotations

import re
import sys
import time

from . import config
from ._log import color, err, ok, warn
from .login import _shot, login

# 档位 -> (弹窗『当前阶段』选项原文, 目标累计检测数, 上一阶段累计检测数)
CMS_STAGE = {
    "ota50":   ("50阶段(50)",       50,    5),
    "ota200":  ("200阶段(200)",     200,   50),
    "ota1k":   ("1000阶段(1000)",   1000,  200),
    "ota5k":   ("5000阶段(5000)",   5000,  1000),
    "ota20k":  ("20000阶段(20000)", 20000, 5000),
    "ota-all": ("全量(-1)",         -1,    20000),
}
CHUNK = 20          # 每次最多处理的 Rule ID 数（列表一页 20 行，保证搜索结果一屏核对完）
BATCH_TYPE = "批量灰度"


# ======================================================================
# 纯逻辑（--selftest）
# ======================================================================

def parse_count(text) -> int | None:
    """『累计检测数』单元格 → int（去千分位逗号；『全量』/-1 → -1；读不出 → None）。"""
    t = str(text or "").replace(",", "").replace("，", "").strip()
    if t in ("全量", "-1"):
        return -1
    return int(t) if re.fullmatch(r"\d+", t) else None


def classify_counts(rows: list[dict], option: str) -> dict:
    """按累计检测数分类。rows: [{'规则ID','累计检测数'}...]。

    返回 {'eligible': [...], 'already': [...], 'mismatch': [(rid, 当前值文本)...]}。
    eligible = 当前恰好处于上一阶段（可以放开）；already = 已是目标阶段；mismatch = 其它值（不动）。
    """
    _label, target, prev = CMS_STAGE[option]
    out = {"eligible": [], "already": [], "mismatch": []}
    for r in rows:
        rid, raw = str(r.get("规则ID", "")).strip(), r.get("累计检测数", "")
        c = parse_count(raw)
        if c == prev:
            out["eligible"].append(rid)
        elif c == target:
            out["already"].append(rid)
        else:
            out["mismatch"].append((rid, str(raw)))
    return out


def chunks(xs: list, n: int = CHUNK):
    for i in range(0, len(xs), n):
        yield xs[i:i + n]


# ======================================================================
# 页面操作
# ======================================================================

def _form_item(page, label: str):
    return page.locator(".ant-form-item").filter(
        has=page.locator(".ant-form-item-label", has_text=label)).first


def open_online_rules(page) -> None:
    """打开「整机升级管理 → 在线规则」页。"""
    page.goto(config.CMS_UPEVENT_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(2500)
    page.locator(".ant-menu-item", has_text=re.compile(r"^\s*在线规则\s*$")).first.click()
    _form_item(page, "批量规则ID").wait_for(state="visible", timeout=15000)
    page.wait_for_timeout(1500)
    print("  ✓ 已打开『在线规则』页")


def set_device_type(page, code: str) -> bool:
    """在『设备类型』下拉里输入设备类型代码并选中联想项（antd v3）。"""
    sel = _form_item(page, "设备类型").locator(".ant-select").first
    try:
        sel.click()
        page.wait_for_timeout(300)
        inp = sel.locator("input.ant-select-search__field").first
        if inp.count():
            inp.fill("")
            inp.type(code, delay=50)
        else:
            page.keyboard.type(code, delay=50)
        opts = page.locator(".ant-select-dropdown:not(.ant-select-dropdown-hidden) .ant-select-dropdown-menu-item")
        for _ in range(25):
            page.wait_for_timeout(200)
            if opts.count() and opts.first.is_visible():
                break
        for i in range(opts.count()):
            o = opts.nth(i)
            if o.is_visible() and code in (o.inner_text() or ""):
                o.click()
                page.wait_for_timeout(400)
                break
        shown = (sel.inner_text() or "").strip()
        if code in shown:
            print(f"  ✓ 设备类型已选：{shown}")
            return True
    except Exception as e:
        print(err(f"设备类型设置异常：{e}"))
    print(err(f"设备类型 {code!r} 没有选上"))
    return False


def search_ids(page, ids: list[str]) -> list[dict]:
    """批量规则ID 填 ids（英文逗号）→ 搜索 → 等结果只含这些 Rule ID → 返回行。"""
    for box in _form_item(page, "规则ID范围").locator("input").all():
        box.fill("")
    _form_item(page, "批量规则ID").locator("input").first.fill(",".join(ids))
    page.get_by_role("button", name=re.compile(r"搜\s*索")).first.click()  # 按钮名里带图标字，不能锚定首尾
    want = set(ids)
    rows: list[dict] = []
    for _ in range(40):
        page.wait_for_timeout(250)
        rows = read_rows(page)
        got = {r.get("规则ID", "") for r in rows}
        if rows and got <= want and len(got) == len(rows):
            break
    page.wait_for_timeout(400)
    return read_rows(page)


def read_rows(page) -> list[dict]:
    """读结果表：[{表头: 单元格文本}]。

    ★ 该页表头和表体是两张 <table>（固定表头布局）：表头表只有 th、没有数据行；数据行在另一张表里。
      所以表头取含『累计检测数』的那张表的 th，数据行取所有可见表里单元格数与表头一致的 tbody tr。
    """
    return page.evaluate(
        """() => {
            const vis = e => e.offsetParent !== null;
            const ht = [...document.querySelectorAll('table')]
                .find(x => vis(x) && [...x.querySelectorAll('th')].some(h => (h.innerText||'').includes('累计检测数')));
            if (!ht) return [];
            const heads = [...ht.querySelectorAll('th')].map(h => (h.innerText||'').replace(/\\s/g,''));
            const out = [], seen = new Set();
            for (const t of document.querySelectorAll('table')) {
                if (!vis(t)) continue;
                for (const r of t.querySelectorAll('tbody tr')) {
                    if (!vis(r)) continue;
                    const c = [...r.querySelectorAll('td')];
                    if (c.length !== heads.length) continue;
                    const o = {};
                    heads.forEach((h, i) => { if (h) o[h] = (c[i].innerText||'').trim(); });
                    if (o['规则ID'] && !seen.has(o['规则ID'])) { seen.add(o['规则ID']); out.push(o); }
                }
            }
            return out;
        }"""
    ) or []


def _check_exact(rows: list[dict], ids: list[str]) -> str:
    """搜索结果必须恰好是 ids；否则返回问题说明。"""
    got = [r.get("规则ID", "") for r in rows]
    if sorted(got) != sorted(ids):
        extra = sorted(set(got) - set(ids))
        miss = sorted(set(ids) - set(got))
        return f"搜索结果与请求不一致（多出 {extra or '-'}，缺少 {miss or '-'}）"
    return ""


def set_batch_type(page, value: str = BATCH_TYPE) -> bool:
    fi = _form_item(page, "批量操作类型")
    fi.locator(".ant-select").first.click()
    page.wait_for_timeout(500)
    page.locator(".ant-select-dropdown:not(.ant-select-dropdown-hidden) li", has_text=value).first.click()
    page.wait_for_timeout(600)
    return batch_type_is(page, value)


def batch_type_is(page, value: str = BATCH_TYPE) -> bool:
    try:
        return value in (_form_item(page, "批量操作类型").locator(".ant-select").first.inner_text() or "")
    except Exception:
        return False


def configure_stage(page, option: str, dry_run: bool = False) -> tuple[bool, str]:
    """『配置最大升级检测数』→ 当前阶段选目标 → 核对累计检测数 → 点弹窗『确 定』。

    dry_run=True：走到点『确 定』之前就点『取 消』（只用于验证页面，不写入）。
    """
    label, target, _prev = CMS_STAGE[option]
    page.get_by_role("button", name=re.compile("配置最大升级检测数")).first.click()
    m = page.locator(".ant-modal", has_text="最大升级检测数配置").last
    try:
        m.wait_for(state="visible", timeout=8000)
    except Exception:
        return False, "没有弹出『最大升级检测数配置』窗口"
    st = m.locator(".ant-form-item").filter(has=page.locator(".ant-form-item-label", has_text="当前阶段")).first
    st.locator(".ant-select").first.click()
    page.wait_for_timeout(500)
    opt = page.locator(".ant-select-dropdown:not(.ant-select-dropdown-hidden) li").filter(
        has_text=re.compile(rf"^\s*{re.escape(label)}\s*$")).first
    if not opt.count():
        _cancel_modal(m)
        return False, f"『当前阶段』里没有选项 {label!r}"
    opt.click()
    page.wait_for_timeout(700)
    chosen = (st.inner_text() or "").replace("当前阶段", "").strip()
    if label not in chosen:
        _cancel_modal(m)
        return False, f"『当前阶段』选中的是 {chosen!r}，不是 {label!r}"
    counts = m.evaluate("""m => [...m.querySelectorAll('.ant-form-item')]
        .filter(f => (f.querySelector('.ant-form-item-label')?.innerText||'').includes('累计检测数'))
        .map(f => [...f.querySelectorAll('input')].map(i => i.value).join(''))""")
    if str(target) not in [str(c).replace(",", "").strip() for c in counts]:
        _cancel_modal(m)
        return False, f"选 {label} 后累计检测数没有变成 {target}（看到 {counts}）"
    print(f"  ✓ 配置：当前阶段 = {label}，累计检测数 = {target}")
    if dry_run:
        _cancel_modal(m)
        page.wait_for_timeout(600)
        return True, "dry-run：已取消，未点确定"
    m.locator("button", has_text=re.compile(r"^\s*确\s*定\s*$")).first.click()
    try:
        m.wait_for(state="hidden", timeout=8000)
    except Exception:
        return False, "点『确 定』后配置窗口没有关闭（可能有必填项或报错）"
    from .active_ota import _detect_ui_error
    ui = _detect_ui_error(page)
    if ui:
        return False, "CMS 报错：" + " ".join(ui.split())
    return True, ""


def _cancel_modal(m) -> None:
    try:
        m.locator("button", has_text=re.compile(r"^\s*取\s*消\s*$")).first.click()
    except Exception:
        pass


def execute_batch(page) -> tuple[bool, str]:
    """点『执行批量操作』；有确认框就在确认框里点确定。点之前再确认批量操作类型是『批量灰度』。"""
    if not batch_type_is(page):
        return False, "批量操作类型不是『批量灰度』，拒绝执行（默认是批量暂停）"
    page.get_by_role("button", name=re.compile("执行批量操作")).first.click()
    page.wait_for_timeout(1200)
    for sel in (".ant-popover:not(.ant-popover-hidden)", ".ant-modal-confirm", ".ant-popconfirm"):
        box = page.locator(sel).last
        try:
            if box.count() and box.is_visible():
                print(f"  确认框：{' '.join((box.inner_text() or '').split())[:120]}")
                box.locator("button", has_text=re.compile(r"^\s*(确\s*定|确\s*认|是|OK)\s*$")).first.click()
                page.wait_for_timeout(1500)
                break
        except Exception:
            continue
    from .active_ota import _detect_ui_error
    ui = _detect_ui_error(page)
    if ui:
        return False, "CMS 报错：" + " ".join(ui.split())
    return True, ""


# ======================================================================
# 对外入口：dass_update 在成功率闸门通过后调用
# ======================================================================

def run_grey_release(p, option: str, code_of: dict, assume_yes: bool = False,
                     dry_run: bool = False) -> set[str] | None:
    """在 CMS 把 Rule ID 灰度放开到 option 对应阶段。

    code_of: {rule_id: 设备类型代码}。返回成功放开 + 本来已是目标阶段的 Rule ID 集合（供后续 DASS 更新）；
    你取消或页面出错返回 None。

    dry_run=True（测试模式）：照常找规则、核对累计检测数、切到『批量灰度』、打开『配置最大升级检测数』选好目标
    阶段并核对自动填的检测数——然后点『取 消』，不点确定、不点『执行批量操作』。不问 yes，CMS 不做任何修改。
    返回 None（测试模式不进入后续 DASS 更新）。
    """
    from .active_offline import group_by_device
    label, target, prev = CMS_STAGE[option]
    mode = "【测试模式：不执行批量操作，CMS 不做任何修改】" if dry_run else ""
    print(color(f"\n======== CMS 灰度放开：{option} → {label}（累计检测数 {prev} → {target}）{mode}========", "bold"))
    tested: list[str] = []
    done, already, mismatch, failed = [], [], [], []
    browser = p.chromium.launch(headless=not config.HEADED)
    context = browser.new_context(viewport={"width": 1600, "height": 900})
    page = context.new_page()
    try:
        login(page)
        open_online_rules(page)
        groups = group_by_device(code_of)

        # —— 检查（只读）——
        plan: list[tuple[str, list[str]]] = []
        for dev, rids in groups.items():
            if not set_device_type(page, dev):
                failed += [(r, f"设备类型 {dev!r} 选不上") for r in rids]
                continue
            for part in chunks(rids):
                rows = search_ids(page, part)
                problem = _check_exact(rows, part)
                if problem:
                    print(warn(f"  {dev}：{problem}"))
                got = {r["规则ID"] for r in rows}
                failed += [(r, "在线规则里搜不到（不在线？）") for r in part if r not in got]
                rows = [r for r in rows if r["规则ID"] in part]
                for r in rows:
                    print(f"  {r['规则ID']}  {dev}  阶段={r.get('阶段名称', '')}  累计检测数={r.get('累计检测数', '')}")
                c = classify_counts(rows, option)
                already += c["already"]
                mismatch += c["mismatch"]
                if c["eligible"]:
                    plan.append((dev, c["eligible"]))
        if already:
            print(f"  已是目标阶段（累计检测数={target}），不再放开：{', '.join(already)}")
        if mismatch:
            print(warn("累计检测数不是上一阶段的值，不动：" + "，".join(f"{r}={v}" for r, v in mismatch)
                       + f"（{option} 要求当前为 {prev}）"))
        todo = [r for _, ids in plan for r in ids]
        if not todo:
            print(warn("没有需要在 CMS 放开的 Rule ID。"))
            _print_grey_summary(option, done, already, mismatch, failed)
            if already and not dry_run:
                from .stage_db import record_stage
                record_stage(already, option, f"update {option} (CMS)", code_of)
            return None if dry_run else set(already)

        # —— 人工确认（测试模式不问：后面不会写入）——
        print(color(f"\n将在 CMS 把 {len(todo)} 个 Rule ID 灰度放开到 {label}：{', '.join(todo)}", "bold"))
        go = assume_yes or dry_run
        if not go:
            if not sys.stdin.isatty():
                print(err("非交互终端无法确认，已取消（需要无人值守请加 --yes）。"))
                return None
            try:
                go = input("  输入 yes 执行，其它任意输入取消：").strip().lower() == "yes"
            except EOFError:
                go = False
        if not go:
            print("已取消，CMS 未做任何修改。")
            return None

        # —— 执行 ——
        for dev, ids in plan:
            if not set_device_type(page, dev):
                failed += [(r, f"设备类型 {dev!r} 选不上") for r in ids]
                continue
            rows = search_ids(page, ids)
            problem = _check_exact(rows, ids)
            if problem:
                failed += [(r, problem + "，未执行") for r in ids]
                continue
            if not set_batch_type(page):
                failed += [(r, "批量操作类型没有切到『批量灰度』，未执行") for r in ids]
                continue
            ok_cfg, why = configure_stage(page, option, dry_run=dry_run)
            if dry_run:
                if ok_cfg:
                    tested += ids
                    print(ok(f"测试通过（未执行）：{dev} {', '.join(ids)} —— 搜索、批量灰度、"
                             f"配置窗口选 {label}、检测数自动为 {target} 均正常；已在配置窗口点『取 消』"))
                    _shot(page, "grey_dryrun")
                else:
                    _shot(page, "grey_dryrun_fail")
                    failed += [(r, why) for r in ids]
                continue
            if not ok_cfg:
                _shot(page, "grey_config_fail")
                failed += [(r, why) for r in ids]
                continue
            ok_ex, why = execute_batch(page)
            if not ok_ex:
                _shot(page, "grey_execute_fail")
                failed += [(r, why) for r in ids]
                continue
            # —— 回读 ——
            page.wait_for_timeout(1500)
            after = {r["规则ID"]: r for r in search_ids(page, ids)}
            for r in ids:
                c = parse_count(after.get(r, {}).get("累计检测数"))
                if c == target:
                    done.append(r)
                else:
                    failed.append((r, f"执行后累计检测数={after.get(r, {}).get('累计检测数', '?')}，不是 {target}"))
            _shot(page, "grey_done")
        _print_grey_summary(option, done, already, mismatch, failed, tested if dry_run else None)
        if dry_run:
            return None
        if done or already:
            # 阶段记录以 CMS 灰度放开成功（回读累计检测数 = 目标值）为准。
            from .stage_db import record_stage
            record_stage(done + already, option, f"update {option} (CMS)", code_of)
        return set(done) | set(already)
    except Exception as e:
        print(err(f"CMS 灰度放开出错：{e}"))
        _shot(page, "grey_error")
        _print_grey_summary(option, done, already, mismatch, failed)
        return None
    finally:
        context.close()
        browser.close()


def _print_grey_summary(option, done, already, mismatch, failed, tested=None) -> None:
    label, target, _ = CMS_STAGE[option]
    if tested is not None:
        print(color(f"\n—— CMS 灰度放开【测试模式】汇总（{label}）——", "bold"))
        print(ok(f"可以放开、流程已走通但未执行（{len(tested)}）：{', '.join(tested) or '-'}"))
        print(color("  测试模式：配置窗口已取消，未点『执行批量操作』，CMS 没有任何修改。", "yellow"))
    else:
        print(color(f"\n—— CMS 灰度放开汇总（{label}）——", "bold"))
        print(ok(f"放开成功（{len(done)}）：{', '.join(done) or '-'}"))
    print(f"  已是目标阶段（{len(already)}）：{', '.join(already) or '-'}")
    if mismatch:
        print(warn(f"阶段不符未动（{len(mismatch)}）：" + "，".join(f"{r}={v}" for r, v in mismatch)))
    if failed:
        print(err(f"失败（{len(failed)}）："))
        for r, why in failed:
            print(color(f"    {r}：{why}", "red"))


def _selftest() -> int:
    good = True

    def check(name, cond):
        nonlocal good
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        good = good and bool(cond)

    check("检测数解析", parse_count("5") == 5 and parse_count("1,000") == 1000
          and parse_count("全量") == -1 and parse_count("-1") == -1 and parse_count("") is None)
    rows = [{"规则ID": "1", "累计检测数": "5"}, {"规则ID": "2", "累计检测数": "50"},
            {"规则ID": "3", "累计检测数": "200"}, {"规则ID": "4", "累计检测数": ""}]
    c = classify_counts(rows, "ota50")
    check("ota50：5 可放开、50 已是目标、其它不动",
          c["eligible"] == ["1"] and c["already"] == ["2"] and [r for r, _ in c["mismatch"]] == ["3", "4"])
    c2 = classify_counts([{"规则ID": "9", "累计检测数": "20,000"}, {"规则ID": "8", "累计检测数": "-1"}], "ota-all")
    check("ota-all：20000 可放开、-1 已是全量", c2["eligible"] == ["9"] and c2["already"] == ["8"])
    check("跳档不放开（ota1k 但当前 5）", classify_counts([{"规则ID": "7", "累计检测数": "5"}], "ota1k")["eligible"] == [])
    check("档位映射完整", set(CMS_STAGE) == {"ota50", "ota200", "ota1k", "ota5k", "ota20k", "ota-all"})
    check("分块", [len(x) for x in chunks(list(range(45)))] == [20, 20, 5])
    check("搜索结果核对", _check_exact([{"规则ID": "1"}], ["1"]) == ""
          and "多出" in _check_exact([{"规则ID": "1"}, {"规则ID": "2"}], ["1"]))
    print("\n自测结果：", "全部通过 ✅" if good else "存在失败 ❌")
    return 0 if good else 1


if __name__ == "__main__":
    sys.exit(_selftest() if "--selftest" in sys.argv[1:] else 0)
