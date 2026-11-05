"""active_ota：CMS OTA 新功能模块（导入建规则 → 逐条上线并处理冲突 → 回读状态汇总）。

用法：
    cd ota_cms_automation
    ./run_active_ota.sh                      # 取 ota_v4 输出目录里最新的 .xlsx
    ./run_active_ota.sh "/路径/模板.xlsx"      # 指定 Excel 模板
    ./.venv/bin/python -m automation.active_ota --selftest     # 纯逻辑自测（不连浏览器）
    CMS_HEADED=true ./.venv/bin/python -m automation.active_ota --probe-online  # 只读 dump 编排视图

三阶段（已确认决策 D1~D12，硬性）：
  阶段1 导入建规则（D5）：走「OTA升级模板导入」入口，把规则名称(D1)/备注(D2)填入导入弹窗
       输入框后导入，导入后回读本批连续 Rule ID range。
  阶段2 逐条上线 + 冲突处理（D6/D7/D8）：从「升级事件(upevent)」列表行内点『规则上下线』打开
       标题为「规则上下线编排」的双区弹窗，在「待选规则」区按规则ID搜出→点「上线」→确认框点
       「确定」。若弹「规则替换」冲突窗：解析被冲突号 XXXXX→先不点弹窗按钮→去「已选规则」
       按 XXXXX 搜、读 mac组 列→mac组非空→回弹窗点「确定」替换；mac组为空→点「取消」中止该条。
  阶段3 状态汇总（D9/D10）：对每个 Rule ID 回读最终状态（online / 冲突中止 / 失败）并打印。

命名/备注规则：
  D1 规则名称 = 设备类型代码 + "_" + 目标版本最后 5 个字符。
     例 MT9655AMAndrGTV + 0100UQ1015(末5=Q1015) → MT9655AMAndrGTV_Q1015。
  D2 备注 = 今天日期 YYYYMMDD + "_USA_" + <数量>，数量=导入 Excel 数据行数。例 20261102_USA_9。

安全边界（硬性）：
- 「确定/删除/替换」按钮点击严格受 D6/D7 条件约束，条件不满足绝不点击。
- mac组判定但凡不确定（搜不到被冲突规则/读不到 mac组列/列文本歧义）一律保守中止该条并记录，不猜。
- Rule ID 比较统一复用 dass_fill.norm_ruleid（去所有非数字分隔符）。
- 不改动 dass_* / login 既有代码；Excel 读取复用 openpyxl（沿用 dass_data 表头→列索引范式）。

编排视图真实选择器（见 .agents/tasks/active-ota-probe-findings.md §8，实跑校正）：
  · 入口：升级事件页 CMS_UPEVENT_URL 列表行内『规则上下线』按钮 → 标题「规则上下线编排」的 antd 弹窗。
  · 双区根容器：弹窗内有且仅 2 张带表头 .vxe-table，**待选表独有 class `vxe-editable`**
    （已选表无此 class）；已选=第 1 张 .vxe-table。区 class 为空、vxe `tid_*` 动态，勿硬编码。
  · 两区搜索框 #ruleId（精准）/#deviceExt（模糊）各出现 2 次（两区各一、id 重复），必须先锚区再取。
  · 列序：规则ID=索引0、mac组=索引5（两区一致）；按列头文本定位列索引更稳。
⚠ 仍**未实跑复现**（环境两区 No Data 所致，见 findings §8.4/8.5/8.8）：行级『上线/下线/编辑』按钮
  文本（是否全角空格）、「确认上线该规则?」小确认框、「规则替换」冲突窗、真实 mac组 单元格占位文案。
  这些按**用户截图/DOM 推断**填入并在注释标注「未实跑复现」，所有写操作均经二次确认 gate（D-f），
  条件不满足只打印不点，确保不误点危险按钮。首次实跑前可 `--probe-online` 再校正这几项。
"""
from __future__ import annotations

import re
import sys
import time
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

from . import config
from .login import login, _shot
# 复用 DASS 侧常量/工具，绝不修改其实现：
from .dass_fill import norm_ruleid, OTA_V4_OUTPUTS
from .dass_data import latest_updated_excel

# 导入 Excel 里取设备类型代码/目标版本的列名（与 dass_data.EXCEL_TO_DASS 同名，含容错别名）。
DEVICE_CODE_HEADERS = ["设备类型代码"]
TARGET_VERSION_HEADERS = ["目标完整版本", "目标版本"]
RULEID_HEADERS = ["Rule ID", "ruleid", "RuleID"]

# mac组 判定为「空」的占位文本（去空白后命中任一即视为未绑定）。
MAC_GROUP_EMPTY_TOKENS = {"", "无", "null", "none", "-", "—", "空"}


# ======================================================================
# 纯逻辑函数（不连浏览器，--selftest 覆盖全部分支）
# ======================================================================

def build_rule_name(device_code: str, target_version: str) -> str:
    """D1：规则名称 = 设备类型代码 + "_" + 目标版本最后 5 个字符。

    例：build_rule_name("MT9655AMAndrGTV", "0100UQ1015") -> "MT9655AMAndrGTV_Q1015"。
    目标版本不足 5 位时取其全部（[-5:] 的自然语义）。两端去空白。
    """
    dc = str(device_code or "").strip()
    tv = str(target_version or "").strip()
    return f"{dc}_{tv[-5:]}"


def build_remark(count: int, today: str | None = None) -> str:
    """D2：备注 = 今天日期 YYYYMMDD + "_USA_" + 数量。

    today 为 None 时用今日日期（YYYYMMDD）；传入字符串则原样用（便于 selftest）。
    例：build_remark(9, "20261102") -> "20261102_USA_9"。
    """
    day = today if today is not None else datetime.now().strftime("%Y%m%d")
    return f"{day}_USA_{int(count)}"


def parse_conflict_ruleid(modal_text: str) -> str:
    """D7/D-d：从「规则替换」弹窗文案解析被冲突的 Rule ID，再经 norm_ruleid 归一化。

    文案形如「当前规则和 65246 冲突,是否替换?」。先用正则取『和 … 冲突』之间的数字，
    再过 norm_ruleid 去所有分隔符，与 dass_fill 口径一致。无匹配返回空串。
    """
    text = str(modal_text or "")
    m = re.search(r"和\s*([0-9,，\s]+?)\s*冲突", text)
    if not m:
        return ""
    return norm_ruleid(m.group(1))


def mac_group_is_bound(cell_text: str) -> bool:
    """D7/D-e：判断「mac组」列单元格是否绑定（非空）。

    去空白后，命中 MAC_GROUP_EMPTY_TOKENS（空/无/null/none/-/—/空）视为未绑定 → False；
    否则视为已绑定 → True。判定标准保守：拿不到文本或为占位符一律 False（不猜着替换）。
    """
    t = str(cell_text or "").strip()
    return t.lower() not in MAC_GROUP_EMPTY_TOKENS


def validate_single_target(device_codes, target_versions) -> tuple[str, str]:
    """D3：校验导入 Excel 只含单一设备类型代码 + 单一目标版本。

    任一集合大小 != 1 → 抛带中文提示的 ValueError（列出实际多个值，不猜）。
    恰好单一 → 返回 (device_code, target_version)。
    """
    dcs = sorted(x for x in device_codes if str(x or "").strip())
    tvs = sorted(x for x in target_versions if str(x or "").strip())
    if len(dcs) != 1:
        raise ValueError(
            f"导入 Excel 的『设备类型代码』应为单一值，实际为 {dcs!r}（共 {len(dcs)} 个）。"
            "请检查模板，不混多个设备类型。"
        )
    if len(tvs) != 1:
        raise ValueError(
            f"导入 Excel 的『目标版本』应为单一值，实际为 {tvs!r}（共 {len(tvs)} 个）。"
            "请检查模板，不混多个目标版本。"
        )
    return dcs[0], tvs[0]


# 阶段3 的每个 Rule ID 最终状态常量。
STATUS_ONLINE = "online"
STATUS_CONFLICT_ABORT = "conflict_abort"
STATUS_FAILED = "failed"


def summarize(results: list[dict]) -> str:
    """D10：把每个 Rule ID 的最终状态格式化成可检索的汇总文本（纯逻辑）。

    results 每项为 dict：
      {"ruleid": str, "status": STATUS_*, "conflict_ruleid"?: str,
       "mac_group_bound"?: bool, "note"?: str}
    输出每行一个 Rule ID 的状态，末尾附三态计数汇总。风格对齐 dass_fill.format_update_detail。
    """
    lines = ["======== active_ota 上线汇总（D10）========"]
    n_online = n_conflict = n_failed = 0
    for r in results:
        rid = r.get("ruleid", "")
        st = r.get("status", "")
        if st == STATUS_ONLINE:
            n_online += 1
            lines.append(f"  [汇总] Rule ID={rid!r}  状态=online ✅（已上线）")
        elif st == STATUS_CONFLICT_ABORT:
            n_conflict += 1
            cf = r.get("conflict_ruleid", "")
            bound = r.get("mac_group_bound", None)
            mac_desc = ("mac组非空" if bound else "mac组为空") if bound is not None else "mac组判定不确定"
            lines.append(
                f"  [汇总] Rule ID={rid!r}  状态=冲突中止 ⛔（被冲突号={cf!r}，{mac_desc}）"
            )
        else:
            n_failed += 1
            note = r.get("note", "")
            lines.append(f"  [汇总] Rule ID={rid!r}  状态=失败 ❌" + (f"（{note}）" if note else ""))
    lines.append(
        f"  —— 总数：{len(results)}  online：{n_online}  "
        f"冲突中止：{n_conflict}  失败：{n_failed}"
    )
    return "\n".join(lines)


def summary_exit_code(results: list[dict]) -> int:
    """全部 online -> 0；有冲突中止/失败 -> 非 0（供 main 退出码，D10）。"""
    if results and all(r.get("status") == STATUS_ONLINE for r in results):
        return 0
    return 1


# ======================================================================
# Excel 读取（轻量，供阶段1；复用 openpyxl，沿用 dass_data 表头→列索引范式）
# ======================================================================

def _pick_header(col_index: dict, candidates: list[str]) -> int | None:
    """从表头->列索引映射里取第一个存在的候选列的索引，取不到返回 None。"""
    for name in candidates:
        if name in col_index:
            return col_index[name]
    return None


def read_import_excel(path: str, sheet_name: str = "升级规则") -> dict:
    """读导入 Excel，返回 {rows, device_codes, target_versions, rule_ids, count}。

    用 openpyxl 按 dass_data 的「表头→列索引」范式读每行的 设备类型代码 / 目标完整版本
    （+Rule ID 若有列则读，用于后续校验），不引入新依赖。rows 为每行原始字段 dict。
    """
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"导入 Excel 文件不存在：{path}")

    wb = load_workbook(str(p), data_only=True)
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.worksheets[0]

    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    col_index = {h: i for i, h in enumerate(headers)}
    dc_idx = _pick_header(col_index, DEVICE_CODE_HEADERS)
    tv_idx = _pick_header(col_index, TARGET_VERSION_HEADERS)
    rid_idx = _pick_header(col_index, RULEID_HEADERS)

    rows: list[dict] = []
    device_codes: set[str] = set()
    target_versions: set[str] = set()
    rule_ids: list[str] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(v not in (None, "") for v in row):
            continue

        def _cell(idx):
            if idx is None or idx >= len(row):
                return ""
            v = row[idx]
            return "" if v is None else str(v).strip()

        dc = _cell(dc_idx)
        tv = _cell(tv_idx)
        rid = _cell(rid_idx)
        rec = {"设备类型代码": dc, "目标版本": tv, "ruleid": rid}
        rows.append(rec)
        if dc:
            device_codes.add(dc)
        if tv:
            target_versions.add(tv)
        if rid:
            rule_ids.append(norm_ruleid(rid))
    wb.close()

    return {
        "rows": rows,
        "device_codes": device_codes,
        "target_versions": target_versions,
        "rule_ids": rule_ids,
        "count": len(rows),
    }


# ======================================================================
# 浏览器驱动工具（文本+role 定位，容忍全角空格；弹窗先取 .ant-modal 作用域）
# ======================================================================

def _normspace(s: str) -> str:
    """去掉半角/全角空格，便于匹配 antd 带全角空格的按钮文本（如『确 定』）。"""
    return str(s or "").replace(" ", "").replace("\u3000", "")


def _active_modal(page):
    """返回当前可见的 antd 弹窗作用域 locator（取最后一个可见的 .ant-modal）。"""
    modals = page.locator(".ant-modal")
    for i in range(modals.count() - 1, -1, -1):
        m = modals.nth(i)
        try:
            if m.is_visible():
                return m
        except Exception:
            continue
    return None


def _click_button_in(scope, page, keyword: str, label: str) -> bool:
    """在 scope 内点击文本（去空白后）包含 keyword 的可见按钮。成功返回 True，否则 False。

    容忍 antd 全角空格；不回退到其它按钮；找不到/不可见时打印并返回 False。
    """
    btns = scope.locator("button")
    for i in range(btns.count()):
        el = btns.nth(i)
        try:
            if not el.is_visible():
                continue
            if keyword in _normspace(el.inner_text()):
                el.scroll_into_view_if_needed(timeout=2000)
                el.click()
                print(f"  ✓ 已点击『{label}』按钮（匹配 {keyword!r}）")
                page.wait_for_timeout(600)
                return True
        except Exception:
            continue
    print(f"  ✗ 未找到可见的『{label}』按钮（关键字 {keyword!r}）。")
    return False


# —— 「规则上下线编排」双区弹窗：导航 + 区作用域（findings §8，实跑校正）——

def open_orchestration_modal(page):
    """导航升级事件页并点行内『规则上下线』，打开标题为「规则上下线编排」的双区弹窗。

    返回弹窗 locator（.ant-modal，含文本「规则上下线编排」）。失败抛 RuntimeError。
    入口与判定信号均为 findings §8.0 实跑校正：
      - URL = config.CMS_UPEVENT_URL（/basicservice/otamanage/upevent）。
      - 行内按钮文本 `规则上下线`（无全角空格），`get_by_role("button", name="规则上下线")`。
      - 弹窗判定：`.ant-modal-title has_text="规则上下线编排"` 可见。
    注意：旧 /otamanage/uprule 单表页不是该视图（findings §8.0 已作废原 §2.1 结论）。
    """
    print(f"  打开升级事件页：{config.CMS_UPEVENT_URL}")
    page.goto(config.CMS_UPEVENT_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(3000)
    try:
        page.get_by_role("button", name="规则上下线").first.click()
    except Exception as e:
        raise RuntimeError(f"未能点开升级事件行内『规则上下线』按钮：{e}")
    try:
        page.locator(".ant-modal-title", has_text="规则上下线编排").first.wait_for(
            state="visible", timeout=10000
        )
    except Exception as e:
        raise RuntimeError(f"点『规则上下线』后未出现「规则上下线编排」弹窗：{e}")
    page.wait_for_timeout(1200)
    modal = page.locator(".ant-modal", has_text="规则上下线编排").last
    print("  ✓ 已打开「规则上下线编排」双区弹窗。")
    return modal


def _orchestration_modal(page):
    """返回当前可见的「规则上下线编排」弹窗 locator；不在场返回 None。"""
    modal = page.locator(".ant-modal", has_text="规则上下线编排").last
    try:
        return modal if modal.count() and modal.is_visible() else None
    except Exception:
        return None


def _pending_table(modal):
    """待选规则表 = 弹窗内独有 class `vxe-editable` 的 .vxe-table（findings §8.1）。"""
    return modal.locator(".vxe-table.vxe-editable").first


def _selected_table(modal):
    """已选规则表 = 弹窗内第 1 张带表头的 .vxe-table（findings §8.1）。"""
    return modal.locator(".vxe-table").first


def _pending_form(modal):
    """待选规则搜索表单 = 弹窗内 form.ant-form-inline（findings §8.1 block 2）。"""
    return modal.locator("form.ant-form-inline").first


def _search_pending_zone(modal, page, rule_id: str) -> bool:
    """在「待选规则」区按规则ID精确搜索（findings §8.2）。

    待选搜索表单 = form.ant-form-inline（block 2），区内 #ruleId placeholder=精准匹配；
    搜索按钮文本无全角空格。先锚待选表单再取 #ruleId，避免与已选区同名控件混淆。
    （设备类型不在此设置——它已在进页面时于已选区设过一次。）
    """
    form = _pending_form(modal)
    form.locator("input#ruleId").first.fill(norm_ruleid(rule_id))
    return _click_button_in(form, page, "搜索", "待选规则-搜索")


def _set_device_type_once(page, modal, device_code: str) -> bool:
    """进页面后【只调用一次】：在「已选规则」区的『设备类型』框设置设备类型代码。

    用户确认的交互：在『设备类型』框【输入】设备类型代码 → 页面自动弹出匹配下拉 → 选中该项
    （不是从静态列表里挑；是输入触发自动匹配）。设备类型只设这一次，后续逐条只用 Rule ID 搜。
    成功返回 True；失败返回 False（上层据此终止整个 online）。

    ★ 已按 findings §9（HEADED 实跑校正）重写 —— 根因：设备类型是 **antd v3 ant-select**，
      旧代码用的是 antd v4 选择器，全部对不上（§9.0）。校正要点：
      · 控件：v3 `.ant-select`（已选区 enabled / 待选区 disabled）；输入框 `input.ant-select-search__field`
        （无 id/placeholder，未展开时 visible:false，必须先 click 容器聚焦）——不是 v4 的
        `ant-select-selection-search-input`（§9.1）。
      · 触发联想：先 `click` 已选区 `.ant-select` 容器聚焦，再 `page.keyboard.type(want)`（§9.2）。
      · 浮层：`.ant-select-dropdown`（**挂在 body 下、不在 modal 内，必须全局 page.locator 取**），
        选项 `.ant-select-dropdown-menu-item`（**v3，不是 v4 的 `.ant-select-item-option`**）（§9.3）。
      · 选中：候选项文字是 `代码（中文名）`（≠ 设备类型代码），必须【包含匹配】或回车选中 `-active`
        高亮项，不能用精确相等（§9.4）。
      · 回读校验：读 v3 `.ant-select-selection-selected-value`（v4 的 `.ant-select-selection-item`
        在此控件取不到）（§9.4）。

    作用域为「已选规则」区：弹窗内含『设备类型』的 `.ant-form-item` 有 2 个，已选区为可编辑
    （`.ant-select` 非 disabled），待选区为 `.ant-select-disabled`（自动带值只读，不设它）（§9.1）。
    """
    want = str(device_code).strip()

    # a) 定位「已选规则」区（可编辑、非 disabled）的『设备类型』ant-select 触发器（§9.1）。
    trigger = None
    try:
        # 优先：含『设备类型』的 form-item 里取非 disabled 的 .ant-select（已选区）。
        cand = modal.locator(
            ".ant-form-item:has-text('设备类型') .ant-select:not(.ant-select-disabled)"
        ).first
        if cand.count():
            trigger = cand
    except Exception:
        trigger = None
    if trigger is None:
        try:
            # 兜底：含『设备类型』的 form-item 里的第 1 个 .ant-select（DOM 顺序第 1 个=已选区）。
            item = modal.locator(".ant-form-item", has_text="设备类型").first
            if item.count():
                cand = item.locator(".ant-select").first
                if cand.count():
                    trigger = cand
        except Exception:
            trigger = None
    if trigger is None:
        print("  ✗ 未定位到「已选规则」区（可编辑）的『设备类型』控件。")
        return False

    def _verify_selected() -> bool:
        """回读该 v3 ant-select 的显示值，确认已选中 want（避免只输文本未选中导致『请输入设备类型』）。"""
        try:
            page.wait_for_timeout(300)
            # v3 显示值容器：.ant-select-selection-selected-value；文字是 `代码（中文名）`，用包含匹配。
            sel = trigger.locator(".ant-select-selection-selected-value").first
            if sel.count():
                txt = (sel.inner_text() or "").strip()
                if txt and want in txt:
                    return True
            # 兜底：读 .ant-select-selection 整体可见文本（覆盖 title 属性里带值的情况）。
            whole = (trigger.inner_text() or "").strip()
            return bool(whole) and want in whole
        except Exception:
            return False

    try:
        trigger.scroll_into_view_if_needed(timeout=2000)
        # b) 先 click 外层容器聚焦并展开下拉（v3 搜索框展开前不可见，必须先聚焦容器）。
        trigger.click()
        page.wait_for_timeout(300)

        # c) 逐字键入设备类型代码触发联想。优先对 v3 搜索框 input 输入；取不到则全局键盘输入（§9.2）。
        search_input = trigger.locator("input.ant-select-search__field").first
        typed_via_input = False
        try:
            if search_input.count():
                search_input.fill("")
                search_input.type(want, delay=60)
                typed_via_input = True
        except Exception:
            typed_via_input = False
        if not typed_via_input:
            page.keyboard.type(want, delay=60)

        # d) 轮询等 v3 联想浮层出现并有可见选项（浮层在 body 下，全局取；选项用 v3 class）（§9.3）。
        dropdown = page.locator(".ant-select-dropdown")
        options = dropdown.locator(".ant-select-dropdown-menu-item")
        appeared = False
        for _ in range(25):
            page.wait_for_timeout(200)
            try:
                if options.count() > 0 and options.first.is_visible():
                    appeared = True
                    break
            except Exception:
                continue

        if appeared:
            n = options.count()
            # d1) 包含匹配点击（候选项文字是 `代码（中文名）`，不能用精确相等）（§9.4）。
            for i in range(n):
                el = options.nth(i)
                try:
                    if not el.is_visible():
                        continue
                    t = (el.inner_text() or "").strip()
                    if want in t:
                        el.click()
                        if _verify_selected():
                            print(f"  ✓ 已选规则区『设备类型』已设置（候选项『{t}』包含匹配）：{want!r}")
                            return True
                except Exception:
                    continue

        # d2) 回车兜底：唯一/首个匹配项通常已带 `-active` 高亮，回车即选中（§9.4）。
        try:
            page.keyboard.press("Enter")
            if _verify_selected():
                print(f"  ✓ 已选规则区『设备类型』已设置（回车选中高亮项）：{want!r}")
                return True
        except Exception:
            pass

        print(f"  ✗ 『设备类型』输入 {want!r} 后未能选中自动匹配项"
              f"（联想浮层出现={appeared}）。若反复失败请 --probe-online 复核 §9 选择器。")
    except Exception as e:
        print(f"  ✗ 『设备类型』设置交互异常：{e}")
    return False


def _search_selected_zone(modal, page, rule_id: str) -> bool:
    """在「已选规则」区按规则ID精确搜索（findings §8.2）。

    findings §8.1：已选区搜索表单与已选表同在弹窗第 0 块，是弹窗内**第 1 个**搜索表单
    （待选表单 form.ant-form-inline 在其后）。#ruleId 两区重复，用「第 1 个 #ruleId」锚已选区，
    并点其所在表单的『搜索』。
    """
    # 已选区 #ruleId = 弹窗内第 1 个（DOM 顺序先于待选 form.ant-form-inline）。
    box = modal.locator("input#ruleId").first
    box.fill(norm_ruleid(rule_id))
    # 已选区搜索按钮 = 弹窗内第 1 个『搜索』（去空白匹配，无全角空格）。
    btns = modal.locator("button")
    for i in range(btns.count()):
        el = btns.nth(i)
        try:
            if not el.is_visible():
                continue
            if "搜索" in _normspace(el.inner_text()):
                el.scroll_into_view_if_needed(timeout=2000)
                el.click()
                print("  ✓ 已点击『已选规则-搜索』按钮（弹窗内第 1 个搜索）")
                page.wait_for_timeout(600)
                return True
        except Exception:
            continue
    print("  ✗ 未找到可见的『已选规则-搜索』按钮。")
    return False


def _read_mac_group_from_table(table) -> str | None:
    """按列头文本『mac组』定位列索引，读该表 tbody 首行该列 innerText（findings §8.3）。

    读不到（无表/无表头/无行/无 mac组 列）返回 None。规则ID 列=索引0、mac组 列=索引5，
    但按列头文本匹配而非硬编码数字，防列序调整。
    """
    try:
        return table.evaluate(
            """(t) => {
                const heads = [...t.querySelectorAll('thead th, thead td')]
                    .map(h => (h.innerText || '').trim());
                const i = heads.findIndex(x =>
                    x.replace(/\\s/g, '').toLowerCase().includes('mac组'.toLowerCase())
                    || x.replace(/\\s/g, '').toLowerCase().includes('mac'));
                const row = t.querySelector('tbody tr');
                if (!row || i < 0) return null;
                const c = row.querySelectorAll('td, th')[i];
                return c ? (c.innerText || '').trim() : null;
            }"""
        )
    except Exception:
        return None


# —— 阶段1：导入建规则 ——

# UI 错误提示：导入点『下一步』或『保存』后，后端可能弹业务错误（实测如
# 『第2行：目标版本已经存在Q0723_to_Q1028, devicecode: MT9655AMAndrGTV』『Service error』）。
# 一旦出现这类可见错误提示，必须立即终止，不再继续向导/保存，并把原文记入错误。
_UI_ERROR_SELECTORS = [
    ".ant-message-error",                 # 顶部红色 message toast
    ".ant-message-custom-content.ant-message-error",
    ".ant-notification-notice-error",     # 右上角错误通知
    ".ant-notice-error",
    ".ant-form-item-explain-error",       # 字段下方红字（如“模板名称不能为空”）
]
# 关键错误词（命中任一即视为错误，覆盖 CMS 自定义样式/文案）。
_UI_ERROR_KEYWORDS = ["Service error", "服务错误", "已经存在", "失败", "错误"]


def _detect_ui_error(page) -> str | None:
    """检测页面上是否出现可见的错误提示。命中返回错误原文（去空白），否则 None。

    先查 ant-design 标准错误节点；再对错误 toast/通知容器里的可见文本做关键词匹配，
    以覆盖自定义样式。只取可见元素，避免命中残留的隐藏节点。
    """
    # 1) 标准错误节点。
    for sel in _UI_ERROR_SELECTORS:
        try:
            loc = page.locator(sel)
            for i in range(min(loc.count(), 10)):
                el = loc.nth(i)
                if el.is_visible():
                    txt = (el.inner_text() or "").strip()
                    if txt:
                        return txt
        except Exception:
            continue
    # 2) 消息/通知容器里的关键词匹配（覆盖非标准 error class 的提示）。
    for container in (".ant-message", ".ant-notification", ".ant-modal-confirm-body"):
        try:
            loc = page.locator(container)
            for i in range(min(loc.count(), 10)):
                el = loc.nth(i)
                if not el.is_visible():
                    continue
                txt = (el.inner_text() or "").strip()
                if txt and any(kw in txt for kw in _UI_ERROR_KEYWORDS):
                    return txt
        except Exception:
            continue
    return None


def _abort_if_ui_error(page, where: str) -> None:
    """在关键写操作后调用：若检测到 UI 错误提示，截图 + 打印原文并抛异常终止。"""
    err = _detect_ui_error(page)
    if err:
        _shot(page, "active_ui_error")
        # 错误文案可能多行，压成单行便于日志阅读。
        one_line = " ".join(err.split())
        raise RuntimeError(f"检测到 UI 错误提示（{where}），已终止：{one_line}")


def _fill_import_field(modal, selectors: list[str], value: str, label: str) -> bool:
    """在导入弹窗作用域内，按候选选择器依次尝试定位可见输入框并填入 value。

    成功填入并回读校验一致后返回 True；全部候选都未命中/回读不符则返回 False。
    回读校验确保『填了但没生效』（如组件未绑定）也算失败，交由调用方中止，避免静默导入。
    """
    for sel in selectors:
        try:
            loc = modal.locator(sel).first
            if loc.count() == 0 or not loc.is_visible():
                continue
            loc.click()
            loc.fill("")
            loc.fill(str(value))
            got = (loc.input_value() or "").strip()
            if got == str(value).strip():
                print(f"  {label} 已填入：{value!r}（选择器 {sel!r}）")
                return True
            print(f"  ⚠ {label} 回读不符：期望 {value!r} 实得 {got!r}（选择器 {sel!r}），继续尝试下一个。")
        except Exception as e:
            print(f"  ⚠ {label} 用选择器 {sel!r} 填写异常：{e}，继续尝试下一个。")
            continue
    return False


def _click_import_confirm(page, modal) -> bool:
    """D-f gate：走完导入【两步向导】。点前断言确实处于导入弹窗。

    实测（用户截图）导入弹窗第一步右下角是『下一步/取消』，点『下一步』后才进入第二步，
    第二步再点完成按钮（『保存』等）导入才真正完成。所以不能点中第一个按钮就返回——
    那正是之前备注看似填了、却没真正完成导入/未点保存的原因。

    流程：
      1) 第一步点『下一步』（若存在）；若没有『下一步』而直接是单步完成按钮，则退化为单步。
      2) 等第二步渲染，点『保存/确定/完成/导入』完成导入。
    条件不满足（弹窗不在场）只打印不点并返回 False。
    """
    if modal is None or not modal.is_visible():
        print("  ✗ [gate] 导入弹窗不在场，拒绝点击导入确认按钮。")
        return False

    # —— 第一步：优先点『下一步』进入第二步；没有『下一步』则视为单步向导。——
    went_next = _click_button_in(modal, page, "下一步", "导入向导第一步(下一步)")
    if went_next:
        page.wait_for_timeout(1800)
        # 点『下一步』后后端可能弹业务错误（目标版本已存在 / Service error）——立即终止。
        _abort_if_ui_error(page, "点『下一步』后")
        modal = _active_modal(page) or modal  # 第二步可能换了 modal 实例
        if modal is None or not modal.is_visible():
            print("  ✗ [gate] 点『下一步』后导入弹窗不在场，拒绝继续。")
            return False

    # —— 第二步（或单步）：点完成/保存按钮真正完成导入。——
    for kw in ("保存", "完成", "确定", "导入"):
        if _click_button_in(modal, page, kw, f"导入完成({kw})"):
            page.wait_for_timeout(1500)
            # 点『保存』后同样可能弹业务错误——立即终止。
            _abort_if_ui_error(page, f"点『{kw}』后")
            return True

    # 单步向导里『下一步』本身可能就是提交（无第二步按钮时）。
    if went_next:
        print("  ⚠ 已点『下一步』但未找到第二步的『保存/完成/确定』按钮——"
              "请确认导入是否已完成（可能第二步按钮文案不同，需按实跑日志校正）。")
        return True

    print("  ✗ 未找到任何导入完成按钮（下一步/保存/完成/确定/导入均未命中）。")
    return False


def run_import(page, excel_path: str) -> tuple[list[str], str, str]:
    """阶段1：导入建规则。返回 (rule_ids_desc_list, name, remark)。

    读 Excel → D3 校验单一值 → 算名称(D1)/备注(D2) → 导航 OTA 模板页、点『导入模板』打开弹窗 →
    在 .ant-modal 作用域内：set_input_files 喂 Excel、填名称(placeholder=必填)、填备注(textarea) →
    走向导（D-f gate）→ 截图。导入后进入『查看规则』按 设备类型代码+目标版本 读 规则ID 列，取本批
    连续区间 max..min（含翻页），断言行数==count 否则告警。
    """
    info = read_import_excel(excel_path)
    count = info["count"]
    if count == 0:
        raise ValueError("导入 Excel 没有数据行。")
    device_code, target_version = validate_single_target(
        info["device_codes"], info["target_versions"]
    )
    name = build_rule_name(device_code, target_version)
    remark = build_remark(count)
    print(f"阶段1 导入：设备类型代码={device_code!r} 目标版本={target_version!r} 行数={count}")
    print(f"  规则名称(D1)={name!r}  备注(D2)={remark!r}")

    # 导航 OTA 模板页并打开导入弹窗（findings §4 实跑确认的选择器）。
    print(f"  打开 OTA 升级模板页：{config.CMS_OTA_URL}")
    page.goto(config.CMS_OTA_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(2500)
    try:
        page.get_by_role("button", name="导入模板").first.click()
    except Exception as e:
        raise RuntimeError(f"未能点开『导入模板』按钮：{e}")
    page.wait_for_timeout(1500)

    modal = _active_modal(page)
    if modal is None:
        raise RuntimeError("点击『导入模板』后未出现导入弹窗（.ant-modal）。")

    # 文件上传（隐藏 input[type=file]，直接喂路径）。
    file_input = modal.locator("input[type='file']").first
    file_input.set_input_files(str(Path(excel_path).resolve()))
    page.wait_for_timeout(1500)

    # 模板名称（placeholder=必填）与备注（textarea）。
    # 模板名称（placeholder=必填）。填不进直接中止——名称必填，否则导入无意义。
    if not _fill_import_field(
        modal,
        [
            "input[placeholder='必填']",
        ],
        name,
        "模板名称",
    ):
        raise RuntimeError("模板名称输入框填写失败（未命中输入框）。")

    # 备注：实测是普通 input（非 textarea），旁边有『备注』标签。按多候选稳健定位。
    # 填不进直接中止——否则会像之前那样静默导入出一个备注为空的模板。
    if not _fill_import_field(
        modal,
        [
            "input[placeholder='备注']",
            "textarea[placeholder='备注']",
            # 兜底：所有非『必填』的可见文本输入框里取第一个（名称框 placeholder=必填，已排除）。
            "input:not([placeholder='必填'])[type='text']",
            "input:not([placeholder='必填']):not([type='file'])",
            "textarea",
        ],
        remark,
        "备注",
    ):
        raise RuntimeError(f"备注输入框填写失败（未命中输入框）。备注值={remark!r}")

    _shot(page, "active_import_dialog")

    # 走向导/完成导入（D-f gate）。
    if not _click_import_confirm(page, _active_modal(page)):
        raise RuntimeError("导入确认按钮点击失败（gate 未满足或按钮缺失）。")
    page.wait_for_timeout(2500)

    # 回读本批 Rule ID：按模板名称在「OTA升级模板」列表定位本批行 → 点「查看规则」→ 读规则ID。
    rule_ids = _collect_rule_ids(page, name, count)
    _shot(page, "active_view_rules")
    return rule_ids, name, remark


def _collect_rule_ids(page, template_name: str, expected: int) -> list[str]:
    """导入后回读本批 Rule ID（升序，数量须等于 Excel 去表头的数据行数）。

    真实路径（用户截图确认）：
      1) 左侧菜单「OTA升级模板」进入模板列表（序号/标题/创建人/备注/导入时间/操作）。
      2) 顶部搜索框按标题搜 template_name 缩小范围，再在「标题」列【精确匹配】本批那一行。
      3) 点该行「查看规则」弹出规则页，读『规则ID』列【全部】ID。
      4) 升序（小→大）返回，并断言数量 == expected（Excel 去表头行数）；不符则告警。
    """
    # —— 1) 进「OTA升级模板」列表（左侧菜单项）——
    if not _open_template_list(page):
        raise RuntimeError("未能进入「OTA升级模板」模板列表页。")

    # —— 2) 顶部搜索框按标题搜，再精确匹配定位本批行，点其「查看规则」——
    if not _open_view_rules_for_template(page, template_name):
        raise RuntimeError(f"未能在模板列表按标题精确匹配到本批行并点『查看规则』：{template_name!r}")

    # —— 3) 读规则页『规则ID』列全部 ID（含翻页）——
    seen: list[str] = []
    page_guard = 0
    while page_guard < 50:
        page_guard += 1
        batch = page.evaluate(
            """() => {
                // vxe-table 表头->列索引，按行读『规则ID』列。
                const tables = Array.from(document.querySelectorAll('.vxe-table, table'));
                const out = [];
                for (const t of tables) {
                    const heads = t.querySelectorAll('thead th, thead td');
                    if (!heads.length) continue;
                    const names = Array.from(heads).map(h => (h.innerText||'').trim());
                    const idxRule = names.findIndex(x => x.indexOf('规则ID') >= 0);
                    if (idxRule < 0) continue;
                    const rows = t.querySelectorAll('tbody tr');
                    for (const r of rows) {
                        const cells = r.querySelectorAll('td, th');
                        if (!cells.length) continue;
                        const rid = cells[idxRule] ? (cells[idxRule].innerText||'').trim() : '';
                        if (rid) out.push(rid);
                    }
                }
                return out;
            }"""
        )
        for raw in batch:
            rid = norm_ruleid(raw)
            if rid and rid not in seen:
                seen.append(rid)

        if len(seen) >= expected:
            break
        if not _goto_next_page(page):
            break
        page.wait_for_timeout(1200)

    # —— 4) 升序（小→大）返回 + 数量校验 ——
    seen_sorted = sorted(set(seen), key=lambda x: int(x) if x.isdigit() else 0)
    if len(seen_sorted) != expected:
        print(f"  ⚠ 回收到 {len(seen_sorted)} 个 Rule ID，期望 {expected}（= Excel 去表头行数）。"
              f"数量不符，请核对是否定位错模板行/翻页不全：{seen_sorted}")
    else:
        print(f"  ✓ 回收本批 Rule ID（{len(seen_sorted)} 个，升序）：{seen_sorted}")
    return seen_sorted


def _open_template_list(page) -> bool:
    """点左侧菜单「OTA升级模板」进入模板列表页。返回是否成功进入。

    用户截图确认入口为左侧导航项「OTA升级模板」（整机升级管理分组下）。
    """
    for sel in (
        "li:has-text('OTA升级模板')",
        ".ant-menu-item:has-text('OTA升级模板')",
        "span:has-text('OTA升级模板')",
        "a:has-text('OTA升级模板')",
    ):
        try:
            loc = page.locator(sel).first
            if loc.count() and loc.is_visible():
                loc.click()
                page.wait_for_timeout(1800)
                return True
        except Exception:
            continue
    print("  ⚠ 未点中左侧菜单「OTA升级模板」（选择器需按实跑校正）。")
    return False


def _open_view_rules_for_template(page, template_name: str) -> bool:
    """在模板列表：顶部搜索框按标题搜 template_name → 「标题」列精确匹配本批行 → 点其「查看规则」。

    选择器按用户截图/DOM 推断，需实跑校正：标题搜索框 placeholder 可能为『标题』或在
    『导入机型/标题』区；列表为 vxe/ant 表格，标题列文本精确等于 template_name。
    """
    # a) 顶部按标题搜索缩小范围（搜不到搜索框则跳过，直接在整表里精确匹配）。
    for sel in (
        "input[placeholder='标题']",
        "input[placeholder*='标题']",
    ):
        try:
            box = page.locator(sel).first
            if box.count() and box.is_visible():
                box.fill("")
                box.fill(template_name)
                # 点搜索按钮（若有）。
                btn = page.get_by_role("button", name="搜索").first
                if btn.count():
                    btn.click()
                page.wait_for_timeout(1500)
                break
        except Exception:
            continue

    # b) 在「标题」列精确匹配本批行，点该行「查看规则」。含翻页兜底。
    page_guard = 0
    while page_guard < 50:
        page_guard += 1
        clicked = page.evaluate(
            """(tname) => {
                const tables = Array.from(document.querySelectorAll('.vxe-table, table'));
                for (const t of tables) {
                    const heads = t.querySelectorAll('thead th, thead td');
                    if (!heads.length) continue;
                    const names = Array.from(heads).map(h => (h.innerText||'').trim());
                    const idxTitle = names.findIndex(x => x.indexOf('标题') >= 0);
                    if (idxTitle < 0) continue;
                    const rows = t.querySelectorAll('tbody tr');
                    for (const r of rows) {
                        const cells = r.querySelectorAll('td, th');
                        if (!cells.length || !cells[idxTitle]) continue;
                        const title = (cells[idxTitle].innerText||'').trim();
                        if (title !== tname) continue;  // 精确匹配
                        // 该行操作列里找『查看规则』按钮/链接并点击。
                        const cand = Array.from(r.querySelectorAll('button, a, span'))
                            .find(e => (e.innerText||'').trim().indexOf('查看规则') >= 0);
                        if (cand) { cand.click(); return true; }
                    }
                }
                return false;
            }""",
            template_name,
        )
        if clicked:
            page.wait_for_timeout(2000)
            return True
        if not _goto_next_page(page):
            break
        page.wait_for_timeout(1200)
    return False


def _goto_next_page(page) -> bool:
    """点 vxe 分页器『下一页』按钮；无下一页或点击失败返回 False。"""
    try:
        nxt = page.locator(".vxe-pager--next-btn").first
        if nxt.count() and nxt.is_enabled():
            cls = nxt.get_attribute("class") or ""
            if "is--disabled" in cls:
                return False
            nxt.click()
            return True
    except Exception:
        pass
    return False


# —— 阶段2：逐条上线 + 冲突处理 ——
# 编排双区容器/搜索框/列序已按 findings §8（实跑校正）回填。仍未实跑复现的几项（行级
# 上线/下线按钮文本是否全角、上线确认框、规则替换冲突窗、真实 mac组 占位文案 —— findings
# §8.4/8.5/8.8）按用户截图/DOM 推断填入，并在各函数注释标注；写操作均经 D-f gate。

def _search_pending_rule(page, modal, rule_id: str) -> bool:
    """在「待选规则」区按规则ID搜出唯一一条（设备类型已在进页面时于已选区设过一次，此处只用 Rule ID）。

    搜不到/异常返回 False（findings §8.1/8.2）。
    """
    try:
        _search_pending_zone(modal, page, rule_id)
        page.wait_for_timeout(1200)
        return True
    except Exception as e:
        print(f"  ✗ 待选规则搜索 {rule_id!r} 失败：{e}")
        return False


def _find_pending_row_by_ruleid(table, rule_id: str):
    """在待选表里按『规则ID』列文本精确锁定目标行，返回 0-based 行索引；找不到/歧义返回 -1。

    findings §9.5 实测：单个 RuleID（如 65369）可能搜出多行，不能用"恰好 1 行"gate。
    改为按规则ID列（列头含『规则ID』，§8.3 索引 0）文本精确等于归一化 rule_id 的行定位。
    若命中多行（同号多行）同样视为歧义返回 -1，由上层保守拒点。
    """
    want = norm_ruleid(rule_id)
    try:
        idx = table.evaluate(
            """(t, want) => {
                const heads = [...t.querySelectorAll('thead th, thead td')]
                    .map(h => (h.innerText || '').trim());
                let col = heads.findIndex(x => x.replace(/\\s/g,'').indexOf('规则ID') >= 0);
                if (col < 0) col = 0;  // §8.3：规则ID 为第 1 列
                const rows = [...t.querySelectorAll('tbody tr')];
                const hits = [];
                rows.forEach((r, i) => {
                    const cells = r.querySelectorAll('td, th');
                    if (!cells.length || !cells[col]) return;
                    const txt = (cells[col].innerText || '').replace(/\\D/g, '');
                    if (txt && txt === want) hits.push(i);
                });
                if (hits.length === 1) return hits[0];
                return -1;  // 0 行或多行同号 -> 歧义
            }""",
            want,
        )
        return int(idx) if idx is not None else -1
    except Exception:
        return -1


def _click_online(page, modal, rule_id: str) -> bool:
    """D-f gate：在待选表里按规则ID列精确锁定目标行后点该行『上线』。失败只打印不点。

    待选表 = 弹窗内独有 vxe-editable 的 .vxe-table（findings §8.1）。
    ★ findings §9.5（HEADED 实跑校正）：
      · 单个 RuleID 可能搜出多行（实测 65369 搜出 2 行），旧代码"搜出恰好 1 行才点"的 gate 会误拒。
        改为按『规则ID』列文本精确等于目标号的行定位（唯一命中才点，多行同号/0 行则保守拒点）。
      · 行内『上线』实测文本为 `上 线`（**半角空格**，非 danger 红色），与 `编 辑` class 完全相同，
        不能靠 class 区分；用 `_click_button_in` 去空白子串匹配 `上线`（兼容半角/全角空格），
        不会误命中 `编辑`。
    """
    try:
        table = _pending_table(modal)
        row_idx = _find_pending_row_by_ruleid(table, rule_id)
        if row_idx < 0:
            rows = table.locator("tbody tr")
            n = rows.count()
            print(f"  ✗ [gate] 待选规则 {rule_id!r} 未按规则ID列精确锁定到唯一目标行"
                  f"（搜出 {n} 行），拒绝点『上线』。")
            return False
        target = table.locator("tbody tr").nth(row_idx)
        return _click_button_in(target, page, "上线", f"上线({rule_id})")
    except Exception as e:
        print(f"  ✗ 点『上线』失败 {rule_id!r}：{e}")
        return False


def _confirm_online(page) -> bool:
    """D6/D-f gate：『确认上线该规则?』小确认框点『确定』。点前断言确认框在场。

    ⚠ findings §8.5：该小确认框**未实跑复现**（探测未点上线），按 antd 二次确认两种常见形态
      （.ant-modal-confirm 弹窗式 / .ant-popconfirm 气泡式）推断定位，文案判定沿用。
    ⚠ 务必在**该确认框作用域**内取『确定』，绝不误点编排弹窗底部的 `确 定`（那是提交整个编排的写操作，
      findings §8.7）。先用文案可见做 gate，再优先取 confirm/popconfirm 容器作用域内的确定。
    """
    try:
        if page.get_by_text("确认上线该规则", exact=False).count() == 0:
            print("  ✗ [gate] 未见『确认上线该规则?』确认框，拒绝点『确定』。")
            return False
        # 优先在二次确认容器作用域内取『确定』（避免误点编排弹窗底部提交按钮）。
        for sel in (".ant-modal-confirm", ".ant-popover", ".ant-popconfirm"):
            scope = page.locator(sel).last
            try:
                if scope.count() and scope.is_visible():
                    return _click_button_in(scope, page, "确定", "确认上线-确定")
            except Exception:
                continue
        print("  ✗ [gate] 未定位到上线确认框的独立作用域（confirm/popconfirm），"
              "为避免误点编排提交按钮，拒绝点『确定』。")
        return False
    except Exception as e:
        print(f"  ✗ 确认上线失败：{e}")
        return False


def _conflict_modal_present(page) -> bool:
    """是否出现『规则替换』冲突弹窗。"""
    try:
        return page.locator(".ant-modal", has_text="规则替换").count() > 0
    except Exception:
        return False


def _read_mac_group(page, modal, conflict_ruleid: str) -> str | None:
    """去「已选规则」区按被冲突号搜出并读 mac组 列文本。读不到返回 None（D-e 保守中止）。

    findings §8.1/8.2/8.3：已选表=弹窗内第 1 张 .vxe-table；已选区搜索用弹窗内第 1 个 #ruleId；
    mac组 列按列头文本定位（索引 5，不硬编码）。
    ⚠ 真实 mac组 单元格占位文案（空/非空样例）在探测环境无数据未实跑复现（findings §8.3/8.6），
      空判定沿用 mac_group_is_bound/MAC_GROUP_EMPTY_TOKENS；读不到一律 None → 上层保守中止。
    """
    try:
        _search_selected_zone(modal, page, conflict_ruleid)
        page.wait_for_timeout(1200)
        return _read_mac_group_from_table(_selected_table(modal))
    except Exception as e:
        print(f"  ✗ 读取被冲突号 {conflict_ruleid!r} 的 mac组 失败：{e}")
        return None


def _click_replace_confirm(page) -> bool:
    """D7/D-f gate：冲突『规则替换』弹窗点『确定』替换。点前断言弹窗仍在场。"""
    if not _conflict_modal_present(page):
        print("  ✗ [gate] 规则替换弹窗不在场，拒绝点『确定』替换。")
        return False
    modal = page.locator(".ant-modal", has_text="规则替换").first
    return _click_button_in(modal, page, "确定", "规则替换-确定")


def _click_replace_cancel(page) -> bool:
    """冲突『规则替换』弹窗点『取消』中止该条。"""
    if not _conflict_modal_present(page):
        return False
    modal = page.locator(".ant-modal", has_text="规则替换").first
    return _click_button_in(modal, page, "取消", "规则替换-取消")


def run_online(page, rule_ids: list[str], device_code: str = "") -> list[dict]:
    """阶段2：逐条上线 + 冲突处理（D6/D7/D8）。返回每个 Rule ID 的处理 result 列表。

    D8：先打印待上线 Rule ID 清单。对每个 rule_id：待选区搜出→点上线→确认框点确定→轮询冲突弹窗：
      · 无冲突 → 记 online 候选；
      · 有冲突 → 先不点弹窗按钮，解析被冲突号 XXXXX → 去已选区读 mac组：
          非空 → 点『确定』替换 → 记 online；
          空/不确定(D-e) → 点『取消』中止该条、记冲突中止。
    所有写按钮点击经 D-f gate，条件不满足只打印不点。
    """
    print("\n======== 阶段2：待上线 Rule ID 清单（D8）========")
    for rid in rule_ids:
        print(f"  · {rid}")

    # 打开「规则上下线编排」双区弹窗（findings §8.0 真实入口）。
    modal = open_orchestration_modal(page)

    # —— 进页面后【只设一次】设备类型：设在「已选规则」区（用户确认）。——
    # 交互：在已选区『设备类型』框输入设备类型代码 → 自动匹配下拉 → 选中。设失败即终止整个 online。
    # 之后逐条只用 Rule ID 搜（待选区上线搜 / 已选区查状态），不再重设设备类型。
    if device_code:
        if not _set_device_type_once(page, modal, device_code):
            print(f"✗ 已选规则区设置『设备类型』={device_code!r} 失败，按约定终止 online（不逐条上线）。")
            return [{"ruleid": rid, "status": STATUS_FAILED,
                     "note": f"设备类型设置失败，终止：{device_code!r}"} for rid in rule_ids]

    results: list[dict] = []
    for rid in rule_ids:
        print(f"\n---- 上线 Rule ID={rid!r} ----")
        res = {"ruleid": rid, "status": STATUS_FAILED}

        if not _search_pending_rule(page, modal, rid):
            res["note"] = "待选规则搜索失败"
            results.append(res)
            continue
        if not _click_online(page, modal, rid):
            res["note"] = "上线按钮 gate 未满足"
            results.append(res)
            continue
        if not _confirm_online(page):
            res["note"] = "上线确认框 gate 未满足"
            results.append(res)
            continue

        page.wait_for_timeout(1500)
        if _conflict_modal_present(page):
            # D7：先不点弹窗按钮，解析被冲突号。
            try:
                body = page.locator(".ant-modal", has_text="规则替换").first.locator(".ant-modal-body").inner_text()
            except Exception:
                body = ""
            conflict_rid = parse_conflict_ruleid(body)
            print(f"  冲突：规则替换弹窗，被冲突号解析为 {conflict_rid!r}")
            _shot(page, f"active_conflict_{rid}")
            if not conflict_rid:
                print("  ⚠ 未能解析被冲突号，保守中止该条（点取消）。")
                _click_replace_cancel(page)
                res["status"] = STATUS_CONFLICT_ABORT
                res["conflict_ruleid"] = ""
                res["mac_group_bound"] = None
                results.append(res)
                continue

            mac_text = _read_mac_group(page, modal, conflict_rid)
            if mac_text is None:
                print(f"  ⚠ mac组 读取不确定，保守中止该条（点取消）。被冲突号={conflict_rid!r}")
                _click_replace_cancel(page)
                res["status"] = STATUS_CONFLICT_ABORT
                res["conflict_ruleid"] = conflict_rid
                res["mac_group_bound"] = None
                results.append(res)
                continue

            if mac_group_is_bound(mac_text):
                print(f"  mac组 非空（{mac_text!r}）→ 回弹窗点『确定』替换上线。")
                if _click_replace_confirm(page):
                    page.wait_for_timeout(1500)
                    res["status"] = STATUS_ONLINE
                else:
                    res["note"] = "替换确定 gate 未满足"
                res["conflict_ruleid"] = conflict_rid
                res["mac_group_bound"] = True
                results.append(res)
            else:
                print(f"  mac组 为空（{mac_text!r}）→ 点『取消』中止该条并记录。")
                _click_replace_cancel(page)
                res["status"] = STATUS_CONFLICT_ABORT
                res["conflict_ruleid"] = conflict_rid
                res["mac_group_bound"] = False
                results.append(res)
        else:
            print("  无冲突弹窗，记为 online 候选（阶段3 回读确认）。")
            _shot(page, f"active_online_{rid}")
            res["status"] = STATUS_ONLINE
            results.append(res)
    return results


# —— 阶段3：状态汇总 ——

def read_final_status(page, rule_ids: list[str], prior_results: list[dict]) -> list[dict]:
    """D9/D10：对每个 rule_id 在『已选规则』按规则ID搜，操作列为『下线』= online，否则未上线。

    与阶段2 的冲突中止结果合并（冲突中止/失败保留阶段2 判定，不覆盖）。
    """
    # D9 回读在「规则上下线编排」弹窗的「已选规则」区进行（findings §8）。阶段2 结束后弹窗
    # 可能仍开着；不在场则重新打开（只读搜索，不点写按钮）。
    modal = _orchestration_modal(page)
    if modal is None:
        try:
            modal = open_orchestration_modal(page)
        except Exception as e:
            print(f"  ⚠ 回读阶段无法打开编排弹窗：{e}")
            modal = None

    prior = {r["ruleid"]: r for r in prior_results}
    merged: list[dict] = []
    for rid in rule_ids:
        p = prior.get(rid, {"ruleid": rid, "status": STATUS_FAILED})
        # 冲突中止 / 已明确失败的保留阶段2 判定。
        if p.get("status") == STATUS_CONFLICT_ABORT:
            merged.append(p)
            continue
        online = _is_online(page, modal, rid) if modal is not None else False
        if online:
            merged.append({**p, "ruleid": rid, "status": STATUS_ONLINE})
        else:
            note = p.get("note", "回读未在『已选规则』找到或操作列非『下线』")
            merged.append({"ruleid": rid, "status": STATUS_FAILED, "note": note})
    _shot(page, "active_final_list")
    return merged


def _is_online(page, modal, rule_id: str) -> bool:
    """D9：在「已选规则」区按规则ID能搜到该行即 online（findings §8/§8.10 第 9 点）。

    已选区=弹窗内第 1 张 .vxe-table（findings §8.1）；按规则ID精确搜出后，若该表 tbody 有行则
    判定 online。保守：搜不到/异常返回 False。
    （上下线状态在编排双区语义为「能否在已选区搜到」，与行内『下线』按钮文本是否全角空格无关，
      故不依赖 §8.4 未复现的行级按钮文本。）
    """
    try:
        _search_selected_zone(modal, page, rule_id)
        page.wait_for_timeout(1200)
        has_row = _selected_table(modal).evaluate(
            """(t) => {
                const row = t.querySelector('tbody tr');
                if (!row) return false;
                // vxe 空表可能渲染占位行；排除无单元格/标注 No Data 的情形。
                const txt = (row.innerText || '').replace(/\\s|\\u3000/g, '');
                if (!txt) return false;
                if (txt.toLowerCase().indexOf('nodata') >= 0 || txt.indexOf('暂无数据') >= 0) return false;
                return row.querySelectorAll('td, th').length > 0;
            }"""
        )
        return bool(has_row)
    except Exception as e:
        print(f"  ✗ online 回读失败 {rule_id!r}：{e}")
        return False


# ======================================================================
# --probe-online：只读 dump 编排视图真实选择器（校正 findings §2.2/3/5）
# ======================================================================

def probe_online(page) -> None:
    """只读 dump「规则上下线编排」双区弹窗结构，用于校正 findings §8。

    入口为升级事件页行内『规则上下线』→「规则上下线编排」弹窗（findings §8.0 真实入口，
    不再用旧 /otamanage/uprule 单表页）。全程只读：仅打开弹窗 dump 结构，绝不点
    上线/下线/编辑/确定/替换/删除 任何写按钮。
    """
    print(f"只读探测「规则上下线编排」编排视图入口：{config.CMS_UPEVENT_URL}")
    modal = open_orchestration_modal(page)  # 仅打开弹窗（只读），不点任何写按钮
    dump = modal.evaluate(
        """(m) => {
            const out = {title: '', tables: [], buttons: [], hasPendingEditable: false,
                         ruleIdCount: 0, deviceExtCount: 0, selectCount: 0, formCount: 0};
            const title = m.querySelector('.ant-modal-title');
            out.title = title ? (title.innerText||'').trim() : '';
            out.ruleIdCount = m.querySelectorAll('input#ruleId').length;
            out.deviceExtCount = m.querySelectorAll('input#deviceExt').length;
            out.selectCount = m.querySelectorAll('.ant-select').length;
            out.formCount = m.querySelectorAll('form.ant-form-inline').length;
            out.hasPendingEditable = !!m.querySelector('.vxe-table.vxe-editable');
            m.querySelectorAll('.vxe-table').forEach(t => {
                const names = Array.from(t.querySelectorAll('thead th, thead td'))
                    .map(x => (x.innerText||'').trim()).filter(Boolean);
                const editable = t.classList.contains('vxe-editable');
                const rowCount = t.querySelectorAll('tbody tr').length;
                out.tables.push({editable, rowCount, headers: names});
            });
            const seen = new Set();
            m.querySelectorAll('button').forEach(b => {
                const t = (b.innerText||'').trim();
                if (t && !seen.has(t)) { seen.add(t); out.buttons.push(t); }
            });
            return out;
        }"""
    )
    print("  弹窗标题：", dump["title"])
    print("  #ruleId 数量：", dump["ruleIdCount"], " #deviceExt 数量：", dump["deviceExtCount"],
          " .ant-select 数量：", dump["selectCount"], " form.ant-form-inline 数量：", dump["formCount"])
    print("  待选表(vxe-editable)在场：", dump["hasPendingEditable"])
    for i, t in enumerate(dump["tables"]):
        print(f"  表[{i}] editable={t['editable']} rows={t['rowCount']} 表头={t['headers']}")
    print("  可见按钮：", dump["buttons"])
    _shot(page, "active_probe_orchestration")
    print("  ⚠ 只读探测完成，未点任何上线/下线/确定/替换/删除按钮。"
          "请用以上实测结果校正 findings §8 并回填 active_ota 阶段2/3 选择器。")


# ======================================================================
# CLI
# ======================================================================

_HELP_TEXT = """\
active_ota —— CMS OTA 规则。拆成两个独立子命令（一条龙后续再 consolidate）。

用法：
  ./run_active_ota.sh collect [EXCEL]      # 导入模板建规则 + 收集本批 Rule ID
  ./run_active_ota.sh online  [EXCEL]      # 从 Excel 读 Rule ID → 逐条上线 + 冲突 + 汇总
  ./run_active_ota.sh --probe-online       # 只读 dump 编排视图选择器
  ./run_active_ota.sh --selftest           # 纯逻辑自测（不连浏览器）
  ./.venv/bin/python -m automation.active_ota <同上>

子命令 / 模式：
  collect [EXCEL]  连 CMS：导入模板 → 建规则 → 去「OTA升级模板」按标题精确匹配本批行 →
                   「查看规则」读『规则ID』列，升序打印本批 Rule ID（数量应 = Excel 去表头行数）。【写线上数据：导入】
  online  [EXCEL]  连 CMS：从 Excel 的【Rule ID】列读本批 Rule ID → 在「升级事件 → 规则上下线」
                   编排弹窗逐条上线 + 处理「规则替换」冲突 → 回读状态汇总。【写线上数据：上线/替换】
  --probe-online   连 CMS，只读 dump「规则上下线编排」视图的选择器，不做任何写操作。【只读/安全】
  --selftest       纯逻辑自测（名称/备注拼接、冲突号解析、mac组判定、汇总），不连浏览器。【离线/安全】
  -h, --help       显示本帮助并退出。

参数：
  EXCEL            可选，Excel 路径；省略则取 OTA_V4_OUTPUTS 下最新的 .xlsx。
                   collect 要求整表单一「设备类型代码」+单一「目标版本」，否则报错中止。
                   online 从表头为【Rule ID】（或 ruleid / RuleID）的列读本批 Rule ID。

决策要点（D1~D12）：
  名称 = 设备类型代码 + "_" + 目标版本末5位（例 MT9655AMAndrGTV_Q1015）。
  备注 = 今天日期YYYYMMDD + "_USA_" + 导入行数（例 20261102_USA_9，行数去表头）。
  冲突「规则替换」：解析被冲突 Rule ID → 去「已选规则」查其 mac组，非空才点「确定」替换上线，
         mac组为空则点「取消」中止该条并记录。上线全自动，不加人工确认。

退出码：0 全部成功；2 缺子命令/无 Rule ID/存在冲突中止或未上线；1 运行出错。

环境：
  CMS_USERNAME / CMS_PASSWORD   在 config/.env 配置（run_active_ota.sh 会检查）。
  CMS_HEADED=true（默认）         开可见浏览器窗口，便于盯着写流程。

⚠ 安全提示：collect 会真实导入规则；online 会真实上线/替换线上规则，均不可逆。
  首次使用建议顺序：先 --selftest，再 --probe-online，再用仅含 1 条的测试 Excel 小批量试跑，确认无误后再整批。
  所有写按钮（导入确定/上线确定/替换确定）均经二次确认 gate，条件不满足只打印不点。"""


_SUBCOMMANDS = ("collect", "online")


def _parse_cli(argv: list[str]) -> tuple[str, str | None]:
    """解析子命令 + 可选 Excel 路径。

    返回 (mode, excel)。mode ∈ {"collect", "online", "probe"}。
      collect [EXCEL]   导入模板建规则 + 收集本批 Rule ID（升序打印）。
      online  [EXCEL]   从 Excel 的『Rule ID』列读本批 Rule ID → 逐条上线 + 冲突 + 汇总。
      --probe-online    只读 dump 编排视图选择器。
    不给 EXCEL 则取 OTA_V4_OUTPUTS 下最新的 .xlsx。
    """
    mode: str | None = None
    excel: str | None = None
    for arg in argv:
        if arg in ("-h", "--help"):
            print(_HELP_TEXT)
            raise SystemExit(0)
        if arg == "--probe-online":
            mode = "probe"
        elif arg == "--selftest":
            continue  # 在 __main__ 处理
        elif arg in _SUBCOMMANDS:
            if mode is not None:
                print(f"✗ 只能指定一个子命令，已有 {mode!r} 又遇到 {arg!r}。")
                raise SystemExit(2)
            mode = arg
        elif arg.startswith("-"):
            print(f"✗ 未知参数：{arg!r}。用法见 -h。")
            raise SystemExit(2)
        elif excel is None:
            excel = arg
        else:
            print(f"✗ 多余参数：{arg!r}。只接受一个 Excel 路径。")
            raise SystemExit(2)

    if mode is None:
        print("✗ 缺少子命令。用法：collect [EXCEL] | online [EXCEL] | --probe-online。详见 -h。")
        raise SystemExit(2)
    return mode, excel


def main() -> int:
    mode, excel = _parse_cli(sys.argv[1:])

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not config.HEADED)
        context = browser.new_context()
        page = context.new_page()
        try:
            login(page)

            if mode == "probe":
                probe_online(page)
                return 0

            if not excel:
                excel = latest_updated_excel(OTA_V4_OUTPUTS)

            # —— 子命令 collect：导入建规则 + 收集本批 Rule ID（升序打印），不上线 ——
            if mode == "collect":
                print(f"[collect] 导入模板 Excel：{excel}")
                rule_ids, name, remark = run_import(page, excel)
                if not rule_ids:
                    print("✗ [collect] 未回收到任何 Rule ID。")
                    return 2
                print(f"\n[collect] 模板名称：{name}")
                print(f"[collect] 备注：{remark}")
                print(f"[collect] 本批 Rule ID（{len(rule_ids)} 个，升序）：")
                for rid in rule_ids:
                    print(f"  {rid}")
                if config.HEADED:
                    print("\n浏览器将在 20 秒后关闭（可自行查看页面）...")
                    time.sleep(20)
                return 0

            # —— 子命令 online：从 Excel『Rule ID』列读本批 Rule ID → 逐条上线 + 冲突 + 汇总 ——
            if mode == "online":
                print(f"[online] 读取 Excel：{excel}")
                info = read_import_excel(excel)
                # read_import_excel 已按 RULEID_HEADERS（含『Rule ID』）读该列并 norm_ruleid。
                rule_ids = sorted(
                    {r for r in info.get("rule_ids", []) if r},
                    key=lambda x: int(x) if x.isdigit() else 0,
                )
                if not rule_ids:
                    print("✗ [online] Excel 的『Rule ID』列没有可用的 Rule ID（列缺失或为空），终止。")
                    return 2
                print(f"[online] 从 Excel 读到 {len(rule_ids)} 个 Rule ID（升序）：{rule_ids}")

                # 设备类型代码：上线搜索的第二个条件（设备类型代码 + Rule ID）。从 Excel『设备类型代码』列读。
                device_codes = {d for d in info.get("device_codes", set()) if str(d or "").strip()}
                if len(device_codes) == 1:
                    device_code = next(iter(device_codes))
                elif len(device_codes) == 0:
                    device_code = ""
                    print("⚠ [online] Excel『设备类型代码』列为空，上线搜索将只用 Rule ID（精度下降）。")
                else:
                    device_code = ""
                    print(f"⚠ [online] Excel 含多个设备类型代码 {sorted(device_codes)!r}，"
                          "无法确定唯一值，上线搜索将只用 Rule ID（精度下降）。")
                if device_code:
                    print(f"[online] 设备类型代码（搜索第二条件）：{device_code!r}")

                # 阶段2：逐条上线 + 冲突处理（按 设备类型代码 + Rule ID 搜索）。
                results = run_online(page, rule_ids, device_code)
                # 阶段3：状态汇总。
                final = read_final_status(page, rule_ids, results)
                print("\n" + summarize(final))

                if config.HEADED:
                    print("\n浏览器将在 20 秒后关闭（可自行查看页面）...")
                    time.sleep(20)
                return summary_exit_code(final)

            print(f"✗ 未知模式：{mode!r}")
            return 2
        except Exception as e:
            print(f"✗ 运行出错：{e}")
            _shot(page, "active_error")
            return 1
        finally:
            context.close()
            browser.close()


# —— 不依赖浏览器的纯函数自测（python -m automation.active_ota --selftest）——
def _selftest() -> int:
    ok = True

    def check(name, cond):
        nonlocal ok
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        if not cond:
            ok = False

    # —— D1：规则名称拼接 ——
    check("D1 MT9655AMAndrGTV + 0100UQ1015 -> MT9655AMAndrGTV_Q1015",
          build_rule_name("MT9655AMAndrGTV", "0100UQ1015") == "MT9655AMAndrGTV_Q1015")
    check("D1 目标版本末5位取 Q1015",
          build_rule_name("X", "0100UQ1015").endswith("_Q1015"))
    check("D1 两端去空白",
          build_rule_name("  DEV  ", "  AAAAABBBBB  ") == "DEV_BBBBB")
    check("D1 目标版本不足5位取全部",
          build_rule_name("DEV", "AB") == "DEV_AB")

    # —— D2：备注拼接 ——
    check("D2 20261102_USA_9", build_remark(9, "20261102") == "20261102_USA_9")
    check("D2 数量取整", build_remark(0, "20261102") == "20261102_USA_0")
    check("D2 默认今日日期为8位YYYYMMDD",
          re.fullmatch(r"\d{8}_USA_5", build_remark(5)) is not None)

    # —— D3：单一值校验 ——
    dc, tv = validate_single_target({"MT9655AMAndrGTV"}, {"0100UQ1015"})
    check("D3 单一值通过返回 (dc, tv)", dc == "MT9655AMAndrGTV" and tv == "0100UQ1015")
    check("D3 单一值忽略空串干扰",
          validate_single_target({"DEV", ""}, {"V", "  "}) == ("DEV", "V"))
    raised_dc = False
    try:
        validate_single_target({"A", "B"}, {"V"})
    except ValueError:
        raised_dc = True
    check("D3 多设备类型代码报错中止", raised_dc)
    raised_tv = False
    try:
        validate_single_target({"A"}, {"V1", "V2"})
    except ValueError:
        raised_tv = True
    check("D3 多目标版本报错中止", raised_tv)
    raised_empty = False
    try:
        validate_single_target(set(), {"V"})
    except ValueError:
        raised_empty = True
    check("D3 空设备类型代码报错（大小!=1）", raised_empty)

    # —— D7：冲突号解析 ——
    check("D7 解析『和 65246 冲突』-> 65246",
          parse_conflict_ruleid("当前规则和 65246 冲突,是否替换?") == "65246")
    check("D7 解析含千分位『和 65,246 冲突』-> 65246",
          parse_conflict_ruleid("当前规则和 65,246 冲突，是否替换?") == "65246")
    check("D7 解析全角逗号『和 65，246 冲突』-> 65246",
          parse_conflict_ruleid("当前规则和 65，246 冲突?") == "65246")
    check("D7 无匹配返回空串", parse_conflict_ruleid("没有冲突信息") == "")
    check("D7 空文本返回空串", parse_conflict_ruleid("") == "")

    # —— D7：mac组 空/非空判定 ——
    check("D7 mac组『测试MT9』非空 -> True", mac_group_is_bound("测试MT9") is True)
    check("D7 mac组 空串 -> False", mac_group_is_bound("") is False)
    check("D7 mac组『无』-> False", mac_group_is_bound("无") is False)
    check("D7 mac组『null』-> False", mac_group_is_bound("null") is False)
    check("D7 mac组『-』-> False", mac_group_is_bound(" - ") is False)
    check("D7 mac组 None -> False", mac_group_is_bound(None) is False)

    # —— D10：汇总格式（三态）——
    results = [
        {"ruleid": "65344", "status": STATUS_ONLINE},
        {"ruleid": "65343", "status": STATUS_CONFLICT_ABORT, "conflict_ruleid": "65246",
         "mac_group_bound": False},
        {"ruleid": "65342", "status": STATUS_FAILED, "note": "上线确认框 gate 未满足"},
    ]
    report = summarize(results)
    check("D10 汇总含每个 Rule ID",
          all(f"Rule ID={r['ruleid']!r}" in report for r in results))
    check("D10 汇总含 online ✅", "online ✅" in report)
    check("D10 汇总含 冲突中止 ⛔", "冲突中止 ⛔" in report)
    check("D10 汇总含被冲突号 65246", "被冲突号='65246'" in report)
    check("D10 汇总含 mac组为空", "mac组为空" in report)
    check("D10 汇总含 失败 ❌", "失败 ❌" in report)
    check("D10 汇总三态计数",
          "总数：3" in report and "online：1" in report
          and "冲突中止：1" in report and "失败：1" in report)
    check("D10 mac组判定不确定描述",
          "mac组判定不确定" in summarize(
              [{"ruleid": "1", "status": STATUS_CONFLICT_ABORT,
                "conflict_ruleid": "2", "mac_group_bound": None}]))

    # —— 退出码 ——
    check("退出码 全 online -> 0",
          summary_exit_code([{"ruleid": "1", "status": STATUS_ONLINE}]) == 0)
    check("退出码 有冲突中止 -> 非0",
          summary_exit_code([{"ruleid": "1", "status": STATUS_CONFLICT_ABORT}]) != 0)
    check("退出码 空结果 -> 非0", summary_exit_code([]) != 0)

    # —— norm_ruleid 复用对齐（与 dass_fill 口径一致）——
    check("norm_ruleid 去逗号对齐", norm_ruleid("65,246") == norm_ruleid("65246"))

    # —— CLI 子命令解析（collect / online / probe / 缺子命令 / 多子命令 / 未知参数）——
    check("CLI collect 带 Excel", _parse_cli(["collect", "/x.xlsx"]) == ("collect", "/x.xlsx"))
    check("CLI online 带 Excel", _parse_cli(["online", "/x.xlsx"]) == ("online", "/x.xlsx"))
    check("CLI collect 不带 Excel", _parse_cli(["collect"]) == ("collect", None))
    check("CLI online 顺序无关（Excel 在前）", _parse_cli(["/x.xlsx", "online"]) == ("online", "/x.xlsx"))
    check("CLI --probe-online", _parse_cli(["--probe-online"]) == ("probe", None))

    def _expect_exit(argv):
        try:
            _parse_cli(argv)
            return False
        except SystemExit as e:
            return e.code == 2

    check("CLI 缺子命令报错退出2", _expect_exit([]))
    check("CLI 缺子命令（仅 Excel）报错退出2", _expect_exit(["/x.xlsx"]))
    check("CLI 多子命令报错退出2", _expect_exit(["collect", "online"]))
    check("CLI 未知参数报错退出2", _expect_exit(["--bogus"]))

    print("\n自测结果：", "全部通过 ✅" if ok else "存在失败 ❌")
    return 0 if ok else 1


if __name__ == "__main__":
    if "--selftest" in sys.argv[1:]:
        sys.exit(_selftest())
    sys.exit(main())
