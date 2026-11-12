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
from ._log import ok, warn, err, color
from .login import login, _shot
# 复用 DASS 侧常量/工具，绝不修改其实现：
from .dass_fill import norm_ruleid, OTA_V4_OUTPUTS
from .dass_data import latest_updated_excel

# 导入 Excel 里取设备类型代码/目标版本的列名（与 dass_data.EXCEL_TO_DASS 同名，含容错别名）。
DEVICE_CODE_HEADERS = ["设备类型代码"]
TARGET_VERSION_HEADERS = ["目标完整版本", "目标版本"]
RULEID_HEADERS = ["Rule ID", "ruleid", "RuleID"]

# collect 回收 Rule ID 后要插入 Excel 的三列（顺序固定），紧跟在『OTA类型』列之后。
# 逻辑照搬 ota_v4 app/core/legacy_engine.py 的 insert_rule_id_columns（用户授权参考）。
RULE_ID_COLUMNS = ["升级方式", "升级日期", "Rule ID"]
RULE_ID_INSERT_AFTER_HEADER = "OTA类型"
# 升级方式由『OTA类型』列的数字代码映射而来（与 ota_v4 OTA_TYPE_CODE_TO_TEXT 一致）。
OTA_TYPE_CODE_TO_TEXT = {"1": "正常升级", "2": "强制升级", "11": "FTE强制升级"}

# mac组 判定为「空」的占位文本（去空白后命中任一即视为未绑定）。
MAC_GROUP_EMPTY_TOKENS = {"", "无", "null", "none", "-", "—", "空"}


# ======================================================================
# 纯逻辑函数（不连浏览器，--selftest 覆盖全部分支）
# ======================================================================

def build_rule_name(device_code: str, target_version) -> str:
    """D1：规则名称 = 设备类型代码 + "_" + 目标版本最后 5 个字符。

    target_version 可传单个字符串，或多个目标版本（list/tuple/set）：多个时去重、排序后，
    各自取末 5 位用 "_" 连接，例 ["V1.3.9.4","V1.3.9.2"] -> DEV_3.9.2_3.9.4。

    例：build_rule_name("MT9655AMAndrGTV", "0100UQ1015") -> "MT9655AMAndrGTV_Q1015"。
    目标版本不足 5 位时取其全部（[-5:] 的自然语义）。两端去空白。
    """
    dc = str(device_code or "").strip()
    if isinstance(target_version, (list, tuple, set, frozenset)):
        tails = sorted({str(x or "").strip()[-5:] for x in target_version if str(x or "").strip()})
        return f"{dc}_" + "_".join(tails)
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


def validate_import_target(device_codes, target_versions) -> tuple[str, list[str]]:
    """D3：校验导入 Excel 只含单一设备类型代码；目标版本允许多个（至少 1 个）。

    设备类型代码个数 != 1，或没有任何目标版本 → 抛带中文提示的 ValueError（列出实际值，不猜）。
    通过 → 返回 (device_code, 排序去重后的目标版本列表)。
    （收集 Rule ID 按模板名称定位，不依赖目标版本，所以一份模板里可含多个目标版本。）
    """
    dcs = sorted(x for x in device_codes if str(x or "").strip())
    tvs = sorted(x for x in target_versions if str(x or "").strip())
    if len(dcs) != 1:
        raise ValueError(
            f"导入 Excel 的『设备类型代码』应为单一值，实际为 {dcs!r}（共 {len(dcs)} 个）。"
            "请检查模板，不混多个设备类型。"
        )
    if not tvs:
        raise ValueError("导入 Excel 的『目标版本』列没有任何值。")
    return dcs[0], tvs


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
            lines.append(f"  [汇总] Rule ID={rid!r}  状态={color('online ✅', 'green')}（已上线）")
        elif st == STATUS_CONFLICT_ABORT:
            n_conflict += 1
            cf = r.get("conflict_ruleid", "")
            bound = r.get("mac_group_bound", None)
            mac_desc = ("mac组非空" if bound else "mac组为空") if bound is not None else "mac组判定不确定"
            lines.append(
                f"  [汇总] Rule ID={rid!r}  状态={color('冲突中止 ⛔', 'red')}（被冲突号={cf!r}，{mac_desc}）"
            )
        else:
            n_failed += 1
            note = r.get("note", "")
            lines.append(f"  [汇总] Rule ID={rid!r}  状态={color('失败 ❌', 'red')}" + (f"（{note}）" if note else ""))
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


def write_rule_id_columns(file_path: str, rule_ids: list[str], sheet_name: str = "升级规则") -> int:
    """collect 回收 Rule ID 后，把『升级方式/升级日期/Rule ID』三列写回【同一个】Excel 文件。

    逻辑照搬 ota_v4 app/core/legacy_engine.py insert_rule_id_columns（用户授权参考）：
      · 三列紧跟在『OTA类型』列之后插入（位置固定），保留其余列顺序/内容；
      · 升级方式 = 该行『OTA类型』数字代码映射（1→正常升级 / 2→强制升级 / 11→FTE强制升级）；
      · 升级日期 = 今天 YYYYMMDD；
      · Rule ID = 按数据行顺序逐行填 rule_ids（已升序，第 1 行放最小值）；
      · 幂等保护：若已含这些列则报错不重复插入；
      · Rule ID 数量必须 == 数据行数，否则报错（不猜）。
    写回传入的 file_path 本身（覆盖更新，确保是一个文件）。返回写入的数据行数。
    """
    from openpyxl.utils import get_column_letter

    if not file_path or not Path(file_path).exists():
        raise ValueError(f"要回写的 Excel 文件不存在：{file_path!r}")

    wb = load_workbook(file_path)
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.worksheets[0]

    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]

    # 幂等：已含任一目标列则拒绝重复插入。
    if any(h in RULE_ID_COLUMNS for h in headers):
        raise ValueError("该 Excel 已含『升级方式/升级日期/Rule ID』列，无需重复写入（幂等保护）。")

    if RULE_ID_INSERT_AFTER_HEADER not in headers:
        raise ValueError(f"表头未找到『{RULE_ID_INSERT_AFTER_HEADER}』列，无法定位三列插入位置。")
    ota_col = headers.index(RULE_ID_INSERT_AFTER_HEADER) + 1  # 1-based
    insert_at = ota_col + 1  # 紧跟 OTA类型 之后

    # 统计数据行数（末尾非空数据行）。
    data_row_count = 0
    for r in range(2, ws.max_row + 1):
        if any(ws.cell(row=r, column=c).value not in (None, "") for c in range(1, ws.max_column + 1)):
            data_row_count = r - 1

    assigned = [str(x) for x in rule_ids]
    if len(assigned) != data_row_count:
        raise ValueError(
            f"Rule ID 数量({len(assigned)}) 与数据行数({data_row_count}) 不匹配，拒绝回写以免错位。"
        )

    today = datetime.now().strftime("%Y%m%d")

    # 插入三列（openpyxl 只移动数据，不移动列宽；这里不额外处理列宽，保持简单）。
    ws.insert_cols(insert_at, amount=len(RULE_ID_COLUMNS))
    for offset, name in enumerate(RULE_ID_COLUMNS):
        cell = ws.cell(row=1, column=insert_at + offset, value=name)
        cell.number_format = "@"

    # 逐行填值：升级方式（按 OTA类型 代码映射）/ 升级日期 / Rule ID（升序逐行）。
    for r in range(2, data_row_count + 2):
        ota_code = str(ws.cell(row=r, column=ota_col).value or "").strip()
        upgrade_method = OTA_TYPE_CODE_TO_TEXT.get(ota_code, "")
        rid = assigned[r - 2]
        for offset, value in enumerate((upgrade_method, today, rid)):
            c = ws.cell(row=r, column=insert_at + offset, value=value)
            c.number_format = "@"

    try:
        wb.save(file_path)
    except PermissionError:
        raise ValueError(
            "回写失败：Excel 文件可能正被 Excel / Numbers 打开并锁定，请关闭后重试。"
        )
    return data_row_count


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


def _pending_rule_id_box(modal):
    """待选区『规则ID』输入框 = 弹窗内【第 2 个】input#ruleId（DOM 顺序：已选区在前、待选区在后）。

    ★ 实跑校正（2026-11-09 截图）：已选区的搜索表单同样带 `form.ant-form-inline`，
      原来的 `modal.locator("form.ant-form-inline").first` 实际取到了【已选区】表单，
      导致上线时把 Rule ID 填进已选区、待选区始终没过滤（搜出 0 行或碰巧在首页才找到）。
    少于 2 个 #ruleId 时抛 RuntimeError，宁可失败也不填错区。
    """
    boxes = modal.locator("input#ruleId")
    n = boxes.count()
    if n < 2:
        raise RuntimeError(f"弹窗内只找到 {n} 个 #ruleId，无法确认待选区输入框（期望 2 个）。")
    return boxes.nth(1)


def _pending_form(modal):
    """待选规则搜索表单 = 待选区 #ruleId 输入框所在的 form（取不到 form 时取最近含按钮的祖先）。"""
    box = _pending_rule_id_box(modal)
    form = box.locator("xpath=ancestor::form[1]")
    if form.count():
        return form.first
    return box.locator("xpath=ancestor::*[.//button][1]").first


def _search_pending_zone(modal, page, rule_id: str) -> bool:
    """在「待选规则」区按规则ID精确搜索（findings §8.2）。

    待选搜索表单 = form.ant-form-inline（block 2），区内 #ruleId placeholder=精准匹配；
    搜索按钮文本无全角空格。先锚待选表单再取 #ruleId，避免与已选区同名控件混淆。
    （设备类型不在此设置——它已在进页面时于已选区设过一次。）
    """
    box = _pending_rule_id_box(modal)
    want = norm_ruleid(rule_id)
    box.fill(want)
    # 填完核对：输入框里确实是目标号，且已选区（第 1 个 #ruleId）没被动过，才继续。
    if (box.input_value() or "").strip() != want:
        print(color(f"  ✗ 待选区『规则ID』框填入后值不符（期望 {want!r}）。", "red"))
        return False
    print(f"  ✓ 已在待选区『规则ID』框输入 {want!r}")
    return _click_button_in(_pending_form(modal), page, "搜索", "待选规则-搜索")


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
        print(color("  ✗ [gate] 导入弹窗不在场，拒绝点击导入确认按钮。", "red"))
        return False

    # —— 第一步：优先点『下一步』进入第二步；没有『下一步』则视为单步向导。——
    went_next = _click_button_in(modal, page, "下一步", "导入向导第一步(下一步)")
    if went_next:
        page.wait_for_timeout(1800)
        # 点『下一步』后后端可能弹业务错误（目标版本已存在 / Service error）——立即终止。
        _abort_if_ui_error(page, "点『下一步』后")
        modal = _active_modal(page) or modal  # 第二步可能换了 modal 实例
        if modal is None or not modal.is_visible():
            print(color("  ✗ [gate] 点『下一步』后导入弹窗不在场，拒绝继续。", "red"))
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
        print(color("  ⚠ 已点『下一步』但未找到第二步的『保存/完成/确定』按钮——"
                    "请确认导入是否已完成（可能第二步按钮文案不同，需按实跑日志校正）。", "yellow"))
        return True

    print(color("  ✗ 未找到任何导入完成按钮（下一步/保存/完成/确定/导入均未命中）。", "red"))
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
    device_code, target_versions = validate_import_target(
        info["device_codes"], info["target_versions"]
    )
    target_version = target_versions[0] if len(target_versions) == 1 else target_versions
    if len(target_versions) > 1:
        print(color(f"  ⚠ Excel 含 {len(target_versions)} 个目标版本 {target_versions}，"
                    "模板名称会带上全部版本末 5 位。", "yellow"))
    name = build_rule_name(device_code, target_versions)
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
        print(color(f"  ⚠ 回收到 {len(seen_sorted)} 个 Rule ID，期望 {expected}（= Excel 去表头行数）。"
                    f"数量不符，请核对是否定位错模板行/翻页不全：{seen_sorted}", "yellow"))
    else:
        print(color(f"  ✓ 回收本批 Rule ID（{len(seen_sorted)} 个，升序）：{seen_sorted}", "green"))
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
        if not _search_pending_zone(modal, page, rule_id):
            return False
        # 等待选表刷新出目标 Rule ID 那一行（最多约 8 秒；旧结果的行不算）。
        # 等不到就照常返回，由上线 gate 按实际行数报错。
        table = _pending_table(modal)
        for _ in range(40):
            page.wait_for_timeout(200)
            if _find_pending_row_by_ruleid(table, rule_id) >= 0:
                break
        page.wait_for_timeout(300)
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
            print(color(f"  ✗ [gate] 待选规则 {rule_id!r} 未按规则ID列精确锁定到唯一目标行"
                        f"（搜出 {n} 行），拒绝点『上线』。", "red"))
            return False
        target = table.locator("tbody tr").nth(row_idx)
        return _click_button_in(target, page, "上线", f"上线({rule_id})")
    except Exception as e:
        print(color(f"  ✗ 点『上线』失败 {rule_id!r}：{e}", "red"))
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
            print(color("  ✗ [gate] 未见『确认上线该规则?』确认框，拒绝点『确定』。", "red"))
            return False
        # 优先在二次确认容器作用域内取『确定』（避免误点编排弹窗底部提交按钮）。
        for sel in (".ant-modal-confirm", ".ant-popover", ".ant-popconfirm"):
            scope = page.locator(sel).last
            try:
                if scope.count() and scope.is_visible():
                    return _click_button_in(scope, page, "确定", "确认上线-确定")
            except Exception:
                continue
        print(color("  ✗ [gate] 未定位到上线确认框的独立作用域（confirm/popconfirm），"
                    "为避免误点编排提交按钮，拒绝点『确定』。", "red"))
        return False
    except Exception as e:
        print(color(f"  ✗ 确认上线失败：{e}", "red"))
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
        if not _search_selected_zone(modal, page, conflict_ruleid):
            return None
        table = _selected_table(modal)
        # 等已选表刷新出被冲突号那一行（最多约 8 秒）；不把旧结果/No Data 当成 mac组 为空。
        for _ in range(40):
            page.wait_for_timeout(200)
            if _find_pending_row_by_ruleid(table, conflict_ruleid) >= 0:
                break
        else:
            print(color(f"  ✗ 已选区搜 {conflict_ruleid!r} 后未出现该行（No Data 或加载超时），mac组 无法确认。", "red"))
            return None
        return _read_mac_group_from_table(table)
    except Exception as e:
        print(color(f"  ✗ 读取被冲突号 {conflict_ruleid!r} 的 mac组 失败：{e}", "red"))
        return None


def _click_replace_confirm(page) -> bool:
    """D7/D-f gate：冲突『规则替换』弹窗点『确定』替换。点前断言弹窗仍在场。"""
    if not _conflict_modal_present(page):
        print(color("  ✗ [gate] 规则替换弹窗不在场，拒绝点『确定』替换。", "red"))
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
            print(color(f"✗ 已选规则区设置『设备类型』={device_code!r} 失败，按约定终止 online（不逐条上线）。", "red"))
            return [{"ruleid": rid, "status": STATUS_FAILED,
                     "note": f"设备类型设置失败，终止：{device_code!r}"} for rid in rule_ids]

    results: list[dict] = []
    for rid in rule_ids:
        print(f"\n---- 上线 Rule ID={rid!r} ----")
        res = {"ruleid": rid, "status": STATUS_FAILED}

        if _already_online(page, modal, rid):
            print(color(f"  ✓ Rule ID {rid!r} 已在『已选规则』中（已上线）→ 跳过。", "green"))
            res["status"] = STATUS_ONLINE
            res["note"] = "已上线，跳过"
            results.append(res)
            continue

        if not _search_pending_rule(page, modal, rid):
            res["note"] = "待选规则搜索失败"
            results.append(res)
            continue
        if not _click_online(page, modal, rid):
            # 点『上线』没成功时，可能其实已弹出「规则替换」冲突窗（挡住了表格）：有就走冲突替换，
            # 没有才记失败；失败时留一张截图方便定位。
            page.wait_for_timeout(800)
            if not _conflict_modal_present(page):
                _shot(page, f"active_gate_fail_{rid}")
                res["note"] = "上线按钮 gate 未满足"
                results.append(res)
                continue
        elif not _confirm_online(page):
            page.wait_for_timeout(800)
            if not _conflict_modal_present(page):
                _shot(page, f"active_gate_fail_{rid}")
                res["note"] = "上线确认框 gate 未满足"
                results.append(res)
                continue

        page.wait_for_timeout(1500)
        if _online_hint_visible(page):
            print(color(f"  ✓ 系统提示 Rule ID {rid!r} 已上线 → 跳过。", "green"))
            _shot(page, f"active_already_online_{rid}")
            res["status"] = STATUS_ONLINE
            res["note"] = "系统提示已上线，跳过"
            results.append(res)
            continue
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
                print(color("  ⚠ 未能解析被冲突号，保守中止该条（点取消）。", "yellow"))
                _click_replace_cancel(page)
                res["status"] = STATUS_CONFLICT_ABORT
                res["conflict_ruleid"] = ""
                res["mac_group_bound"] = None
                results.append(res)
                continue

            # 冲突弹窗的遮罩盖住整个编排弹窗，开着的时候点不到下面的『已选规则-搜索』（实跑截图：
            # 被冲突号已填进框但表仍是 No Data）。所以先点『取消』关掉弹窗，再去已选区读 mac组。
            _click_replace_cancel(page)
            page.wait_for_timeout(800)

            # 判定原则：查被冲突 Rule ID 的 mac组；非空 → 替换；为空 → 中止该条。
            mac_text = _read_mac_group(page, modal, conflict_rid)
            if mac_text is None:
                print(color(f"  ⚠ mac组 读取不确定，保守中止该条。被冲突号={conflict_rid!r}", "yellow"))
                res["status"] = STATUS_CONFLICT_ABORT
                res["conflict_ruleid"] = conflict_rid
                res["mac_group_bound"] = None
                results.append(res)
                continue

            if not mac_group_is_bound(mac_text):
                print(color(f"  被冲突号 {conflict_rid} 的 mac组 为空（{mac_text!r}）→ 中止该条并记录。", "red"))
                res["status"] = STATUS_CONFLICT_ABORT
                res["conflict_ruleid"] = conflict_rid
                res["mac_group_bound"] = False
                results.append(res)
                continue

            # mac组 非空 → 重新对这条点一次『上线』，第二次弹出冲突窗时点『确定』替换。
            print(color(f"  被冲突号 {conflict_rid} 的 mac组 非空（{mac_text!r}）→ 重新上线并点『确定』替换。", "green"))
            res["conflict_ruleid"] = conflict_rid
            res["mac_group_bound"] = True
            if not (_search_pending_rule(page, modal, rid)
                    and _click_online(page, modal, rid)
                    and _confirm_online(page)):
                _shot(page, f"active_gate_fail_{rid}")
                res["note"] = "冲突后重新上线 gate 未满足"
                results.append(res)
                continue
            page.wait_for_timeout(1500)
            if not _conflict_modal_present(page):
                print(color("  重新上线后未再出现冲突弹窗，记为 online 候选（阶段3 回读确认）。", "yellow"))
                res["status"] = STATUS_ONLINE
            elif _click_replace_confirm(page):
                page.wait_for_timeout(1500)
                res["status"] = STATUS_ONLINE
            else:
                res["note"] = "替换确定 gate 未满足"
            results.append(res)
        else:
            print(color("  无冲突弹窗，记为 online 候选（阶段3 回读确认）。", "green"))
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
    online_ids = [r["ruleid"] for r in merged if r.get("status") == STATUS_ONLINE]
    if online_ids:
        from .stage_db import record_cms
        record_cms(online_ids, "在线", "online")
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
        print(color(f"  ✗ online 回读失败 {rule_id!r}：{e}", "red"))
        return False


def _already_online(page, modal, rule_id: str) -> bool:
    """上线前预检：已选规则区按 Rule ID 能精确搜到该行 → 该规则已经在线，可跳过。

    精确匹配规则ID列（不是只要有行就算）；最多等约 3 秒，搜不到/异常一律返回 False（照常上线）。
    """
    try:
        if not _search_selected_zone(modal, page, rule_id):
            return False
        table = _selected_table(modal)
        for _ in range(15):
            page.wait_for_timeout(200)
            if _find_pending_row_by_ruleid(table, rule_id) >= 0:
                return True
        return False
    except Exception:
        return False


_ONLINE_HINTS = ("已上线", "已经上线", "已在线", "已处于上线")


def _online_hint_visible(page) -> bool:
    """页面上的提示（toast / 通知 / 对话框）是否表示该规则已经上线。"""
    try:
        for sel in (".ant-message", ".ant-notification", ".ant-modal-confirm", ".ant-popover"):
            loc = page.locator(sel)
            for i in range(min(loc.count(), 5)):
                txt = (loc.nth(i).inner_text() or "").replace(" ", "")
                if any(h in txt for h in _ONLINE_HINTS):
                    return True
    except Exception:
        pass
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

_HELP_COMMON = """\
Excel：省略时用 prep 最近生成的那份；也可写路径（相对路径按当前目录）。
退出码：0 成功；1 运行出错；2 参数问题 / 有未上线或冲突中止；publish 检查点失败 11/12/13/14。
CMS_HEADED=true（默认）开可见浏览器；导入、上线、替换都会改线上数据，不可撤销。"""

_HELP_BY_MODE = {
    "collect": """\
active_ota collect [EXCEL] —— CMS 导入模板建规则，回收本批 Rule ID 写回 Excel（publish 的第 1-2 步；不再单独作为 deployota 子命令）

  1) OTA升级模板 → 导入模板：名称 = 设备类型代码_目标版本末5位（多个目标版本全部带上），备注 = YYYYMMDD_USA_行数
  2) 在模板列表按名称找到本批 → 查看规则 → 读规则ID（数量应 = Excel 数据行数）
  3) 把『升级方式 / 升级日期 / Rule ID』三列写回同一个 Excel
要求：整表只有一个设备类型代码；目标版本可以有多个。之后用 ./deployota online 上线。
""",
    "online": """\
deployota online [EXCEL | Rule ID...] [--device CODE] —— CMS 上线

  ./deployota online                      用 prep 最近生成的 Excel 里的 Rule ID 列
  ./deployota online 路径.xlsx            指定 Excel
  ./deployota online 50000-50005          直接给 Rule ID（单个 / 区间 / 逗号，可混用）
  ./deployota online 50000 --device X     找不到设备类型代码时手工指定（一般不需要）

逐条：已在『已选规则』的跳过 → 『待选规则』按 Rule ID 搜索 → 上 线 → 确认。
出现『规则替换』冲突：查被冲突规则的 mac 组，非空才替换，为空或读不到则中止该条。
结束后回读状态，打印上线汇总和部署信息汇总表，并记录 CMS 状态『在线』。
设备类型代码按 Rule ID 自动找：部署库 → prep Excel → CMS 在线规则页。
""",
    "publish": """\
deployota publish [EXCEL] [--no-dass] —— 一步完成：导入 → Rule ID → 上线 → DASS 场测

  检查点 ① 导入成功  ② 得到 Rule ID（数量 = Excel 行数，回写并读回一致）
         ③ 全部上线  ④ DASS init 填写完成（之后仍需你输入 approve 才提交；记录阶段『场测』）
  任一检查点失败：红框提醒原因和下一步，后面不再执行。--no-dass 只做前 3 个。
  Excel 已含 Rule ID（导入过）时拒绝重复导入，只需上线请用 ./deployota online。
""",
}


def _help_text(mode: str | None) -> str:
    if mode in _HELP_BY_MODE:
        return _HELP_BY_MODE[mode] + "\n" + _HELP_COMMON
    return "\n".join(_HELP_BY_MODE.values()) + "\n" + _HELP_COMMON





# —— publish：导入(upload) + 上线(online) 合并，三个检查点逐一把关 ——
CP_NAMES = {1: "导入上传成功", 2: "得到 Rule ID", 3: "上线成功", 4: "DASS init 填报完成"}
# 退出码：检查点 n 失败 → 10+n；前置检查不通过 → 2。
EXIT_CP_FAIL = {1: 11, 2: 12, 3: 13, 4: 14}


def _print_deploy_summary(excel: str, final: list[dict] | None = None) -> None:
    """上线完成后打印部署信息汇总（与 dass 更新后的汇总同一张表，目标版本为『目标完整版本』）。

    final 给定时只在至少有一条 online 时打印；读不到 Excel 或没有 Rule ID 列时静默跳过。
    """
    if final is not None and not any(r.get("status") == STATUS_ONLINE for r in final):
        return
    try:
        from .summary import print_summary
        print("")
        print_summary(excel)
    except Exception as e:
        print(color(f"（部署汇总显示跳过：{e}）", "dim"))


def format_checkpoint_report(cps: list[dict]) -> str:
    """把三个检查点的结果排成可检索的报告（纯逻辑）。

    cps 每项：{"n": 1|2|3, "status": "pass"|"fail"|"skip", "detail": str}。
    全部 pass → 末行绿色『全部检查点通过』；有 fail → 末行红色『发布未完成』并点出第一个失败的检查点。
    """
    lines = ["======== publish 检查点报告 ========"]
    first_fail = None
    for cp in cps:
        n, st, detail = cp["n"], cp["status"], cp.get("detail", "")
        title = f"检查点{n} {CP_NAMES[n]}"
        if st == "pass":
            lines.append(color(f"  ✓ {title}", "green") + (f"：{detail}" if detail else ""))
        elif st == "fail":
            lines.append(color(f"  ✗ {title}：失败", "red") + (f" —— {detail}" if detail else ""))
            if first_fail is None:
                first_fail = n
        else:
            lines.append(color(f"  - {title}：未执行", "yellow") + (f"（{detail}）" if detail else ""))
    if first_fail is None and all(c["status"] == "pass" for c in cps):
        done = "已导入、拿到 Rule ID、全部上线成功" + ("、DASS init 已填报" if len(cps) >= 4 else "")
        lines.append(color(f"  ★ 全部检查点通过：{done}。", "green"))
    else:
        lines.append(color(f"  ✗ 发布未完成：卡在检查点{first_fail or '?'}（{CP_NAMES.get(first_fail, '')}）。"
                           "请按上面的提示处理，处理前不要直接重跑。", "red"))
    return "\n".join(lines)


def _cp_alert(n: int, reason: str, hints: list[str]) -> None:
    """检查点失败时的醒目提醒：红色框 + 原因 + 下一步建议。"""
    bar = "!" * 64
    print("\n" + color(bar, "red"))
    print(color(f"!! 检查点{n}【{CP_NAMES[n]}】失败", "red"))
    print(color(f"!! 原因：{reason}", "red"))
    for h in hints:
        print(color(f"!! 建议：{h}", "red"))
    print(color(bar, "red") + "\n")


def _run_dass_init(excel: str) -> int:
    """在子进程里跑 `dass_fill init <excel>`（它自己起浏览器；子进程继承终端，人工核对/批复照常交互）。"""
    import subprocess
    root = Path(__file__).resolve().parent.parent
    return subprocess.run([sys.executable, "-m", "automation.dass_fill", "init", excel], cwd=str(root)).returncode


def run_publish(page, excel: str, with_dass: bool = True) -> int:
    """导入(upload) + 上线(online) + DASS init 合并执行，四个检查点（with_dass=False 则只前三个）：
      1) 导入上传成功   —— 导入向导走完且无 UI 错误，模板在『OTA升级模板』列表里找到；
      2) 得到 Rule ID   —— 回收到的 Rule ID 数 == Excel 数据行数，并成功回写、回读一致；
      3) 上线成功       —— 全部 Rule ID 回读为已上线，无失败/冲突中止。
      4) DASS init 填报完成 —— 上线全部成功后，用同一份 Excel 跑 dass init（填写后仍要你在终端人工核对、
                          输入 approve 才提交）；退出码 0 才算通过，部分行没填成功(3)/出错都算失败。
    任一检查点失败：立刻红色醒目提醒（原因+下一步），后面的步骤不再执行，返回 10+n。
    返回 0 全部通过；2 前置检查不通过（未动 CMS）。
    """
    steps = (1, 2, 3, 4) if with_dass else (1, 2, 3)
    cps = [{"n": n, "status": "skip", "detail": "前一步未通过，未执行"} for n in steps]

    def finish(code: int) -> int:
        print("\n" + format_checkpoint_report(cps))
        if config.HEADED:
            print("\n浏览器将在 20 秒后关闭（可自行查看页面）...")
            time.sleep(20)
        return code

    def fail(n: int, reason: str, hints: list[str], shot: str | None = None) -> int:
        cps[n - 1].update(status="fail", detail=reason)
        _cp_alert(n, reason, hints)
        if shot:
            _shot(page, shot)
        return finish(EXIT_CP_FAIL[n])

    # —— 前置检查（不动 CMS）——
    info = read_import_excel(excel)
    count = info["count"]
    if count == 0:
        print(color("✗ [publish] 导入 Excel 没有数据行，终止（未动 CMS）。", "red"))
        return 2
    if any(info.get("rule_ids", [])):
        print(color("✗ [publish] 这份 Excel 已经含有 Rule ID，说明之前已导入过；再导入会触发 CMS "
                    "『目标版本已经存在』。终止（未动 CMS）。", "red"))
        print(color("  只需要上线请用：./deployota online   需要重新导入请先确认 CMS 上的旧模板已处理。", "red"))
        return 2
    try:
        device_code, target_versions = validate_import_target(info["device_codes"], info["target_versions"])
    except ValueError as e:
        print(color(f"✗ [publish] {e}（未动 CMS）", "red"))
        return 2
    name = build_rule_name(device_code, target_versions)
    print(f"[publish] Excel：{excel}\n[publish] 设备类型代码={device_code!r}  模板名称={name!r}  数据行数={count}")

    # —— 检查点1：导入上传成功 ——
    print(color("\n▶ 检查点1：导入模板（upload）", "bold"))
    try:
        rule_ids, name, remark = run_import(page, excel)
    except Exception as e:
        return fail(1, str(e), [
            f"到 CMS『OTA升级模板』列表核对标题 {name!r} 是否已经创建。",
            "已创建就不要直接重跑（会报『目标版本已经存在』）：先处理掉旧模板，或手工取 Rule ID 填回 Excel 后用 ./deployota online。",
            "未创建：按上面的错误原文修正 Excel 后重跑 ./deployota publish。",
        ], "publish_cp1_fail")
    cps[0].update(status="pass", detail=f"模板 {name!r} 已导入（备注 {remark!r}）")
    print(color(f"✓ 检查点1 通过：模板 {name!r} 已导入。", "green"))

    # —— 检查点2：得到 Rule ID ——
    print(color("\n▶ 检查点2：核对 Rule ID 并回写 Excel", "bold"))
    if not rule_ids:
        return fail(2, "导入后未回收到任何 Rule ID。", [
            f"模板 {name!r} 很可能已在 CMS 创建，请去『OTA升级模板→查看规则』确认。",
            "不要重跑 publish；确认规则存在后，把 Rule ID 填进 Excel 的『Rule ID』列再用 ./deployota online。",
        ], "publish_cp2_fail")
    if len(rule_ids) != count:
        return fail(2, f"回收到 {len(rule_ids)} 个 Rule ID，但 Excel 有 {count} 行数据（{rule_ids}）。", [
            f"到 CMS 核对模板 {name!r} 的规则数量；数量对不上时不能自动上线，以免漏传或误传。",
            "核对无误后，把 Rule ID 填进 Excel 再用 ./deployota online。",
        ], "publish_cp2_fail")
    try:
        n_written = write_rule_id_columns(excel, rule_ids)
        back = sorted((r for r in read_import_excel(excel).get("rule_ids", []) if r),
                      key=lambda x: int(x) if x.isdigit() else 0)
    except Exception as e:
        return fail(2, f"Rule ID 回写 Excel 失败：{e}", [
            f"规则已在 CMS 创建（{rule_ids}），请手工填入 Excel 的『Rule ID』列后用 ./deployota online。",
        ], "publish_cp2_fail")
    if back != sorted(rule_ids, key=lambda x: int(x) if x.isdigit() else 0):
        return fail(2, f"回写后回读 Excel 的 Rule ID 与回收结果不一致：写入 {rule_ids}，读回 {back}。", [
            "请打开 Excel 核对『Rule ID』列后再用 ./deployota online。",
        ], "publish_cp2_fail")
    cps[1].update(status="pass", detail=f"{len(rule_ids)} 个（{rule_ids[0]}~{rule_ids[-1]}），已回写 Excel {n_written} 行并回读一致")
    print(color(f"✓ 检查点2 通过：{len(rule_ids)} 个 Rule ID，已回写 Excel。", "green"))

    # —— 检查点3：上线成功 ——
    print(color("\n▶ 检查点3：逐条上线（online）", "bold"))
    try:
        results = run_online(page, rule_ids, device_code)
        final = read_final_status(page, rule_ids, results)
    except Exception as e:
        return fail(3, f"上线过程出错：{e}", [
            "Rule ID 已在 Excel 里。先到 CMS『规则上下线编排』核对哪些已上线，再用 ./deployota online 重跑（已上线的会自动跳过）。",
        ], "publish_cp3_fail")
    print("\n" + summarize(final))
    bad = [r for r in final if r.get("status") != STATUS_ONLINE]
    if bad:
        detail = "；".join(
            f"{r['ruleid']}={'冲突中止' if r.get('status') == STATUS_CONFLICT_ABORT else '失败'}"
            + (f"（{r['note']}）" if r.get("note") else "")
            for r in bad)
        return fail(3, f"{len(bad)}/{len(final)} 条未上线：{detail}", [
            "未上线的 Rule ID 见上方汇总；冲突中止 = 被冲突规则的 mac 组为空或无法确认，需人工处理。",
            "处理完后用 ./deployota online 重跑（已上线的会自动跳过）。",
        ], "publish_cp3_fail")
    cps[2].update(status="pass", detail=f"{len(final)}/{len(final)} 条全部在线")
    print(color(f"✓ 检查点3 通过：{len(final)} 条 Rule ID 全部上线。", "green"))
    if not with_dass:
        _print_deploy_summary(excel)
        return finish(0)

    # —— 检查点4：DASS init 填报 ——
    print(color("\n▶ 检查点4：DASS init 填报（填写后需你在终端人工核对并输入 approve 才提交）", "bold"))
    try:
        rc = _run_dass_init(excel)
    except Exception as e:
        return fail(4, f"无法启动 DASS 填报：{e}", [
            "CMS 导入和上线都已完成，不要重跑 publish。单独跑：./deployota dass init",
        ])
    redo = "CMS 导入和上线都已完成，不要重跑 publish。排查后单独跑：./deployota dass init"
    if rc == 3:
        return fail(4, "DASS init 有行没填成功（退出码 3）。", [
            "见上方 dass 日志里标 ⚠ 的行；到 DASS 填报页核对，补填或删掉多插的行。", redo])
    if rc != 0:
        return fail(4, f"DASS init 异常退出（退出码 {rc}）。", [
            "见上方 dass 日志里的错误原文与 artifacts/ 里的截图。", redo])
    cps[3].update(status="pass", detail=f"{len(final)} 条已填入 DASS（提交与否以你在终端的 approve 为准）")
    print(color("✓ 检查点4 通过：DASS init 已填写完成。", "green"))
    return finish(0)


_SUBCOMMANDS = ("collect", "online", "publish")


# Rule ID 规格：65370 / 65370-65379 / 65370,65375（可混用）。用来区分『Rule ID 参数』和『Excel 路径』。
_RULE_SPEC_RE = re.compile(r"^\d[\d,，\-~\s]*$")


def split_rule_args(argv: list[str]) -> tuple[list[str], list[str], str]:
    """从参数里拿出 Rule ID 规格和 --device。返回 (剩余参数, Rule ID 规格列表, 设备类型代码)。"""
    rest, specs, device = [], [], ""
    i = 0
    while i < len(argv):
        a = argv[i]
        if a == "--device":
            if i + 1 >= len(argv):
                print("✗ --device 后面要跟设备类型代码。")
                raise SystemExit(2)
            device = argv[i + 1].strip()
            i += 2
            continue
        if a.startswith("--device="):
            device = a.split("=", 1)[1].strip()
        elif _RULE_SPEC_RE.match(a.strip()):
            specs.append(a)
        else:
            rest.append(a)
        i += 1
    return rest, specs, device


def _print_online_summary_by_ids(online_ids: list[str], code_of: dict) -> None:
    """按 Rule ID 上线后的部署信息汇总：明细取部署库/来源 Excel（目标版本为目标完整版本）。"""
    if not online_ids:
        return
    from .active_offline import group_by_device
    from .dass_update import resolve_rules
    from .summary import render_summary
    from . import depdb
    batches = depdb.load_all()
    resolved, _ = resolve_rules(online_ids, batches)
    details = {rec["ruleid"]: rec for rec, _src in resolved}
    today = datetime.now().strftime("%Y%m%d")

    def deploy_date(rid: str) -> str:
        """部署时间取部署库批次（最新的有效批次）；库里没有才用今天。"""
        owners = [b for b in batches if str(b.get("状态", depdb.STATUS_ACTIVE)) == depdb.STATUS_ACTIVE
                  and rid in [x.strip() for x in str(b.get("Rule ID列表", "")).split(",")]]
        return max((str(b.get("部署时间", "")) for b in owners), default="") or today

    for dev, rids in group_by_device({r: code_of.get(r, "") for r in online_ids}).items():
        rows = []
        for rid in rids:
            d = details.get(rid) or {}
            rows.append({"Rule ID": rid, "升级类型": d.get("升级类型") or "-", "机型信息": d.get("机型") or "-",
                         "机型扩展码": d.get("设备扩展信息") or "-", "源版本": d.get("源版本") or "-",
                         "目标版本": d.get("目标版本") or "-", "部署时间": deploy_date(rid)})
        print("")
        print(render_summary(dev, rows, "online"))
    missing = [r for r in online_ids if r not in details]
    if missing:
        print(color(f"  （{', '.join(missing)} 在部署库/来源 Excel 里没有明细，表中只有 Rule ID）", "dim"))


def online_by_rule_ids(specs: list[str], device_override: str = "") -> int:
    """`deployota online 50000-50005`：按 Rule ID 上线。

    设备类型代码取自部署库（--device 可覆盖；查不到的不处理）。按设备类型分组，每组走与 Excel 模式完全相同的
    run_online（待选区搜索 → 上线 → 确认；冲突时看被冲突规则的 mac 组非空才替换；已上线的自动跳过），
    组内回读状态，最后打印上线汇总和部署信息汇总表。
    """
    from .dass_perf_query import parse_ruleid_spec
    from .active_offline import device_codes_for, group_by_device
    from . import depdb
    ids: list[str] = []
    try:
        for s in specs:
            ids.extend(parse_ruleid_spec(s))
    except ValueError as e:
        print(color(f"✗ Rule ID 参数有误：{e}", "red"))
        return 2
    ids = list(dict.fromkeys(ids))
    from .device_lookup import from_cms_online_rules, print_sources, resolve_local
    code_of, src, missing = resolve_local(ids, device_override)
    print(color(f"[online] 按 Rule ID 上线：{len(ids)} 个", "bold"))
    print_sources(code_of, src, missing, color, lambda t: color("⚠ " + t, "yellow"))
    no_dev: list[str] = []

    final_all: list[dict] = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not config.HEADED)
        context = browser.new_context()
        page = context.new_page()
        try:
            login(page)
            if missing:
                # 『在线规则』页能查到的说明已经在线（后面 run_online 会自动跳过）；查不到的只能 --device。
                found = from_cms_online_rules(page, missing)
                code_of.update(found)
                for r in found:
                    print(f"  {r}：CMS『在线规则』查到设备类型代码 {found[r]}（已在线）")
                no_dev = [r for r in missing if r not in found]
                if no_dev:
                    print(color(f"⚠ 哪里都查不到设备类型代码，这些不处理：{', '.join(no_dev)}"
                                "（还没上线的规则 CMS 侧查不到，请用 --device 指定）", "yellow"))
            groups = group_by_device(code_of)
            for dev, rids in groups.items():
                print(color(f"\n======== 设备类型 {dev}：{len(rids)} 个 Rule ID ========", "bold"))
                results = run_online(page, rids, dev)
                # 回读要在切换设备类型之前做：已选区的搜索带着当前设备类型过滤。
                final_all += read_final_status(page, rids, results)
        except Exception as e:
            print(color(f"✗ 运行出错：{e}", "red"))
            _shot(page, "active_error")
            done = {r["ruleid"] for r in final_all}
            final_all += [{"ruleid": r, "status": STATUS_FAILED, "note": f"未处理：{e}"} for r in ids
                          if r not in done and r not in no_dev]
        finally:
            if config.HEADED:
                print("\n浏览器将在 20 秒后关闭（可自行查看页面）...")
                time.sleep(20)
            context.close()
            browser.close()
    print("\n" + summarize(final_all))
    _print_online_summary_by_ids([r["ruleid"] for r in final_all if r.get("status") == STATUS_ONLINE], code_of)
    if no_dev:
        print(color(f"⚠ 未处理（查不到设备类型代码，可用 --device 指定）：{', '.join(no_dev)}", "yellow"))
        return 2
    return summary_exit_code(final_all)


def _parse_cli(argv: list[str]) -> tuple[str, str | None]:
    """解析子命令 + 可选 Excel 路径。

    返回 (mode, excel)。mode ∈ {"collect", "online", "publish", "probe"}。
      publish [EXCEL]   collect + online 合并，3 个检查点。
      collect [EXCEL]   导入模板建规则 + 收集本批 Rule ID（升序打印）。
      online  [EXCEL]   从 Excel 的『Rule ID』列读本批 Rule ID → 逐条上线 + 冲突 + 汇总。
      --probe-online    只读 dump 编排视图选择器。
    不给 EXCEL 则取 OTA_V4_OUTPUTS 下最新的 .xlsx。
    """
    mode: str | None = None
    excel: str | None = None
    for arg in argv:
        if arg in ("-h", "--help"):
            print(_help_text(mode))
            raise SystemExit(0)
        if arg == "--probe-online":
            mode = "probe"
        elif arg in ("--selftest", "--no-dass"):
            continue  # --selftest 在 __main__ 处理；--no-dass 在 main 里读（只对 publish 生效）
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
        print("✗ 缺少子命令。用法：collect [EXCEL] | online [EXCEL] | publish [EXCEL] | --probe-online。详见 -h。")
        raise SystemExit(2)
    return mode, excel


def main() -> int:
    argv, rule_specs, device_override = split_rule_args(sys.argv[1:])
    mode, excel = _parse_cli(argv)
    if rule_specs or device_override:
        if mode != "online":
            print(color(f"✗ 只有 online 支持直接给 Rule ID（{mode} 请给 Excel）。", "red"))
            return 2
        if excel:
            print(color("✗ Rule ID 和 Excel 不能同时给，二选一。", "red"))
            return 2
        if not rule_specs:
            print(color("✗ 给了 --device 但没有 Rule ID。例：./deployota online 50000-50005", "red"))
            return 2
        return online_by_rule_ids(rule_specs, device_override)

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
                    print(color("✗ [collect] 未回收到任何 Rule ID。", "red"))
                    return 2
                print(f"\n[collect] 模板名称：{name}")
                print(f"[collect] 备注：{remark}")
                print(f"[collect] 本批 Rule ID（{len(rule_ids)} 个，升序）：")
                for rid in rule_ids:
                    print(f"  {rid}")

                # 回写：把『升级方式/升级日期/Rule ID』三列写回传入的同一个 Excel（紧跟 OTA类型 列后，
                # Rule ID 升序逐行、第 1 行放最小值；升级方式按 OTA类型 代码映射；升级日期=今天）。
                try:
                    n = write_rule_id_columns(excel, rule_ids)
                    print("[collect] " + color(
                        f"✓ 已把『升级方式/升级日期/Rule ID』三列回写到 Excel（{n} 行）：{excel}", "green"))
                except Exception as e:
                    print("[collect] " + color(f"✗ 回写 Excel 三列失败：{e}", "red"))
                    if config.HEADED:
                        print("\n浏览器将在 20 秒后关闭（可自行查看页面）...")
                        time.sleep(20)
                    return 2

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
                    print(color("✗ [online] Excel 的『Rule ID』列没有可用的 Rule ID（列缺失或为空），终止。", "red"))
                    return 2
                print(f"[online] 从 Excel 读到 {len(rule_ids)} 个 Rule ID（升序）：{rule_ids}")

                # 设备类型代码：上线搜索的第二个条件（设备类型代码 + Rule ID）。从 Excel『设备类型代码』列读。
                device_codes = {d for d in info.get("device_codes", set()) if str(d or "").strip()}
                if len(device_codes) == 1:
                    device_code = next(iter(device_codes))
                elif len(device_codes) == 0:
                    device_code = ""
                    print(color("⚠ [online] Excel『设备类型代码』列为空，上线搜索将只用 Rule ID（精度下降）。", "yellow"))
                else:
                    device_code = ""
                    print(color(f"⚠ [online] Excel 含多个设备类型代码 {sorted(device_codes)!r}，"
                                "无法确定唯一值，上线搜索将只用 Rule ID（精度下降）。", "yellow"))
                if device_code:
                    print(f"[online] 设备类型代码（搜索第二条件）：{device_code!r}")

                # 阶段2：逐条上线 + 冲突处理（按 设备类型代码 + Rule ID 搜索）。
                results = run_online(page, rule_ids, device_code)
                # 阶段3：状态汇总。
                final = read_final_status(page, rule_ids, results)
                print("\n" + summarize(final))
                _print_deploy_summary(excel, final)

                if config.HEADED:
                    print("\n浏览器将在 20 秒后关闭（可自行查看页面）...")
                    time.sleep(20)
                return summary_exit_code(final)

            # —— 子命令 publish：导入 + 取 Rule ID + 上线，三个检查点 ——
            if mode == "publish":
                return run_publish(page, excel, with_dass="--no-dass" not in sys.argv[1:])

            print(color(f"✗ 未知模式：{mode!r}", "red"))
            return 2
        except Exception as e:
            print(color(f"✗ 运行出错：{e}", "red"))
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
    dc, tv = validate_import_target({"MT9655AMAndrGTV"}, {"0100UQ1015"})
    check("D3 单一值通过返回 (dc, [tv])", dc == "MT9655AMAndrGTV" and tv == ["0100UQ1015"])
    check("D3 单一值忽略空串干扰",
          validate_import_target({"DEV", ""}, {"V", "  "}) == ("DEV", ["V"]))
    raised_dc = False
    try:
        validate_import_target({"A", "B"}, {"V"})
    except ValueError:
        raised_dc = True
    check("D3 多设备类型代码报错中止", raised_dc)
    check("D3 多目标版本允许，返回排序列表",
          validate_import_target({"A"}, {"V1.3.9.4", "V1.3.9.2"}) == ("A", ["V1.3.9.2", "V1.3.9.4"]))
    check("D1 多目标版本名称带全部末5位",
          build_rule_name("MT9666AndrWH80E", ["V1.3.9.4", "V1.3.9.2"]) == "MT9666AndrWH80E_3.9.2_3.9.4")
    raised_tv = False
    try:
        validate_import_target({"A"}, {"", "  "})
    except ValueError:
        raised_tv = True
    check("D3 目标版本全空报错中止", raised_tv)
    raised_empty = False
    try:
        validate_import_target(set(), {"V"})
    except ValueError:
        raised_empty = True
    check("D3 空设备类型代码报错（大小!=1）", raised_empty)

    # —— publish：检查点报告 ——
    rep_ok = format_checkpoint_report([{"n": n, "status": "pass", "detail": "x"} for n in (1, 2, 3)])
    check("publish 全通过含『全部检查点通过』", "全部检查点通过" in rep_ok and "✗" not in rep_ok)
    rep_bad = format_checkpoint_report([
        {"n": 1, "status": "pass"}, {"n": 2, "status": "fail", "detail": "数量不符"}, {"n": 3, "status": "skip"}])
    check("publish 失败报告点出检查点2且后续未执行",
          "卡在检查点2" in rep_bad and "检查点3 上线成功：未执行" in rep_bad and "数量不符" in rep_bad)
    check("publish 退出码 11/12/13/14", EXIT_CP_FAIL == {1: 11, 2: 12, 3: 13, 4: 14})
    check("publish 四项全通过文案含 DASS",
          "DASS init 已填报" in format_checkpoint_report([{"n": n, "status": "pass"} for n in (1, 2, 3, 4)]))
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
