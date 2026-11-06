"""depdb_backfill —— CMS 部署历史回补（【只读】抓取）。

登录 CMS → 进「整机升级管理 → OTA升级模板」列表 → 逐条读 标题/备注/创建人/创建时间 →
按创建人过滤 `devinw`（其余跳过）→ 点「查看规则」进二级页读 规则ID/设备类型代码/
设备扩展信息/源版本/目标版本/升级方式 → 构造一条 depdb 部署记录 →
经 depdb.append_record 去重合并写入 → 翻页直到列表「创建时间」早于 --since（默认 2026-04-01）即停。
备注含 VOID（忽略大小写）的记录以 force_status=STATUS_VOID 落为『作废』，其余按『有效』写入。

用法：
    cd ota_cms_automation
    ./.venv/bin/python -m automation.depdb_backfill --selftest        # 纯逻辑自测（不连浏览器，仅临时库）
    ./.venv/bin/python -m automation.depdb_backfill --help            # 列出全部参数（默认 dry-run）
    CMS_HEADED=true ./.venv/bin/python -m automation.depdb_backfill   # 由【用户】亲自跑的 HEADED 全量抓取（默认 dry-run）
    CMS_HEADED=true ./.venv/bin/python -m automation.depdb_backfill --commit  # 真正写库（需用户显式加 --commit）

===================== 安全红线（硬性，所有维护者必须遵守）=====================
- 【全程只读】：脚本对 CMS 只做 登录 / 翻页 / 点「查看规则」查看 / 读取文本。
  绝不点击 删除/上线/下线/新建/批量修改/导入/提交/替换 或任何写按钮。
- 【--dry-run 默认开启】：只抓取并打印逐条预览 + 统计，不写库；真正写库必须显式加 --commit。
- 【绝不清库】：只经 depdb.append_record 追加/去重/作废替换，绝不调用底层整体重写清空；
  真实部署库现有记录必须保留。
- 【实现期代理不代跑全量】：实现/验证只运行离线 --selftest（假数据 + OTA_DEPLOY_DB_DIR 临时库）。
  真正连生产 CMS 的 HEADED 全量抓取/入库留给用户亲自执行。
- 【绝不打印凭据值】：凭据只由 login 从 config/.env 读取，本模块不读取/不打印其值。

返回码：0 正常；1 运行出错；2 参数非法。
"""
from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime

from playwright.sync_api import sync_playwright

from . import config
from ._log import ok, warn, err, color
from .login import login, _shot
from .active_ota import (
    _open_template_list,
    _goto_next_page,
    norm_ruleid,
)
from . import depdb
from .summary import _rule_id_range

DEFAULT_SINCE = "2026-04-01"
DEFAULT_OWNER = "devinw"


# ======================================================================
# CLI 解析（仿 dass_perf_query._build_parser）
# ======================================================================

def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="python -m automation.depdb_backfill",
        description="CMS 部署历史回补（只读抓取 OTA升级模板 → 构造 depdb 记录 → 去重合并写入）。",
    )
    p.add_argument(
        "--since",
        default=DEFAULT_SINCE,
        help=f"回溯截止日期（YYYY-MM-DD）；列表『创建时间』早于它即停止翻页。默认 {DEFAULT_SINCE}。",
    )
    p.add_argument(
        "--owner",
        default=DEFAULT_OWNER,
        help=f"只处理该创建人的模板，其余跳过。默认 {DEFAULT_OWNER}。",
    )
    mode = p.add_mutually_exclusive_group()
    mode.add_argument(
        "--commit",
        dest="commit",
        action="store_true",
        help="真正写库（经 append_record 去重合并）。不加时默认 dry-run，只打印预览+统计不写库。",
    )
    mode.add_argument(
        "--dry-run",
        dest="dry_run",
        action="store_true",
        help="只抓取并打印逐条预览+统计，不写库（默认行为，显式写出可读性更好）。",
    )
    p.add_argument(
        "--max-pages",
        type=int,
        default=100,
        help="翻页保护上限，防失控翻页。默认 100。",
    )
    p.add_argument(
        "--limit",
        type=int,
        default=None,
        help="仅抓前 N 个 owner 模板（调试用）；不传则不限。",
    )
    p.add_argument(
        "--selftest",
        action="store_true",
        help="运行纯逻辑自测（不连浏览器，仅写临时库），完成后退出。",
    )
    return p


# ======================================================================
# 纯逻辑函数（不连浏览器，--selftest 覆盖全部分支）
# ======================================================================

def remark_has_void(remark: str) -> bool:
    """备注是否含 VOID（忽略大小写）。例 '20261104_USA_8_VOID' / '...void' → True。"""
    return "void" in str(remark or "").lower()


# 备注里的 8 位日期（20YYMMDD，如 20261105）。
_DATE8_RE = re.compile(r"(20\d{6})")


def parse_created(created: str) -> datetime | None:
    """多格式容错解析创建时间字符串为 datetime；无法解析返回 None。

    支持 'YYYY-MM-DD HH:MM:SS' / 'YYYY-MM-DD' / 'YYYY/MM/DD HH:MM:SS' / 'YYYY/MM/DD'。
    """
    s = str(created or "").strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d",
                "%Y/%m/%d %H:%M:%S", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def deploy_date_from(remark: str, created: str) -> str:
    """部署时间（YYYYMMDD）：优先备注里的 8 位日期，取不到则取创建时间的日期部分。

    例 ('20261105_USA_1', '2026-11-06 10:00:00') → '20261105'；
       ('USA_batch',      '2026-11-05 22:58:08') → '20261105'。
    两者都取不到返回空串。
    """
    m = _DATE8_RE.search(str(remark or ""))
    if m:
        return m.group(1)
    dt = parse_created(created)
    if dt is not None:
        return dt.strftime("%Y%m%d")
    return ""


def is_before_since(created: str, since_str: str) -> bool:
    """创建时间是否【早于】--since。

    - 创建时间解析失败 → 返回 False（保守按『不早于』处理，宁可多抓也不漏抓）。
    - since 解析失败 → 返回 False（不截断）。
    - 严格『早于』：等于 since 当日 00:00:00 不算早于。
    """
    dt = parse_created(created)
    if dt is None:
        return False
    since = parse_created(since_str)
    if since is None:
        return False
    return dt < since


def _first_nonempty(rule_rows: list[dict], key: str) -> str:
    """二级页规则行里某列的首个非空值（去空白）。"""
    for r in rule_rows:
        v = str(r.get(key, "") or "").strip()
        if v:
            return v
    return ""


def build_backfill_record(row: dict, rule_rows: list[dict], owner: str = DEFAULT_OWNER) -> dict:
    """由模板行 + 二级页规则行构造一条 depdb 记录字典（§E 字段映射）。只构造 dict，不写库。

    row       = {title, creator, remark, created}
    rule_rows = [{规则ID, 设备类型代码, 设备扩展信息, 源版本, 目标版本, 升级方式}, ...]

    字段：
      - 设备类型代码/源版本/目标版本/升级类型（←升级方式）：取二级页首行非空值；
      - 设备扩展码集合：各行『设备扩展信息』去重排序后 ','.join（口径同 build_record_from_excel.ext_set）；
      - Rule ID列表：各行『规则ID』norm_ruleid 去重、按数值升序 ','.join；
      - Rule ID范围：复用 summary._rule_id_range（lo==hi 则单值，否则 'lo ~ hi'）；
      - Owner：优先模板行创建人，缺省用 owner 入参；
      - 部署时间：deploy_date_from（备注优先/回退创建时间）；
      - 设备类型名称：''（二级页无该列）；
      - Excel路径：f'CMS:{标题}'（标记来源为 CMS 抓取）；
      - 批次号：f'{设备类型代码}_{范围去空格}_{部署时间}'；
      - 含 VOID：附『作废时间』= 部署时间或创建时间（append_record 落库时再兜底）。
    """
    title = str(row.get("title", "") or "").strip()
    creator = str(row.get("creator", "") or "").strip()
    remark = str(row.get("remark", "") or "")
    created = str(row.get("created", "") or "").strip()

    device_code = _first_nonempty(rule_rows, "设备类型代码")
    src_version = _first_nonempty(rule_rows, "源版本")
    tgt_version = _first_nonempty(rule_rows, "目标版本")
    upgrade_type = _first_nonempty(rule_rows, "升级方式")

    # Rule ID：归一、去重、数值升序。
    rids: list[str] = []
    seen: set[str] = set()
    for r in rule_rows:
        rid = norm_ruleid(str(r.get("规则ID", "") or ""))
        if rid and rid not in seen:
            seen.add(rid)
            rids.append(rid)
    rids.sort(key=lambda x: int(x) if x.isdigit() else 0)

    # Rule ID 范围：喂 _rule_id_range 期望的 [{"Rule ID": ...}] 形态。
    lo, hi = _rule_id_range([{"Rule ID": x} for x in rids])
    rng = lo if lo == hi else f"{lo} ~ {hi}"

    # 设备扩展码集合：去重排序拼接。
    ext_set = sorted({
        str(r.get("设备扩展信息", "") or "").strip()
        for r in rule_rows
        if str(r.get("设备扩展信息", "") or "").strip()
    })

    deploy_time = deploy_date_from(remark, created)

    rec = {
        "设备类型代码": device_code,
        "设备类型名称": "",
        "源版本": src_version,
        "目标版本": tgt_version,
        "升级类型": upgrade_type,
        "设备扩展码集合": ",".join(ext_set),
        "Rule ID列表": ",".join(rids),
        "Rule ID范围": rng,
        "规则条数": len(rids),
        "Owner": creator or owner,
        "部署时间": deploy_time,
        "批次号": f"{device_code}_{rng.replace(' ', '')}_{deploy_time}",
        "作废时间": "",
        "Excel路径": f"CMS:{title}",
    }
    if remark_has_void(remark):
        # 作废时间优先用部署时间，缺省用创建时间（append_record 落库时仍会兜底填 now）。
        rec["作废时间"] = deploy_time or created
    return rec


# ======================================================================
# 只读浏览器读取函数（读列范式照抄 active_ota._collect_rule_ids 的 page.evaluate）
# ======================================================================

def read_template_rows(page) -> list[dict]:
    """读 OTA升级模板列表页当前页每行 {title, creator, remark, created}。

    用一次 page.evaluate：vxe/普通表 表头文本→列索引（『标题』『创建人』『备注』、
    时间列同时试『创建时间』或『导入时间』），逐行取 innerText。只读，不点任何按钮。
    """
    return page.evaluate(
        """() => {
            const tables = Array.from(document.querySelectorAll('.vxe-table, table'));
            const out = [];
            const seenRows = new Set();
            for (const t of tables) {
                const heads = t.querySelectorAll('thead th, thead td');
                if (!heads.length) continue;
                const names = Array.from(heads).map(h => (h.innerText||'').trim());
                const idxTitle = names.findIndex(x => x.indexOf('标题') >= 0);
                const idxCreator = names.findIndex(x => x.indexOf('创建人') >= 0);
                const idxRemark = names.findIndex(x => x.indexOf('备注') >= 0);
                let idxTime = names.findIndex(x => x.indexOf('创建时间') >= 0);
                if (idxTime < 0) idxTime = names.findIndex(x => x.indexOf('导入时间') >= 0);
                if (idxTitle < 0) continue;  // 不是模板列表
                const rows = t.querySelectorAll('tbody tr');
                for (const r of rows) {
                    const cells = r.querySelectorAll('td, th');
                    if (!cells.length || !cells[idxTitle]) continue;
                    const get = (i) => (i >= 0 && cells[i]) ? (cells[i].innerText||'').trim() : '';
                    const title = get(idxTitle);
                    if (!title) continue;
                    const key = title + '||' + get(idxTime);
                    if (seenRows.has(key)) continue;  // vxe 固定列可能重复渲染
                    seenRows.add(key);
                    out.push({
                        title: title,
                        creator: get(idxCreator),
                        remark: get(idxRemark),
                        created: get(idxTime),
                    });
                }
            }
            return out;
        }"""
    )


def _read_rule_rows_current_page(page) -> list[dict]:
    """读二级『查看规则』页当前页每行的各列（表头文本→列索引，范式同 _collect_rule_ids）。

    列：规则ID/设备类型代码/设备扩展信息/源版本/目标版本/升级方式。只读。
    """
    return page.evaluate(
        """() => {
            const tables = Array.from(document.querySelectorAll('.vxe-table, table'));
            const out = [];
            const seenRows = new Set();
            for (const t of tables) {
                const heads = t.querySelectorAll('thead th, thead td');
                if (!heads.length) continue;
                const names = Array.from(heads).map(h => (h.innerText||'').trim());
                const idxRule = names.findIndex(x => x.indexOf('规则ID') >= 0);
                if (idxRule < 0) continue;  // 不是规则表
                const idxCode = names.findIndex(x => x.indexOf('设备类型代码') >= 0);
                const idxExt = names.findIndex(x => x.indexOf('设备扩展信息') >= 0);
                const idxSrc = names.findIndex(x => x.indexOf('源版本') >= 0);
                const idxTgt = names.findIndex(x => x.indexOf('目标版本') >= 0);
                const idxWay = names.findIndex(x => x.indexOf('升级方式') >= 0);
                const rows = t.querySelectorAll('tbody tr');
                for (const r of rows) {
                    const cells = r.querySelectorAll('td, th');
                    if (!cells.length) continue;
                    const get = (i) => (i >= 0 && cells[i]) ? (cells[i].innerText||'').trim() : '';
                    const rid = get(idxRule);
                    if (!rid) continue;
                    const key = rid + '||' + get(idxExt);
                    if (seenRows.has(key)) continue;
                    seenRows.add(key);
                    out.push({
                        '规则ID': rid,
                        '设备类型代码': get(idxCode),
                        '设备扩展信息': get(idxExt),
                        '源版本': get(idxSrc),
                        '目标版本': get(idxTgt),
                        '升级方式': get(idxWay),
                    });
                }
            }
            return out;
        }"""
    )


def _click_view_rules_by_row_index(page, data_idx: int) -> dict:
    """按【行索引】点模板列表第 data_idx 个数据行自己的行内「查看规则」。只读。

    与 read_template_rows 完全同口径：锚「表头含『标题』」的模板表，对有非空标题的
    tbody tr 按出现顺序计数（跳过无标题行、去 vxe 固定列重复），命中目标 data_idx 时
    【只在该行元素内】找文本含『查看规则』的 button/a/span 并 .click()。
    绝不用顶部搜索框、绝不按标题文本跨行匹配（重名标题无法区分）。

    返回 {"ok": bool, "title": str}——title 为被点击行的标题，供调用方校验是否与期望一致。
    """
    return page.evaluate(
        """(targetIdx) => {
            const tables = Array.from(document.querySelectorAll('.vxe-table, table'));
            const seenRows = new Set();
            let dataIdx = -1;
            for (const t of tables) {
                const heads = t.querySelectorAll('thead th, thead td');
                if (!heads.length) continue;
                const names = Array.from(heads).map(h => (h.innerText||'').trim());
                const idxTitle = names.findIndex(x => x.indexOf('标题') >= 0);
                let idxTime = names.findIndex(x => x.indexOf('创建时间') >= 0);
                if (idxTime < 0) idxTime = names.findIndex(x => x.indexOf('导入时间') >= 0);
                if (idxTitle < 0) continue;  // 不是模板列表
                const rows = t.querySelectorAll('tbody tr');
                for (const r of rows) {
                    const cells = r.querySelectorAll('td, th');
                    if (!cells.length || !cells[idxTitle]) continue;
                    const get = (i) => (i >= 0 && cells[i]) ? (cells[i].innerText||'').trim() : '';
                    const title = get(idxTitle);
                    if (!title) continue;
                    const key = title + '||' + get(idxTime);
                    if (seenRows.has(key)) continue;  // vxe 固定列可能重复渲染
                    seenRows.add(key);
                    dataIdx += 1;
                    if (dataIdx !== targetIdx) continue;
                    // 命中目标行：只在该行内找『查看规则』元素并点击。
                    const els = r.querySelectorAll('button, a, span');
                    for (const el of els) {
                        if ((el.innerText||'').indexOf('查看规则') >= 0) {
                            el.click();
                            return {ok: true, title: title};
                        }
                    }
                    return {ok: false, title: title};  // 该行无『查看规则』
                }
            }
            return {ok: false, title: ''};  // 索引越界/未找到模板表
        }""",
        data_idx,
    )


def _return_to_template_list(page) -> str:
    """二级『查看规则』页返回模板列表。返回 'tab'（点 antd tab 返回，保持翻页位置）或
    'reopen'（兜底重开列表，会回到第 1 页）。

    首选点 antd tab『OTA升级模板』——实测保持当前翻页位置，适合逐行遍历；点击异常/
    tab 不在场时兜底调 _open_template_list（会重置回第 1 页，仅异常兜底用）。
    """
    try:
        tab = page.locator(".ant-tabs-tab", has_text="OTA升级模板").first
        if tab.count() > 0 and tab.is_visible():
            tab.click()
            page.wait_for_timeout(1000)
            return "tab"
    except Exception:
        pass
    _open_template_list(page)
    return "reopen"


def collect_rule_rows(page, row_idx: int, title: str, max_pages: int = 100) -> list[dict]:
    """点模板列表第 row_idx 个数据行自己的「查看规则」进二级页，读全部规则行（含翻页累计）。只读。

    按【行索引】点该行行内『查看规则』（_click_view_rules_by_row_index，不搜标题、不跨行匹配，
    可区分重名标题），读当前页规则行后用 _goto_next_page 翻页累计，_shot 截图关键步骤。
    读完由调用方 _return_to_template_list 返回列表。title 仅用于截图命名与点中行一致性校验。
    """
    clicked = _click_view_rules_by_row_index(page, row_idx)
    if not clicked.get("ok"):
        print(warn(f"未能在模板列表第 {row_idx} 行点开行内『查看规则』：{title!r}"
                   f"（点中标题={clicked.get('title','')!r}，跳过该模板）。"))
        return []
    clicked_title = str(clicked.get("title", "") or "").strip()
    if clicked_title and title and clicked_title != title:
        print(warn(f"第 {row_idx} 行点中标题 {clicked_title!r} 与预期 {title!r} 不符"
                   f"（行索引口径漂移，跳过该模板）。"))
        return []
    page.wait_for_timeout(1500)

    rows: list[dict] = []
    seen: set[tuple] = set()
    guard = 0
    while guard < max_pages:
        guard += 1
        batch = _read_rule_rows_current_page(page)
        for r in batch:
            key = (norm_ruleid(str(r.get("规则ID", ""))), str(r.get("设备扩展信息", "")))
            if key in seen:
                continue
            seen.add(key)
            rows.append(r)
        if not _goto_next_page(page):
            break
        page.wait_for_timeout(1000)

    _shot(page, f"backfill_rules_{norm_ruleid(title) or 'tpl'}")
    return rows


# ======================================================================
# 主循环（只读遍历 + dry-run/commit 写库）
# ======================================================================

def run_backfill(page, args) -> dict:
    """遍历 OTA升级模板列表，过滤 owner，构造并（dry-run 预览 / commit 写库）记录。返回统计 dict。"""
    stats = {
        "scanned": 0,        # 扫描到的模板行数
        "owner_hits": 0,     # 命中 owner 的模板数
        "built": 0,          # 成功构造的记录数
        "written": 0,        # 新写入
        "skipped_dup": 0,    # 完全重复跳过
        "replaced": 0,       # 作废替换（同身份键新 Rule ID）
        "voided": 0,         # VOID 作废落库
        "empty_rules": 0,    # 二级页读不到规则行、跳过
        "void_empty": 0,     # VOID 批次且二级页空（规则已清空）、跳过（含在 empty_rules 内）
        "platforms": set(),  # 设备类型代码集合
        "rule_ids": set(),   # Rule ID 全集（去重）
    }

    if not _open_template_list(page):
        raise RuntimeError("未能进入「OTA升级模板」模板列表页。")

    page_no = 0
    while page_no < args.max_pages:
        page_no += 1
        rows = read_template_rows(page)
        print(color(f"── 第 {page_no} 页：读到 {len(rows)} 行模板 ──", "bold"))
        stats["scanned"] += len(rows)

        # 命中 owner 的目标行（先一次性读入内存再逐个进二级页，避免 DOM 丢上下文）。
        # 每项记下该行在本页数据行里的【索引 i】（read_template_rows 返回顺序的下标），
        # 逐行处理时用该索引点行内『查看规则』——与 _click_view_rules_by_row_index 同口径。
        targets: list[dict] = []
        all_before_since = len(rows) > 0
        for i, row in enumerate(rows):
            created = row.get("created", "")
            if not is_before_since(created, args.since):
                all_before_since = False
            creator = str(row.get("creator", "") or "").strip()
            if creator != args.owner:
                continue  # 其他创建人跳过
            if is_before_since(created, args.since):
                continue  # 早于 since 的行跳过（整页都早则下方判停）
            targets.append({"row": row, "idx": i})

        stats["owner_hits"] += len(targets)

        for item in targets:
            if args.limit is not None and stats["built"] >= args.limit:
                print(warn(f"已达 --limit={args.limit}，停止抓取。"))
                return stats
            row = item["row"]
            row_idx = item["idx"]
            title = str(row.get("title", "") or "").strip()
            remark = str(row.get("remark", "") or "")
            print(color(f"  · 处理模板[行{row_idx}]：{title!r}  备注={remark!r}  创建={row.get('created','')}", "dim"))

            rule_rows = collect_rule_rows(page, row_idx, title, max_pages=args.max_pages)
            if not rule_rows:
                stats["empty_rules"] += 1
                if remark_has_void(remark):
                    stats["void_empty"] += 1
                # 读完二级页后返回列表（保持翻页位置），继续下一行。
                _return_to_template_list(page)
                continue

            rec = build_backfill_record(row, rule_rows, owner=args.owner)
            stats["built"] += 1
            if rec.get("设备类型代码"):
                stats["platforms"].add(rec["设备类型代码"])
            for x in str(rec.get("Rule ID列表", "")).split(","):
                if x.strip():
                    stats["rule_ids"].add(x.strip())

            is_void = remark_has_void(remark)
            tag = color("[VOID→作废]", "red") if is_void else color("[有效]", "green")
            print(f"    {tag} 设备={rec['设备类型代码']} {rec['源版本']} → {rec['目标版本']} "
                  f"[{rec['升级类型']}]  Rule ID {rec['Rule ID范围']}（{rec['规则条数']} 条）  "
                  f"部署={rec['部署时间']}  来源={rec['Excel路径']}")

            if args.commit:
                wrote, msg = depdb.append_record(
                    rec, force_status=depdb.STATUS_VOID if is_void else None
                )
                if is_void:
                    stats["voided"] += 1
                if wrote:
                    if "作废" in msg and "旧记录" in msg:
                        stats["replaced"] += 1
                    stats["written"] += 1
                    print(ok(f"    写库：{msg}"))
                else:
                    stats["skipped_dup"] += 1
                    print(color(f"    跳过：{msg}", "dim"))
            else:
                print(color("    （dry-run：未写库）", "dim"))

            # 读完二级页后返回列表（保持翻页位置），继续下一行。
            _return_to_template_list(page)

        # 截止判断：整页所有行都早于 since → 停止翻页（已进入更早历史）。
        if all_before_since:
            print(warn(f"本页所有模板创建时间均早于 --since={args.since}，停止翻页。"))
            break

        if not _goto_next_page(page):
            print(color("已到最后一页（无下一页）。", "dim"))
            break
        page.wait_for_timeout(1000)
    else:
        print(warn(f"已达 --max-pages={args.max_pages} 翻页保护上限，停止。"))

    return stats


def _print_summary(stats: dict, committed: bool) -> None:
    print("")
    print(color("────────── 回补总体统计 / Backfill Summary ──────────", "bold"))
    print(f"  扫描模板行：{stats['scanned']}    命中 owner：{stats['owner_hits']}    "
          f"构造记录：{stats['built']}")
    if committed:
        print(f"  新写入：{color(str(stats['written']), 'green')}    "
              f"跳过重复：{stats['skipped_dup']}    "
              f"作废替换：{stats['replaced']}    "
              f"VOID 作废：{color(str(stats['voided']), 'red')}")
    print(f"  二级页无规则跳过：{stats['empty_rules']}"
          f"（其中 VOID 批次规则已清空：{stats.get('void_empty', 0)}）")
    platforms = sorted(stats["platforms"])
    rid_nums = sorted(int(x) for x in stats["rule_ids"] if x.isdigit())
    rid_span = f"{rid_nums[0]} ~ {rid_nums[-1]}" if rid_nums else "-"
    print(f"  平台数（设备类型）：{color(str(len(platforms)), 'bold')}"
          f"    Rule ID 总数：{color(str(len(stats['rule_ids'])), 'green')}"
          f"    范围：{color(rid_span, 'green')}")
    if platforms:
        print(f"  平台：{', '.join(platforms)}")
    if not committed:
        print(warn("当前为 dry-run：未写库。要真正写入请显式加 --commit。"))


# ======================================================================
# 运行入口
# ======================================================================

def _run(args) -> int:
    print(f"回补参数：--since={args.since}  --owner={args.owner}  "
          f"{'--commit（写库）' if args.commit else 'dry-run（默认，不写库）'}  "
          f"--max-pages={args.max_pages}" + (f"  --limit={args.limit}" if args.limit else ""))
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not config.HEADED)
        context = browser.new_context()
        page = context.new_page()
        # 只读抓取：不注册任何 dialog 应答，不点任何写按钮。
        try:
            login(page)
            print(f"打开 OTA 升级模板页：{config.CMS_OTA_URL}")
            page.goto(config.CMS_OTA_URL, wait_until="domcontentloaded")
            page.wait_for_timeout(2500)
            stats = run_backfill(page, args)
            _print_summary(stats, committed=args.commit)
            _shot(page, "backfill_done")
            return 0
        except Exception as e:
            print(err(f"回补出错：{e}"))
            _shot(page, "backfill_error")
            return 1
        finally:
            context.close()
            browser.close()


def main() -> int:
    try:
        args = _build_parser().parse_args()
    except SystemExit as e:
        # argparse 的 --help 返回 0，参数错误返回 2。
        return int(e.code) if isinstance(e.code, int) else 2

    if args.selftest:
        return _selftest()

    # 互斥组：未给任一则默认 dry-run。commit 优先于 dry_run（互斥组保证不会同给）。
    if not getattr(args, "commit", False):
        args.commit = False
    return _run(args)


# ======================================================================
# 纯逻辑 + depdb 集成自测（不连浏览器；仅写 OTA_DEPLOY_DB_DIR 临时库）
# ======================================================================

def _selftest() -> int:
    import os
    import tempfile

    ok_all = True

    def check(name, cond):
        nonlocal ok_all
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        if not cond:
            ok_all = False

    # —— remark_has_void：大小写不敏感 ——
    check("VOID 大写识别", remark_has_void("20261104_USA_8_VOID") is True)
    check("void 小写识别", remark_has_void("20261104_usa_8_void") is True)
    check("混合大小写 Void 识别", remark_has_void("x_Void_y") is True)
    check("无 VOID 返回 False", remark_has_void("20261105_USA_1") is False)
    check("空备注返回 False", remark_has_void("") is False)

    # —— deploy_date_from：备注优先 / 回退创建时间 ——
    check("备注含 8 位日期优先", deploy_date_from("20261105_USA_1", "2026-11-06 10:00:00") == "20261105")
    check("备注无日期回退创建时间(带时分秒)", deploy_date_from("USA_batch", "2026-11-05 22:58:08") == "20261105")
    check("备注无日期回退创建时间(仅日期)", deploy_date_from("USA_batch", "2026-11-05") == "20261105")
    check("两者都无返回空串", deploy_date_from("USA", "") == "")

    # —— parse_created：多格式 ——
    check("解析 YYYY-MM-DD HH:MM:SS", parse_created("2026-11-05 22:58:08") is not None)
    check("解析 YYYY-MM-DD", parse_created("2026-11-05") is not None)
    check("解析 YYYY/MM/DD HH:MM:SS", parse_created("2026/11/05 22:58:08") is not None)
    check("无法解析返回 None", parse_created("not-a-date") is None)
    check("空串返回 None", parse_created("") is None)

    # —— is_before_since：边界 ——
    check("等于 since 当日不算早于", is_before_since("2026-04-01 00:00:00", "2026-04-01") is False)
    check("早一天算早于", is_before_since("2026-03-31 23:59:59", "2026-04-01") is True)
    check("晚于 since 不算早于", is_before_since("2026-05-01", "2026-04-01") is False)
    check("创建时间解析失败不漏抓(False)", is_before_since("garbage", "2026-04-01") is False)
    check("since 解析失败不截断(False)", is_before_since("2020-01-01", "bad-since") is False)

    # —— build_backfill_record：字段映射 ——
    row = {
        "title": "MT8195Andr32P1S_Q1031",
        "creator": "devinw",
        "remark": "20261105_USA_3",
        "created": "2026-11-06 22:58:08",
    }
    rule_rows = [
        {"规则ID": "65,152", "设备类型代码": "MT8195Andr32P1S", "设备扩展信息": "B",
         "源版本": "0100AQ0723", "目标版本": "0100AQ1031", "升级方式": "正常升级"},
        {"规则ID": "65150", "设备类型代码": "MT8195Andr32P1S", "设备扩展信息": "A",
         "源版本": "0100AQ0723", "目标版本": "0100AQ1031", "升级方式": "正常升级"},
        {"规则ID": "65150", "设备类型代码": "MT8195Andr32P1S", "设备扩展信息": "A",
         "源版本": "", "目标版本": "", "升级方式": ""},  # 重复 rule id + 空列，验证首行非空/去重
    ]
    rec = build_backfill_record(row, rule_rows, owner="devinw")
    check("设备类型代码取首行非空", rec["设备类型代码"] == "MT8195Andr32P1S")
    check("源版本取首行非空", rec["源版本"] == "0100AQ0723")
    check("目标版本取首行非空", rec["目标版本"] == "0100AQ1031")
    check("升级类型来自升级方式首行非空", rec["升级类型"] == "正常升级")
    check("Rule ID 归一去重升序", rec["Rule ID列表"] == "65150,65152")
    check("Rule ID 范围正确", rec["Rule ID范围"] == "65150 ~ 65152")
    check("规则条数=2", rec["规则条数"] == 2)
    check("设备扩展码集合去重排序", rec["设备扩展码集合"] == "A,B")
    check("设备类型名称留空", rec["设备类型名称"] == "")
    check("Excel路径=CMS:标题", rec["Excel路径"] == "CMS:MT8195Andr32P1S_Q1031")
    check("部署时间来自备注 8 位日期", rec["部署时间"] == "20261105")
    check("Owner 取创建人", rec["Owner"] == "devinw")
    check("批次号拼接", rec["批次号"] == "MT8195Andr32P1S_65150~65152_20261105")
    check("非 VOID 作废时间为空", rec["作废时间"] == "")

    # creator 为空时 Owner 回退到 owner 入参。
    rec_noowner = build_backfill_record(
        {**row, "creator": ""}, rule_rows, owner="fallbackowner"
    )
    check("creator 空时 Owner 回退入参", rec_noowner["Owner"] == "fallbackowner")

    # VOID 记录带作废时间。
    rec_void = build_backfill_record(
        {**row, "remark": "20261104_USA_8_VOID"}, rule_rows, owner="devinw"
    )
    check("VOID 记录作废时间=部署时间", rec_void["作废时间"] == "20261104")

    # 范围 lo==hi → 单值。
    rec_single = build_backfill_record(
        row, [{"规则ID": "65150", "设备类型代码": "D", "设备扩展信息": "A",
               "源版本": "s", "目标版本": "t", "升级方式": "正常升级"}], owner="devinw"
    )
    check("单个 Rule ID 范围为单值(非 a ~ b)", rec_single["Rule ID范围"] == "65150")

    # —— depdb 集成自测（仅临时库，绝不触碰真实库）——
    tmp = tempfile.mkdtemp(prefix="depdb_backfill_selftest_")
    old = os.environ.get("OTA_DEPLOY_DB_DIR")
    os.environ["OTA_DEPLOY_DB_DIR"] = tmp
    try:
        print(f"[selftest] depdb 临时库目录 = {tmp}")
        # 普通有效记录。
        normal = build_backfill_record(row, rule_rows, owner="devinw")
        wrote, _ = depdb.append_record(normal)
        check("普通记录写入返回新写入", wrote)
        recs = depdb.load_all()
        check("普通记录状态为『有效』", len(recs) == 1 and recs[-1].get("状态") == depdb.STATUS_ACTIVE)

        # VOID 记录（force_status=STATUS_VOID），用不同身份键避免与情况2耦合。
        void_row = {**row, "title": "DEV_Y_tpl", "remark": "20261104_USA_8_VOID"}
        void_rule_rows = [
            {"规则ID": "70000", "设备类型代码": "DEV_Y", "设备扩展信息": "Z",
             "源版本": "1.0", "目标版本": "2.0", "升级方式": "强制升级"},
        ]
        void_rec = build_backfill_record(void_row, void_rule_rows, owner="devinw")
        wrote, _ = depdb.append_record(
            void_rec, force_status=depdb.STATUS_VOID if remark_has_void(void_row["remark"]) else None
        )
        check("VOID 记录写入返回新写入", wrote)
        recs = depdb.load_all()
        last = recs[-1]
        check("VOID 记录末条状态为『作废』", last.get("状态") == depdb.STATUS_VOID)
        check("VOID 记录『作废时间』非空", bool(str(last.get("作废时间", "")).strip()))
        check("未污染真实库(临时库共 2 条)", len(recs) == 2)
    finally:
        if old is None:
            os.environ.pop("OTA_DEPLOY_DB_DIR", None)
        else:
            os.environ["OTA_DEPLOY_DB_DIR"] = old

    # —— _build_parser：默认 dry-run（commit=False）、参数齐全 ——
    parser = _build_parser()
    ns = parser.parse_args([])
    check("默认 since", ns.since == DEFAULT_SINCE)
    check("默认 owner", ns.owner == DEFAULT_OWNER)
    check("默认未加 --commit", ns.commit is False)
    check("默认 max-pages=100", ns.max_pages == 100)
    check("默认 limit=None", ns.limit is None)
    ns_commit = parser.parse_args(["--commit"])
    check("--commit 置 True", ns_commit.commit is True)

    print("\n自测结果：", "全部通过 ✅" if ok_all else "存在失败 ❌")
    return 0 if ok_all else 1


if __name__ == "__main__":
    if "--selftest" in sys.argv[1:]:
        sys.exit(_selftest())
    sys.exit(main())
