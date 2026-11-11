"""deployota summary —— 从【更新后的 Excel】（已含 Rule ID）打印 OTA 部署信息汇总。

数据全部来自更新后的 Excel（collect 回写 Rule ID 后的那份，sheet『升级规则』），不连任何网络。
打印内容对应部署汇总表：设备类型代码 / 部署规则总数 / Rule ID 范围，逐条列
Rule ID / 机型信息 / 机型扩展码 / 源版本 / 目标版本 / 部署时间，末尾灰度提醒。

用法：
    ./deployota summary <更新后的excel>
    ./deployota summary --file <更新后的excel>
    ./.venv/bin/python -m automation.summary <excel>
"""
from __future__ import annotations

import argparse
import re
import sys

from openpyxl import load_workbook

from . import config as cms_config  # noqa: F401  触发 config/.env 加载（与其它命令一致）
from ._log import ok, warn, err, color

# 汇总表列 -> Excel 列名（按更新后的 Excel 实际表头；列表 = 依次尝试的候选列名）。
# 目标版本显示 Excel『目标完整版本』列（形如 0100UQ1104）；老 Excel 没有该列时退回『目标版本』列。
_COLS = [
    ("Rule ID", "Rule ID"),
    ("升级类型", "升级方式"),
    ("机型信息", "内部机型信息"),
    ("机型扩展码", "设备扩展信息"),
    ("源版本", "源版本"),
    ("目标版本", ["目标完整版本", "目标版本"]),
    ("部署时间", "升级日期"),
]
# 不显示、但供部署库沿用的列：Excel『目标版本』（版本变化，形如 Q0305_to_Q1104）。
# 部署库的去重身份键一直用它，改成完整版本会让同一批被当成新批次重复入库，所以单独保留。
_EXTRA_COLS = [("版本变化", "目标版本")]

_RULEID_CANDIDATES = ["Rule ID", "ruleid", "RuleID"]


def _read_rows(xlsx_path: str):
    """读更新后 Excel 的『升级规则』sheet，返回 (device_code, rows[list[dict]])。

    rows 每项含汇总表需要的 6 列文本；仅保留 Rule ID 非空的真实数据行。
    """
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb["升级规则"] if "升级规则" in wb.sheetnames else wb.worksheets[0]
    hdr = [str(ws.cell(1, c).value or "").strip() for c in range(1, ws.max_column + 1)]
    idx = {h: i for i, h in enumerate(hdr)}

    def col_index(name_or_candidates):
        if isinstance(name_or_candidates, list):
            for n in name_or_candidates:
                if n in idx:
                    return idx[n]
            return None
        return idx.get(name_or_candidates)

    rid_i = col_index(_RULEID_CANDIDATES)
    dc_i = col_index("设备类型代码")
    if rid_i is None:
        raise ValueError("该 Excel 未找到『Rule ID』列——它需要是 collect 回写 Rule ID 后的【更新后 Excel】。")

    def _cands(src):
        if src == "Rule ID":
            return _RULEID_CANDIDATES
        return src if isinstance(src, list) else [src]
    col_i = {disp: col_index(_cands(src)) for disp, src in _COLS + _EXTRA_COLS}

    rows = []
    device_code = ""
    for r in range(2, ws.max_row + 1):
        def cell(i):
            return str(ws.cell(r, i + 1).value).strip() if i is not None and ws.cell(r, i + 1).value is not None else ""
        rid = cell(rid_i)
        if not rid:
            continue
        if not device_code and dc_i is not None:
            device_code = cell(dc_i)
        rows.append({disp: cell(col_i[disp]) for disp, _src in _COLS + _EXTRA_COLS})
    return device_code, rows


def _rule_id_range(rows):
    """返回 Rule ID 的 (最小, 最大) 文本；非纯数字时按出现顺序取首尾。"""
    rids = [r.get("Rule ID", "").strip() for r in rows if r.get("Rule ID", "").strip()]
    if not rids:
        return "", ""
    nums = [x for x in rids if x.isdigit()]
    if len(nums) == len(rids):
        return min(nums, key=int), max(nums, key=int)
    return rids[0], rids[-1]


def _wcwidth(s) -> int:
    """按显示宽度计长度：CJK/全角字符记 2，其余记 1。"""
    w = 0
    for ch in str(s):
        o = ord(ch)
        # CJK 统一表意、全角标点、假名、全角符号等记 2 宽。
        if (0x1100 <= o <= 0x115F or 0x2E80 <= o <= 0xA4CF or 0xAC00 <= o <= 0xD7A3
                or 0xF900 <= o <= 0xFAFF or 0xFE30 <= o <= 0xFE4F
                or 0xFF00 <= o <= 0xFF60 or 0xFFE0 <= o <= 0xFFE6
                or 0x1F300 <= o <= 0x1FAFF or 0x2600 <= o <= 0x27BF):
            w += 2
        else:
            w += 1
    return w


def _pad(s, w):
    return str(s) + " " * max(0, w - _wcwidth(s))


def _box_table(rows, headers: list[str] | None = None) -> str:
    """把 rows 渲染成 Unicode 方框线表格（表头着色、中文宽度对齐）。headers 默认取上线汇总的列。"""
    headers = headers or [disp for disp, _ in _COLS]
    colw = {h: _wcwidth(h) for h in headers}
    for row in rows:
        for h in headers:
            colw[h] = max(colw[h], _wcwidth(row.get(h, "")))

    def line(left, mid, right):
        return left + mid.join("─" * (colw[h] + 2) for h in headers) + right

    def cells(vals):
        # 先用纯文本按显示宽度 pad 并拼好整行（对齐只基于纯文本，绝不把颜色码算进宽度）。
        parts = [" " + _pad(vals.get(h, ""), colw[h]) + " " for h in headers]
        return "│" + "│".join(parts) + "│"

    out = [line("┌", "┬", "┐")]
    # 表头整行拼好后再整体着色（一个 color 包裹整行，颜色码不参与任何对齐计算）。
    out.append(color(cells({h: h for h in headers}), "bold"))
    out.append(line("├", "┼", "┤"))
    for row in rows:
        out.append(cells(row))
    out.append(line("└", "┴", "┘"))
    return "\n".join(out)


# 下线汇总的列：与上线汇总相同，只是最后一列是『下线时间』。
OFFLINE_HEADERS = ["Rule ID", "升级类型", "机型信息", "机型扩展码", "源版本", "目标版本", "下线时间"]


def render_summary(device_code: str, rows: list[dict], kind: str = "online", when: str = "") -> str:
    """渲染一个设备类型的汇总块（带颜色的文本）。kind='online' 上线汇总，'offline' 下线汇总。

    下线汇总与上线同一版式，但用红色标题、⛔ 标记、『下线规则总数 / 下线时间』字样和红色提示框，
    一眼能和上线区分开。
    """
    lo, hi = _rule_id_range(rows)
    rng = lo if lo == hi else f"{lo} ~ {hi}"
    out: list[str] = []
    if kind == "offline":
        title = "⛔  OTA 下线信息汇总 / OTA Offline Summary  【已下线】"
        tone, headers = "red", OFFLINE_HEADERS
    else:
        title = "📦  OTA 部署信息汇总 / OTA Deployment Summary"
        tone, headers = "bold", None
    tw = _wcwidth(title)
    out.append(color("╔" + "═" * (tw + 2) + "╗", tone))
    out.append(color(f"║ {title} ║", tone))
    out.append(color("╚" + "═" * (tw + 2) + "╝", tone))
    if kind == "offline":
        out.append(f"  设备类型代码：{color(device_code or '(未知)', 'bold')}"
                   f"    ·    {color('下线规则总数', 'red')}：{color(str(len(rows)), 'bold')} 条"
                   f"    ·    Rule ID 范围：{color(rng, 'red')}"
                   f"    ·    下线时间：{color(when or '-', 'bold')}")
    else:
        out.append(f"  设备类型代码：{color(device_code or '(未知)', 'bold')}"
                   f"    ·    部署规则总数：{color(str(len(rows)), 'bold')} 条"
                   f"    ·    Rule ID 范围：{color(rng, 'green')}")
    out.append("")
    out.append(_box_table(rows, headers))
    out.append("")
    if kind == "offline":
        tip = f"⛔ 已下线：以上 {len(rows)} 条规则已从『已选规则』移出，设备将不再收到这些升级。Rule ID 范围：{rng}"
        tip_tone = "red"
    else:
        tip = f"⚠ 灰度提醒：进行下一阶段灰度发布时，请提供本次的 Rule ID 信息。本次部署 Rule ID 范围：{rng}"
        tip_tone = "yellow"
    bw = _wcwidth(tip)
    out.append(color("┌" + "─" * (bw + 2) + "┐", tip_tone))
    out.append(color(f"│ {tip} │", tip_tone))
    out.append(color("└" + "─" * (bw + 2) + "┘", tip_tone))
    return "\n".join(out)


def print_summary(xlsx_path: str) -> int:
    device_code, rows = _read_rows(xlsx_path)
    if not rows:
        print(err("未从该 Excel 读到任何含 Rule ID 的数据行。"), file=sys.stderr)
        return 1
    print(render_summary(device_code, rows, "online"))
    print(ok(f"汇总完成（complete）：共 {len(rows)} 条规则。"))
    return 0


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="deployota summary",
        description="从更新后的 Excel（含 Rule ID）打印 OTA 部署信息汇总。",
    )
    p.add_argument("excel", nargs="?", default=None, help="更新后的 Excel 路径（含 Rule ID）。")
    p.add_argument("--file", dest="file", default=None, help="同 excel 位置参数，二选一。")
    return p


def main() -> int:
    args = _build_parser().parse_args()
    path = args.file or args.excel
    if not path:
        print(err("请提供更新后的 Excel 路径：deployota summary <excel>（或 --file）。"), file=sys.stderr)
        return 2
    import os
    if not os.path.isfile(path):
        print(err(f"找不到 Excel 文件：{path}"), file=sys.stderr)
        return 2
    try:
        return print_summary(path)
    except Exception as e:
        print(err(f"汇总失败：{e}"), file=sys.stderr)
        return 1


# —— 不连网的纯逻辑自测 ——
def _selftest() -> int:
    import tempfile, os
    from openpyxl import Workbook
    ok_all = True

    def check(name, cond):
        nonlocal ok_all
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        if not cond:
            ok_all = False

    # 造一个更新后的 Excel（含 Rule ID / 内部机型信息 / 设备扩展信息 / 源版本 / 目标版本 / 升级日期）。
    wb = Workbook(); ws = wb.active; ws.title = "升级规则"
    ws.append(["设备类型代码", "内部机型信息", "设备扩展信息", "OTA类型", "升级方式",
               "升级日期", "Rule ID", "源版本", "目标版本"])
    ws.append(["MT9655AMAndrGTV", "55U75SUA(15)", "55U75SUA15524", "2", "强制升级",
               "20261105", "65345", "0100UQ0723", "Q0723_to_Q1028"])
    ws.append(["MT9655AMAndrGTV", "85U75SUA(15)", "85U75SUA15019", "2", "强制升级",
               "20261105", "65348", "0100UQ0723", "Q0723_to_Q1028"])
    p = os.path.join(tempfile.mkdtemp(), "upd.xlsx"); wb.save(p)

    dc, rows = _read_rows(p)
    check("设备类型代码读出", dc == "MT9655AMAndrGTV")
    check("读到 2 行", len(rows) == 2)
    check("Rule ID 范围 65345~65348", _rule_id_range(rows) == ("65345", "65348"))
    check("机型扩展码映射自 设备扩展信息", rows[0]["机型扩展码"] == "55U75SUA15524")
    check("部署时间映射自 升级日期", rows[0]["部署时间"] == "20261105")
    check("无『目标完整版本』列时目标版本退回版本变化值", rows[0]["目标版本"] == "Q0723_to_Q1028")
    # 有『目标完整版本』列 → 汇总显示完整版本；版本变化单独保留给部署库。
    wb2 = Workbook(); ws2 = wb2.active; ws2.title = "升级规则"
    ws2.append(["设备类型代码", "Rule ID", "源版本", "目标完整版本", "目标版本", "升级日期"])
    ws2.append(["DEV", "65370", "0100QQ0305", "0100UQ1104", "Q0305_to_Q1104", "20261109"])
    p2 = os.path.join(tempfile.mkdtemp(), "full.xlsx"); wb2.save(p2)
    _, rows2 = _read_rows(p2)
    check("目标版本显示『目标完整版本』", rows2[0]["目标版本"] == "0100UQ1104")
    check("版本变化另存（供部署库身份键）", rows2[0]["版本变化"] == "Q0305_to_Q1104")
    check("汇总表不显示『版本变化』列", "版本变化" not in render_summary("DEV", rows2))
    off = render_summary("DEV", [{"Rule ID": "65370", "目标版本": "0100UQ1104", "下线时间": "2026-11-10 15:30"}],
                         "offline", "2026-11-10 15:30")
    check("下线汇总有明显下线标识与下线时间",
          "下线信息汇总" in off and "已下线" in off and "下线时间" in off and "2026-11-10 15:30" in off)
    check("升级类型映射自 升级方式", rows[0]["升级类型"] == "强制升级")

    print("\n自测结果：", "全部通过 ✅" if ok_all else "存在失败 ❌")
    return 0 if ok_all else 1


if __name__ == "__main__":
    if "--selftest" in sys.argv[1:]:
        sys.exit(_selftest())
    sys.exit(main())
