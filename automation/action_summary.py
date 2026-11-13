"""命令结束时的『操作总结』：逐步列出做了什么、结果如何，最后给一句结论。

用法（命令入口）：
    log = action_summary.start("update ota50", ["65345", ...])
    try:
        ...  # 各步骤里调用 action_summary.step(名称, 状态, 说明)
    finally:
        action_summary.finish()

状态：ok（完成）/ warn（部分完成或需要你注意）/ fail（失败或中止）/ skip（本次不需要或没执行）。
各模块调用 step() 时若当前没有在记录（例如被别的命令间接调用），什么都不做。
"""
from __future__ import annotations

from ._log import color

OK, WARN, FAIL, SKIP = "ok", "warn", "fail", "skip"

# 同意语句里的阶段写法
APPROVAL_STAGE = {"ota50": "OTA50", "ota200": "OTA200", "ota1k": "OTA1000", "ota5k": "OTA5000",
                  "ota20k": "OTA20000", "ota-all": "OTA完全放开"}


def short_rate(text: str) -> str:
    """'83.00%' → '83%'，'70.50%' → '70.5%'，'49.61%' 不变（只去末尾的 0，不四舍五入）。"""
    t = str(text or "").strip().replace(" ", "")
    if not t:
        return ""
    num = t.rstrip("%")
    if "." in num:
        num = num.rstrip("0").rstrip(".")
    return num + "%"


def approval_sentence(option: str, ids: list[str], day: str, rates: dict | None = None) -> str:
    """同意（OTA200，20261113，50032~50033）50032:83%，50033:70%。

    rates 给了就在后面逐个跟上成功率（rid:成功率，按 Rule ID 升序；查不到写『未查到』）。
    """
    from .rule_info import compress_ids
    ids = sorted(dict.fromkeys(ids), key=lambda x: int(x) if x.isdigit() else 0)
    head = f"同意（{APPROVAL_STAGE[option]}，{day}，{compress_ids(ids)}）"
    if rates is None:
        return head
    return head + "，".join(f"{rid}:{short_rate((rates.get(rid) or {}).get('text', '')) or '未查到'}" for rid in ids)
_ICON = {OK: ("✓", "green"), WARN: ("⚠", "yellow"), FAIL: ("✗", "red"), SKIP: ("·", "dim")}


class ActionLog:
    def __init__(self, title: str, ids: list[str] | None = None):
        self.title = title
        self.ids = list(ids or [])
        self.steps: list[dict] = []
        self.option = ""                 # update 的档位（ota50…），有值才生成同意语句
        self.rates: dict[str, dict] = {}  # {rid: {"text": '49.09%', "src": 'OTA50 阶段'}}
        self.approved: list[str] = []    # 真正升级成功（CMS 放开 + DASS 写入）的 Rule ID
        self.submitted: bool | None = None  # DASS 是否已提交

    def set(self, name: str, status: str, detail: str = "") -> None:
        """同名步骤只保留最后一次结果（后面的判断会覆盖前面的）。"""
        for s in self.steps:
            if s["name"] == name:
                s.update(status=status, detail=detail)
                return
        self.steps.append({"name": name, "status": status, "detail": detail})

    def verdict(self) -> tuple[str, str]:
        st = [s["status"] for s in self.steps]
        if not st:
            return FAIL, "没有执行任何步骤"
        if FAIL in st:
            first = next(s for s in self.steps if s["status"] == FAIL)
            return FAIL, f"未完成：卡在『{first['name']}』—— {first['detail']}"
        if WARN in st:
            return WARN, "部分完成，请看上面标 ⚠ 的步骤"
        return OK, "全部完成"

    def render(self) -> str:
        from .rule_info import compress_ids
        lines = [color(f"\n======== 操作总结：{self.title} ========", "bold")]
        if self.ids:
            lines.append(f"  Rule ID：{compress_ids(self.ids)}（{len(self.ids)} 个）")
        for i, s in enumerate(self.steps, 1):
            icon, tone = _ICON.get(s["status"], ("?", "dim"))
            text = f"  {i}. {icon} {s['name']}" + (f"：{s['detail']}" if s["detail"] else "")
            lines.append(color(text, tone))
        v, msg = self.verdict()
        icon, tone = _ICON[v]
        lines.append(color(f"  结论：{icon} {msg}", tone))
        if self.option and self.approved:
            from datetime import datetime
            sentence = approval_sentence(self.option, self.approved, datetime.now().strftime("%Y%m%d"), self.rates)
            lines.append("")
            lines.append(color(f"  {sentence}", "green" if self.submitted else "yellow"))
            if any(not (self.rates.get(r) or {}).get("text") for r in self.approved):
                lines.append(color("  （『未查到』= 安卓升级展示报表里暂无这条的数据）", "dim"))
            if any("参考" in (self.rates.get(r) or {}).get("src", "") for r in self.approved):
                lines.append(color("  （部分成功率取自报表当前阶段，仅供参考，见上方闸门明细）", "dim"))
            if self.submitted is not True:
                lines.append(color("  ⚠ DASS 尚未提交，提交完成后再发这句。", "yellow"))
        return "\n".join(lines)


CURRENT: ActionLog | None = None


def start(title: str, ids: list[str] | None = None) -> ActionLog:
    global CURRENT
    CURRENT = ActionLog(title, ids)
    return CURRENT


def step(name: str, status: str, detail: str = "") -> None:
    if CURRENT is not None:
        CURRENT.set(name, status, detail)


def set_option(option: str) -> None:
    if CURRENT is not None and option in APPROVAL_STAGE:
        CURRENT.option = option


def set_rates(rates: dict) -> None:
    """rates: {rid: {"text": '49.09%', "src": '...'}}；已有的不覆盖为空。"""
    if CURRENT is not None:
        for rid, r in rates.items():
            if r.get("text") or rid not in CURRENT.rates:
                CURRENT.rates[rid] = r


def set_approved(ids: list[str]) -> None:
    if CURRENT is not None:
        CURRENT.approved = list(dict.fromkeys(ids))


def set_submitted(ok: bool) -> None:
    if CURRENT is not None:
        CURRENT.submitted = ok


def finish() -> None:
    """打印总结并结束记录（重复调用无副作用）。"""
    global CURRENT
    if CURRENT is not None:
        print(CURRENT.render())
        CURRENT = None


def _selftest() -> int:
    import re
    good = True

    def check(name, cond):
        nonlocal good
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        good = good and bool(cond)

    plain = lambda t: re.sub(r"\x1b\[[0-9;]*m", "", t)
    log = ActionLog("update ota50", ["65345", "65346"])
    log.set("CMS 灰度放开到 50阶段", OK, "成功 2 个")
    log.set("DASS 更新到 OTA50", OK, "写入 2 个")
    log.set("DASS 提交", OK, "已提交")
    check("全部完成", log.verdict()[0] == OK and "全部完成" in plain(log.render()))
    log.set("DASS 提交", WARN, "你没有输入 approve，未提交")
    check("同名步骤覆盖、部分完成", len(log.steps) == 3 and log.verdict()[0] == WARN)
    log.set("CMS 灰度放开到 50阶段", FAIL, "累计检测数不符")
    check("失败时结论点出卡在哪一步", "卡在『CMS 灰度放开到 50阶段』" in log.verdict()[1])
    check("渲染含 Rule ID 范围和序号", "65345~65346" in plain(log.render()) and "1. ✗" in plain(log.render()))
    check("没有记录时 step() 不报错", step("x", OK) is None)
    check("同意语句格式", approval_sentence("ota50", ["65348", "65345", "65346", "65347"], "20261113")
          == "同意（OTA50，20261113，65345~65348）")
    check("各档写法", [APPROVAL_STAGE[k] for k in ("ota1k", "ota5k", "ota20k", "ota-all")]
          == ["OTA1000", "OTA5000", "OTA20000", "OTA完全放开"])
    log2 = ActionLog("update ota200", ["65345", "65346"])
    log2.option, log2.approved, log2.submitted = "ota200", ["65345", "65346"], True
    log2.rates = {"65345": {"text": "61.20%", "src": "OTA50 阶段"}}
    out = plain(log2.render())
    check("同意语句后跟每个 Rule ID 的成功率", "65345~65346）65345:61.2%，65346:未查到" in out
          and "同意（OTA200，" in out and "尚未提交" not in out)
    check("成功率写法", (short_rate("83.00%"), short_rate("70.50%"), short_rate("49.61%"), short_rate(""))
          == ("83%", "70.5%", "49.61%", ""))
    check("示例格式", approval_sentence("ota200", ["50033", "50032"], "20261113",
                                    {"50032": {"text": "83.00%"}, "50033": {"text": "70.00%"}})
          == "同意（OTA200，20261113，50032~50033）50032:83%，50033:70%")
    log2.submitted = None
    check("未提交时提示", "尚未提交" in plain(log2.render()))
    print("\n自测结果：", "全部通过 ✅" if good else "存在失败 ❌")
    return 0 if good else 1


if __name__ == "__main__":
    import sys
    sys.exit(_selftest() if "--selftest" in sys.argv[1:] else 0)
