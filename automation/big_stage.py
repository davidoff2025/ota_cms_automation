"""放 20K 时按『源版本用户数』决定：放 20K / 再放一轮 20K / 直接完全开放（纯逻辑）。

只在 ./deployota update ota20k 时用（update ota-all 不检查用户数）。整批一起判断：

  当前阶段     这批里有『源版本用户数』> 20000 的？   动作              内部阶段键
  OTA5000      有（至少一条）                         放到 20K          ota20k
  OTA5000      没有（全部 ≤ 20000）                   完全开放          ota5k-all
  OTA20000     有（至少一条）                         再放一轮 20K      ota20k-again   （CMS 累计 +20000，DASS 不动）
  OTA20000     没有（全部 ≤ 20000）                   完全开放          ota-all

『当前阶段』取自阶段记录（rule_stages）：整批都在 20K → 第二轮；都在 5K（或没有记录）→ 第一次；混着 → 停。
源版本用户数取自成功率闸门同一张报表、同一行。有 Rule ID 读不到用户数 → 停（可 override：只按读到的判断）。
"""
from __future__ import annotations

THRESHOLD = 20000

ACTION_TEXT = {
    "ota20k": "放到 20K（累计 6255 → 26255）",
    "ota20k-again": "再放一轮 20K（累计检测数 +20000；DASS 保持 OTA20000，不改）",
    "ota5k-all": "完全开放（5000阶段 → 全量；DASS 当前阶段 → 完全开放）",
    "ota-all": "完全开放（20000阶段 → 全量；DASS 当前阶段 → 完全开放）",
}


def parse_users(text) -> int | None:
    """『源版本用户数』单元格 → int（去千分位逗号 / 空白）；读不出 → None。"""
    t = "".join(str(text or "").replace(",", "").replace("，", "").split())
    return int(t) if t.isdigit() else None


def start_option(stages: dict[str, str]) -> tuple[str | None, str]:
    """{rid: 阶段记录里的当前阶段（'5K' / '20K' / '' …）} → (起始内部阶段键, 问题说明)。

    全部 20K → 'ota20k-again'；全部 5K 或没有记录 → 'ota20k'；混着（或在别的阶段）→ (None, 说明)。
    """
    at20 = sorted(r for r, s in stages.items() if s == "20K")
    other = sorted(r for r, s in stages.items() if s not in ("20K", "5K", ""))
    if other:
        return None, "这些 Rule ID 不在 5K / 20K：" + "、".join(f"{r}={stages[r]}" for r in other)
    if at20 and len(at20) != len(stages):
        rest = sorted(set(stages) - set(at20))
        return None, f"一批里混着 20K（{'、'.join(at20)}）和 5K（{'、'.join(rest)}），请分开跑"
    return ("ota20k-again" if at20 else "ota20k"), ""


def decide(option: str, users: dict[str, int | None], allow_missing: bool = False) -> dict:
    """option ∈ {'ota20k', 'ota20k-again'}；users = {rid: 源版本用户数 或 None}。

    返回 {"action": 内部阶段键 或 None(停), "big": [(rid, n) > 阈值], "missing": [rid 读不到], "reason": 说明}。
    allow_missing=True（你已 override）：读不到的不参与判断。
    """
    missing = sorted(r for r, n in users.items() if n is None)
    big = sorted(((r, n) for r, n in users.items() if n is not None and n > THRESHOLD), key=lambda x: -x[1])
    if missing and not allow_missing:
        return {"action": None, "big": big, "missing": missing,
                "reason": f"{len(missing)} 个 Rule ID 读不到『源版本用户数』：{'、'.join(missing)}"}
    known = [n for n in users.values() if n is not None]
    if not known:
        return {"action": None, "big": [], "missing": missing, "reason": "一个都读不到『源版本用户数』"}
    if big:
        act = "ota20k" if option == "ota20k" else "ota20k-again"
        why = f"{len(big)} 条 > {THRESHOLD:,}（最大 {big[0][0]}={big[0][1]:,}）"
    else:
        act = "ota5k-all" if option == "ota20k" else "ota-all"
        why = f"全部 ≤ {THRESHOLD:,}（最大 {max(known):,}）"
    return {"action": act, "big": big, "missing": missing, "reason": why}


def format_table(users: dict[str, int | None], rates: dict | None = None, notes: dict | None = None) -> list[str]:
    """逐条：Rule ID / 成功率 / 源版本用户数 / 是否 > 20000（/ 来源说明）。"""
    rates, notes = rates or {}, notes or {}
    out = ["  Rule ID   成功率      源版本用户数   > 20,000？"]
    for rid in sorted(users, key=lambda x: int(x) if x.isdigit() else 0):
        n = users[rid]
        rate = (rates.get(rid) or {}).get("text", "") or "-"
        flag = "读不到" if n is None else ("是" if n > THRESHOLD else "否")
        note = notes.get(rid, "")
        out.append(f"  {rid:<9} {rate:<11} {('-' if n is None else f'{n:,}'):>12}   {flag}"
                   + (f"   （{note}）" if note else ""))
    return out


def _selftest() -> int:
    good = True

    def check(name, cond):
        nonlocal good
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        good = good and bool(cond)

    check("用户数解析", parse_users("92,108") == 92108 and parse_users(" 6229 ") == 6229
          and parse_users("") is None and parse_users("-") is None)
    check("起始阶段：全 5K / 没记录 → ota20k", start_option({"1": "5K", "2": ""}) == ("ota20k", ""))
    check("起始阶段：全 20K → ota20k-again", start_option({"1": "20K", "2": "20K"}) == ("ota20k-again", ""))
    check("起始阶段：混着 → 停", start_option({"1": "20K", "2": "5K"})[0] is None
          and start_option({"1": "1K"})[0] is None)
    d = decide("ota20k", {"64350": 92108, "64351": 12940})
    check("5K + 有一条 > 20000 → 20K（整批）", d["action"] == "ota20k" and d["big"] == [("64350", 92108)])
    check("5K + 全部 ≤ 20000 → 完全开放", decide("ota20k", {"1": 6229, "2": 20000})["action"] == "ota5k-all")
    check("20K + 有一条 > 20000 → 再一轮 20K", decide("ota20k-again", {"1": 48555, "2": 151})["action"] == "ota20k-again")
    check("20K + 全部 ≤ 20000 → 完全开放", decide("ota20k-again", {"1": 309})["action"] == "ota-all")
    check("正好 20000 算低于", decide("ota20k", {"1": 20000})["action"] == "ota5k-all")
    m = decide("ota20k", {"1": 30000, "2": None})
    check("读不到用户数 → 停", m["action"] is None and m["missing"] == ["2"])
    check("override：只按读到的判断", decide("ota20k", {"1": 3000, "2": None}, allow_missing=True)["action"] == "ota5k-all")
    check("一个都读不到 → 停（override 也停）", decide("ota20k", {"1": None}, allow_missing=True)["action"] is None)
    t = format_table({"64350": 92108, "64351": None}, {"64350": {"text": "29.92%"}})
    check("逐条表格", "92,108" in t[1] and "29.92%" in t[1] and "读不到" in t[2])
    print("\n自测结果：", "全部通过 ✅" if good else "存在失败 ❌")
    return 0 if good else 1


if __name__ == "__main__":
    import sys
    sys.exit(_selftest() if "--selftest" in sys.argv[1:] else 0)
