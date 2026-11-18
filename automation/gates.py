"""update 的间隔闸门（纯逻辑）：距离进入【上一阶段】必须超过 48 小时。

  - 时间来源 1：本地阶段记录（rule_stages）进入上一阶段的时间（精确到秒）：now − 那个时间 > 48 小时。
    20K 再放一轮：取『进入 20K』和最近一条『20K 再放一轮』CMS 备注里较晚的那个。
  - 时间来源 2（本地没有记录时）：报表里的『当前阶段开放时间』（只有日期）：今天 ≥ 开放日期 + 2 天就算满足。
  - 两个都没有 → 不通过（不知道开放了多久）。
  - ota50（场测 → OTA50）不检查。
"""
from __future__ import annotations

import re
from datetime import datetime, timedelta

MIN_HOURS = 48

# 内部阶段键 → 上一阶段（本地阶段记录里的键）
PREV_STAGE_KEY = {"ota200": "50", "ota1k": "200", "ota5k": "1K", "ota20k": "5K", "ota5k-all": "5K",
                  "ota20k-again": "20K", "ota-all": "20K"}


def _parse_ts(t) -> datetime | None:
    try:
        return datetime.strptime(str(t)[:19], "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return None


def _parse_day(t) -> datetime | None:
    d = re.sub(r"\D", "", str(t or ""))[:8]
    try:
        return datetime.strptime(d, "%Y%m%d") if len(d) == 8 else None
    except ValueError:
        return None


def entered_at(rid: str, option: str, state: dict, events: list[dict]) -> datetime | None:
    """本地记录里这条进入上一阶段的时间；没有返回 None。"""
    key = PREV_STAGE_KEY.get(option)
    s = state.get(rid) or {}
    t = _parse_ts((s.get("阶段") or {}).get(key, "")) if key else None
    if key == "20K":
        from .stage_db import AGAIN_NOTE
        notes = [_parse_ts(e.get("时间")) for e in events
                 if str(e.get("Rule ID", "")) == rid and e.get("类型") == "CMS"
                 and str(e.get("备注", "")).startswith(AGAIN_NOTE)]
        cands = [x for x in [t] + notes if x]
        t = max(cands) if cands else None
    return t


def check_interval(ids: list[str], option: str, state: dict, events: list[dict], report_open: dict,
                   now: datetime | None = None) -> list[dict]:
    """返回每条：{rid, ok, since(显示用), hours(已过小时，可能 None), need(还差小时), source}。option=ota50 → 全部 ok。"""
    now = now or datetime.now()
    out = []
    for rid in ids:
        if option not in PREV_STAGE_KEY:
            out.append({"rid": rid, "ok": True, "since": "-", "hours": None, "need": 0, "source": "不检查"})
            continue
        t = entered_at(rid, option, state, events)
        if t:
            h = (now - t).total_seconds() / 3600
            out.append({"rid": rid, "ok": h > MIN_HOURS, "since": t.strftime("%Y-%m-%d %H:%M"), "hours": h,
                        "need": max(0.0, MIN_HOURS - h), "source": "阶段记录"})
            continue
        d = _parse_day(report_open.get(rid, ""))
        if d:
            ok = now.date() >= (d + timedelta(days=2)).date()
            h = (now - d).total_seconds() / 3600
            out.append({"rid": rid, "ok": ok, "since": d.strftime("%Y-%m-%d"), "hours": h,
                        "need": 0 if ok else max(0.0, (d + timedelta(days=2) - now).total_seconds() / 3600),
                        "source": "报表开放日期（+2 天）"})
            continue
        out.append({"rid": rid, "ok": False, "since": "?", "hours": None, "need": None, "source": "查不到开放时间"})
    return out


def format_interval(rows: list[dict], stage_label: str) -> list[str]:
    lines = [f"\n======== 间隔闸门：距进入 {stage_label} 必须 > {MIN_HOURS} 小时 ========",
             f"  {'Rule ID':8} {'进入时间':18} {'已过':>9}   结果（来源）"]
    for r in rows:
        past = "-" if r["hours"] is None else f"{r['hours']:.1f} h"
        if r["ok"]:
            res = "通过 ✅"
        elif r["need"] is None:
            res = "不通过 ❌ 不知道开放了多久"
        else:
            res = f"不通过 ❌ 还差 {r['need']:.1f} h"
        lines.append(f"  {r['rid']:8} {r['since']:18} {past:>9}   {res}（{r['source']}）")
    bad = [r["rid"] for r in rows if not r["ok"]]
    lines.append(f"  结论：{'全部通过' if not bad else f'{len(bad)} 条不满足：' + '、'.join(bad)}")
    return lines


def _selftest() -> int:
    good = True

    def check(name, cond):
        nonlocal good
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        good = good and bool(cond)

    now = datetime(2026, 11, 18, 10, 0, 0)
    state = {"a": {"阶段": {"1K": "2026-11-16 09:00:00"}}, "b": {"阶段": {"1K": "2026-11-16 11:00:00"}},
             "c": {"阶段": {"20K": "2026-11-10 09:00:00"}}}
    ev = [{"Rule ID": "c", "类型": "CMS", "备注": "20K 再放一轮：c 26255→46255", "时间": "2026-11-17 09:00:00"}]
    r = {x["rid"]: x for x in check_interval(["a", "b"], "ota5k", state, ev, {}, now)}
    check("阶段记录：49h 通过、47h 不通过（还差 1h）",
          r["a"]["ok"] and not r["b"]["ok"] and abs(r["b"]["need"] - 1.0) < 0.01)
    r2 = check_interval(["c"], "ota20k-again", state, ev, {}, now)[0]
    check("20K 再一轮：按最近一轮的时间算", not r2["ok"] and r2["since"] == "2026-11-17 09:00")
    check("ota-all 也按 20K 最近一轮", not check_interval(["c"], "ota-all", state, ev, {}, now)[0]["ok"])
    r3 = {x["rid"]: x for x in check_interval(["x", "y", "z"], "ota200", {}, [],
                                              {"x": "20261116", "y": "2026-11-17"}, now)}
    check("报表日期：11/16 开放 → 11/18 通过；11/17 → 不通过；没有 → 不通过",
          r3["x"]["ok"] and not r3["y"]["ok"] and not r3["z"]["ok"] and r3["z"]["need"] is None)
    check("ota50 不检查", all(x["ok"] for x in check_interval(["q"], "ota50", {}, [], {}, now)))
    lines = format_interval(check_interval(["a", "b"], "ota5k", state, ev, {}, now), "OTA1000")
    check("表格", any("还差 1.0 h" in l for l in lines) and "1 条不满足：b" in lines[-1])
    print("\n自测结果：", "全部通过 ✅" if good else "存在失败 ❌")
    return 0 if good else 1


if __name__ == "__main__":
    import sys
    sys.exit(_selftest() if "--selftest" in sys.argv[1:] else 0)
