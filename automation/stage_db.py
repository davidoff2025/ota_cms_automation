"""Rule ID 灰度阶段记录（每个 Rule ID 一条时间线）。

存储：<OTADATA>/rule_stages.jsonl（只追加的事件日志）+ rule_stages.csv（每个 Rule ID 一行，供 Excel 查看）。

两类信息分开记：
  · 灰度阶段：场测 → 50 → 200 → 1K → 5K → 20K → 完全开放，每个阶段一个执行时间。
      - `dass init` 填写完成 → 场测；
      - `update otaXX`：CMS 灰度放开成功（回读累计检测数 = 目标值）时记；--no-cms / `dass otaXX` 时按 DASS 更新成功记。
      - 只前进不回退：已在更后阶段的不写；同一阶段只保留第一次的时间（重跑不会改日期）；跳档照写并标注。
  · CMS 在线状态：在线 / 已下线（+时间）。`online` / `publish` 上线成功记『在线』，`offline` 下线成功记『已下线』。
      下线只改在线状态，不动阶段时间线（offline 不改 DASS，也不该改阶段日期）。

用法（查看）：
    ./deployota query 65370            # Rule ID 查询结果里带当前阶段与时间线
    ./.venv/bin/python -m automation.stage_db --export    # 重新导出 CSV
"""
from __future__ import annotations

import csv
import json
import os
import sys
from datetime import datetime

from . import depdb

STAGES = ["场测", "50", "200", "1K", "5K", "20K", "完全开放"]
OPTION_STAGE = {"init": "场测", "ota50": "50", "ota200": "200", "ota1k": "1K",
                "ota5k": "5K", "ota20k": "20K", "ota-all": "完全开放"}
JSONL_NAME = "rule_stages.jsonl"
CSV_NAME = "rule_stages.csv"
CSV_FIELDS = ["Rule ID", "设备类型代码", "当前阶段"] + STAGES + ["CMS状态", "CMS状态时间"]


def _path(name: str) -> str:
    return os.path.join(depdb.db_dir(), name)


def load_events() -> list[dict]:
    p = _path(JSONL_NAME)
    if not os.path.isfile(p):
        return []
    out = []
    with open(p, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                try:
                    out.append(json.loads(line))
                except ValueError:
                    continue
    return out


def build_state(events: list[dict]) -> dict:
    """事件 → {rid: {"设备类型代码", "阶段": {阶段: 时间}, "当前阶段", "CMS状态", "CMS状态时间"}}（纯逻辑）。"""
    st: dict[str, dict] = {}
    for e in events:
        rid = str(e.get("Rule ID", "")).strip()
        if not rid:
            continue
        s = st.setdefault(rid, {"设备类型代码": "", "阶段": {}, "当前阶段": "", "CMS状态": "", "CMS状态时间": ""})
        if e.get("设备类型代码"):
            s["设备类型代码"] = e["设备类型代码"]
        if e.get("类型") == "阶段":
            stage = e.get("阶段", "")
            if stage in STAGES and stage not in s["阶段"]:
                s["阶段"][stage] = e.get("时间", "")
        elif e.get("类型") == "CMS":
            s["CMS状态"], s["CMS状态时间"] = e.get("状态", ""), e.get("时间", "")
    for s in st.values():
        reached = [x for x in STAGES if x in s["阶段"]]
        s["当前阶段"] = reached[-1] if reached else ""
    return st


def plan_stage(state: dict, ids: list[str], stage: str) -> tuple[list[str], list[tuple[str, str]], list[str]]:
    """决定哪些要写（纯逻辑）。返回 (要写, [(跳过的rid, 原因)], 跳档的rid)。"""
    idx = STAGES.index(stage)
    write, skip, jump = [], [], []
    for rid in ids:
        s = state.get(rid)
        if s and stage in s["阶段"]:
            skip.append((rid, f"已有{stage}记录（{s['阶段'][stage]}）"))
            continue
        cur = s["当前阶段"] if s else ""
        if cur and STAGES.index(cur) > idx:
            skip.append((rid, f"已在{cur}阶段，不回退"))
            continue
        if cur and idx > STAGES.index(cur) + 1:
            jump.append(rid)
        write.append(rid)
    return write, skip, jump


def _append(events: list[dict]) -> None:
    if not events:
        return
    os.makedirs(depdb.db_dir(), exist_ok=True)
    with open(_path(JSONL_NAME), "a", encoding="utf-8") as f:
        for e in events:
            f.write(json.dumps(e, ensure_ascii=False) + "\n")
    export_csv()


def _device_of(ids: list[str], code_of: dict | None, state: dict) -> dict:
    out = dict(code_of or {})
    missing = [r for r in ids if not out.get(r)]
    for r in list(missing):
        if state.get(r, {}).get("设备类型代码"):
            out[r] = state[r]["设备类型代码"]
    missing = [r for r in ids if not out.get(r)]
    if missing:
        from .active_offline import device_codes_for
        found, _ = device_codes_for(missing, depdb.load_all())
        out.update(found)
    return out


def record_stage(ids: list[str], option: str, source: str, code_of: dict | None = None,
                 when: str | None = None, quiet: bool = False) -> list[str]:
    """记录 ids 到达 option 对应的阶段。返回实际写入的 rid。写入失败只提示，不影响主流程。"""
    try:
        stage = OPTION_STAGE[option]
        ids = list(dict.fromkeys(str(x).strip() for x in ids if str(x).strip()))
        state = build_state(load_events())
        write, skip, jump = plan_stage(state, ids, stage)
        t = when or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        dev = _device_of(write, code_of, state)
        _append([{"Rule ID": r, "设备类型代码": dev.get(r, ""), "类型": "阶段", "阶段": stage,
                  "时间": t, "来源": source, "跳档": r in jump} for r in write])
        if not quiet:
            from ._log import color
            if write:
                print(color(f"✓ 阶段记录：{len(write)} 个 Rule ID → {stage}（{t}）", "green"))
            if jump:
                print(color(f"  ⚠ 跳档记录：{', '.join(jump)}（前一阶段没有记录）", "yellow"))
            if skip:
                print(color("  · 阶段记录未写：" + "；".join(f"{r} {why}" for r, why in skip), "dim"))
        return write
    except Exception as e:
        print(f"（阶段记录写入跳过：{e}）")
        return []


def record_cms(ids: list[str], status: str, source: str, code_of: dict | None = None,
               when: str | None = None) -> list[str]:
    """记录 CMS 在线状态（『在线』/『已下线』）。状态没变的不重复写。不影响阶段时间线。"""
    try:
        ids = list(dict.fromkeys(str(x).strip() for x in ids if str(x).strip()))
        state = build_state(load_events())
        write = [r for r in ids if state.get(r, {}).get("CMS状态") != status]
        t = when or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        dev = _device_of(write, code_of, state)
        _append([{"Rule ID": r, "设备类型代码": dev.get(r, ""), "类型": "CMS", "状态": status,
                  "时间": t, "来源": source} for r in write])
        return write
    except Exception as e:
        print(f"（CMS 状态记录写入跳过：{e}）")
        return []


def export_csv() -> str:
    state = build_state(load_events())
    p = _path(CSV_NAME)
    with open(p, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        w.writeheader()
        for rid in sorted(state, key=lambda x: int(x) if x.isdigit() else 0):
            s = state[rid]
            row = {"Rule ID": rid, "设备类型代码": s["设备类型代码"], "当前阶段": s["当前阶段"],
                   "CMS状态": s["CMS状态"], "CMS状态时间": s["CMS状态时间"]}
            row.update({k: s["阶段"].get(k, "") for k in STAGES})
            w.writerow(row)
    return p


def timeline(s: dict | None) -> str:
    """'场测 2026-11-10 → 50 2026-11-12'；没有记录返回空串。"""
    if not s:
        return ""
    return " → ".join(f"{k} {s['阶段'][k][:10]}" for k in STAGES if k in s["阶段"])


def _selftest() -> int:
    import tempfile
    good = True

    def check(name, cond):
        nonlocal good
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        good = good and bool(cond)

    old = os.environ.get("OTA_DEPLOY_DB_DIR")
    os.environ["OTA_DEPLOY_DB_DIR"] = tempfile.mkdtemp(prefix="stage_selftest_")
    try:
        code = {"1": "DEV", "2": "DEV", "3": "DEV"}
        check("init → 场测", record_stage(["1", "2"], "init", "dass init", code, "2026-11-10 10:00:00", True) == ["1", "2"])
        check("重跑 init 不重复", record_stage(["1"], "init", "dass init", code, "2026-11-11 10:00:00", True) == [])
        check("ota50 前进", record_stage(["1"], "ota50", "update", code, "2026-11-12 10:00:00", True) == ["1"])
        check("不回退：已到 50 不再写场测", record_stage(["1"], "init", "x", code, quiet=True) == [])
        w = record_stage(["2", "3"], "ota1k", "update", code, "2026-11-13 10:00:00", True)
        check("跳档照写（含无记录的）", w == ["2", "3"])
        check("ota200 不回退（2 已在 1K）", record_stage(["2"], "ota200", "update", code, quiet=True) == [])
        st = build_state(load_events())
        check("当前阶段", st["1"]["当前阶段"] == "50" and st["2"]["当前阶段"] == "1K")
        check("同一阶段保留第一次时间", st["1"]["阶段"]["场测"] == "2026-11-10 10:00:00")
        record_cms(["1"], "已下线", "offline", code, "2026-11-14 10:00:00")
        st = build_state(load_events())
        check("下线只改 CMS 状态、不动阶段", st["1"]["CMS状态"] == "已下线" and st["1"]["当前阶段"] == "50")
        check("CMS 状态不变不重复写", record_cms(["1"], "已下线", "offline", code) == [])
        check("时间线", timeline(st["1"]) == "场测 2026-11-10 → 50 2026-11-12")
        rows = list(csv.DictReader(open(_path(CSV_NAME), encoding="utf-8-sig")))
        check("CSV 每个 Rule ID 一行带各阶段日期", len(rows) == 3 and rows[0]["50"].startswith("2026-11-12"))
    finally:
        if old is None:
            os.environ.pop("OTA_DEPLOY_DB_DIR", None)
        else:
            os.environ["OTA_DEPLOY_DB_DIR"] = old
    print("\n自测结果：", "全部通过 ✅" if good else "存在失败 ❌")
    return 0 if good else 1


if __name__ == "__main__":
    if "--selftest" in sys.argv[1:]:
        sys.exit(_selftest())
    if "--export" in sys.argv[1:]:
        print(export_csv())
        sys.exit(0)
    print(__doc__)
