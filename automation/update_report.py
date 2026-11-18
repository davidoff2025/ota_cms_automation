"""deployota update 结束后的放量汇总：收集结果 → 记录（OTADATA/update_runs.jsonl）→ HTML（带灰度进度条）→ 邮件。

邮件：To = 数据统计报告（OTA_MAIL_CC，默认 sys.datareports@hisense.com）+ --mail-to 给的人；Bcc = devinw（OTA_MAIL_BCC）。
什么时候发：有规则在 CMS 真正放开 → 发（失败的条目也写进去）；执行了但一条都没放开 → 发失败通知；
          --dry-run / 你取消 / 什么都没改（本来就在目标阶段）→ 不发。--no-mail 不发（照样记录）。
"""
from __future__ import annotations

import html as _html
import os
from datetime import datetime

from ._log import color

# 内部阶段键 → (上一阶段, 本次阶段)，用来写『OTA5000 → OTA20000』
TRANSITION = {
    "ota50": ("场测", "OTA50"), "ota200": ("OTA50", "OTA200"), "ota1k": ("OTA200", "OTA1000"),
    "ota5k": ("OTA1000", "OTA5000"), "ota20k": ("OTA5000", "OTA20000"), "ota5k-all": ("OTA5000", "完全开放"),
    "ota20k-again": ("OTA20000", "OTA20000（再一轮）"), "ota-all": ("OTA20000", "完全开放"),
}
STATUS_TEXT = {"progress": "放量完成", "partial": "部分完成", "failed": "未放开", "none": "无变化"}


def _h(x) -> str:
    return _html.escape(str(x if x is not None else ""))


# ----------------------------------------------------------------------
# 收集（在 action_summary.finish() 之前调用）
# ----------------------------------------------------------------------

def collect(requested: str, ids: list[str], dry_run: bool = False, no_cms: bool = False) -> dict:
    from . import action_summary as A
    from . import cms_grey, depdb
    from .dass_perf import LAST_USERS
    from .dass_update import resolve_rules
    cur = A.CURRENT
    run = dict(cms_grey.LAST_RUN) if not dry_run else {}
    final = run.get("option") or (cur.option if cur else "") or requested
    done = list(run.get("done", [])) if run else (list(cur.approved) if (cur and no_cms) else [])
    already = list(run.get("already", []))
    failed = [tuple(x) for x in run.get("failed", [])]
    mismatch = [tuple(x) for x in run.get("mismatch", [])]
    before = dict(run.get("before", {}))
    steps = list(cur.steps) if cur else []
    cancelled = any(k in str(s.get("detail", "")) for s in steps for k in ("你取消了", "你没有确认", "无法确认"))
    any_fail = any(s.get("status") == "fail" for s in steps)
    if dry_run:
        status = "none"
    elif done:
        status = "partial" if (failed or mismatch or any_fail) else "progress"
    elif cancelled or (not failed and not any_fail):
        status = "none"
    else:
        status = "failed"
    recs = {r["ruleid"]: r for r, _ in resolve_rules(ids, depdb.load_all())[0]}
    from .cms_grey import CMS_STAGE
    st = CMS_STAGE.get(final, {})
    rows = []
    fail_why = dict(failed)
    mis_why = dict(mismatch)
    for rid in ids:
        rec = recs.get(rid, {})
        if rid in done:
            res = "已放开"
            if st.get("total") is None and st:
                cnt = (before.get(rid) or 0) + st.get("step", 0)
            else:
                cnt = st.get("total", "")
        elif rid in already:
            res, cnt = "本来已是目标阶段", st.get("total", "")
        elif rid in fail_why:
            res, cnt = f"失败：{fail_why[rid]}", before.get(rid, "")
        elif rid in mis_why:
            res, cnt = f"不在上一阶段：{mis_why[rid]}", ""
        else:
            res, cnt = "未处理", ""
        rate = ((cur.rates.get(rid) if cur else None) or {}).get("text", "")
        rows.append({"ruleid": rid, "device": rec.get("方案号", ""), "model": rec.get("机型", ""),
                     "ext": rec.get("设备扩展信息", ""), "src": rec.get("源版本", ""), "tgt": rec.get("目标版本", ""),
                     "rate": rate, "users": LAST_USERS.get(rid), "result": res,
                     "count": "全量" if cnt == -1 else cnt, "ok": rid in done})
    sentence = ""
    if cur and cur.option and cur.approved:
        sentence = A.approval_sentence(cur.option, cur.approved, datetime.now().strftime("%Y%m%d"), cur.rates)
    frm, to = TRANSITION.get(final, ("", final))
    return {"time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "requested": requested, "action": final,
            "from": frm, "to": to, "ids": list(ids), "done": done, "already": already, "failed": failed,
            "mismatch": mismatch, "rows": rows, "steps": steps, "status": status,
            "dass_submitted": (cur.submitted if cur else None), "sentence": sentence,
            "verdict": (cur.verdict()[1] if cur else ""), "dry_run": dry_run}


# ----------------------------------------------------------------------
# HTML
# ----------------------------------------------------------------------

FONT = "font-family:-apple-system,'Segoe UI','PingFang SC','Microsoft YaHei',Arial,sans-serif;"


def _bar_html(cells: list[dict]) -> str:
    """一条进度条（表格画的：Outlook / 手机都正常）。"""
    tds = []
    for i, c in enumerate(cells):
        if i:
            lit = cells[i - 1]["state"] not in ("todo",) and c["state"] not in ("todo",)
            tds.append(f'<td style="padding:0;" valign="top"><div style="height:3px;margin-top:15px;'
                       f'background:{"#10b981" if lit else "#e5e7eb"};min-width:18px;"></div></td>')
        if c["state"] == "done":
            dot = ('<div style="width:30px;height:30px;line-height:30px;border-radius:15px;background:#10b981;'
                   'color:#fff;font-size:15px;font-weight:700;margin:0 auto;">✓</div>')
            tone = "#047857"
        elif c["state"] == "skipped":
            dot = ('<div style="width:26px;height:26px;line-height:26px;border-radius:15px;background:#f3f4f6;'
                   'border:2px dashed #d1d5db;color:#9ca3af;font-size:14px;margin:0 auto;">–</div>')
            tone = "#9ca3af"
        elif c["state"] == "current":
            dot = ('<div style="width:26px;height:26px;line-height:26px;border-radius:15px;background:#2563eb;'
                   'border:2px solid #bfdbfe;color:#fff;font-size:13px;font-weight:700;margin:0 auto;">●</div>')
            tone = "#1d4ed8"
        else:
            dot = ('<div style="width:26px;height:26px;line-height:26px;border-radius:15px;background:#fff;'
                   'border:2px solid #d1d5db;color:#9ca3af;font-size:13px;margin:0 auto;">○</div>')
            tone = "#9ca3af"
        label = c["label"] + (f" ×{c['rounds']}" if c.get("rounds", 0) > 1 else "")
        if c["state"] == "skipped":
            c = dict(c, date="跳过")
        weight = "700" if c["state"] == "current" else "500"
        tds.append(f'<td align="center" valign="top" style="padding:0 2px;min-width:58px;">{dot}'
                   f'<div style="{FONT}font-size:12px;font-weight:{weight};color:{tone};margin-top:6px;'
                   f'white-space:nowrap;">{_h(label)}</div>'
                   f'<div style="{FONT}font-size:11px;color:#6b7280;margin-top:2px;">{_h(c["date"]) or "&nbsp;"}</div></td>')
    return ('<table role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;">'
            f'<tr>{"".join(tds)}</tr></table>')


def progress_groups(ids: list[str]) -> list[tuple[list[str], list[dict]]]:
    """按进度相同分组：[(Rule IDs, 进度格)]（同一批阶段都一样时只有一组）。"""
    from .stage_db import build_state, load_events, progress_cells, rounds_20k
    ev = load_events()
    state = build_state(ev)
    groups: dict[tuple, tuple[list[str], list[dict]]] = {}
    for rid in ids:
        s = state.get(rid)
        cells = progress_cells(s, rounds_20k(rid, s, ev))
        key = tuple((c["state"], c["date"], c["rounds"]) for c in cells)
        groups.setdefault(key, ([], cells))[0].append(rid)
    return list(groups.values())


def render_html(reports: list[dict], originator: str = "") -> str:
    from .rule_info import compress_ids
    devices = sorted({r["device"] for rep in reports for r in rep["rows"] if r["device"]})
    all_ids = [i for rep in reports for i in rep["ids"]]
    when = reports[0]["time"] if reports else ""
    head_status = ("未放开" if all(r["status"] == "failed" for r in reports) else
                   "部分完成" if any(r["status"] in ("partial", "failed") for r in reports) else "放量完成")
    tone = {"放量完成": "#1e40af", "部分完成": "#b45309", "未放开": "#b91c1c"}[head_status]
    sections = []
    for rep in reports:
        n_ok = len(rep["done"])
        st_tone = {"progress": "#047857", "partial": "#b45309", "failed": "#b91c1c", "none": "#6b7280"}[rep["status"]]
        bars = "".join(
            f'<div style="margin:8px 0 14px 0;"><div style="{FONT}font-size:12px;color:#6b7280;margin-bottom:6px;">'
            f'{_h("全部 " + str(len(gids)) + " 条" if len(progress_groups(rep["ids"])) == 1 else compress_ids(gids))}'
            f'</div>{_bar_html(cells)}</div>'
            for gids, cells in progress_groups(rep["ids"]))
        th = "".join(f'<th align="left" style="{FONT}font-size:12px;color:#374151;background:#f3f4f6;padding:8px 9px;'
                     f'border-bottom:2px solid #e5e7eb;white-space:nowrap;">{h}</th>'
                     for h in ("Rule ID", "机型", "扩展码", "源版本", "目标版本", "成功率", "源版本用户数",
                               "累计检测数", "结果"))
        trs = []
        for i, r in enumerate(rep["rows"]):
            bg = "#ffffff" if i % 2 == 0 else "#fafafa"
            res_tone = "#047857" if r["ok"] else ("#b91c1c" if r["result"].startswith("失败") else "#6b7280")
            users = "-" if r["users"] is None else f"{r['users']:,}"
            cells = [f'<b style="color:#1d4ed8;">{_h(r["ruleid"])}</b>', _h(r["model"] or "-"), _h(r["ext"] or "-"),
                     f'<span style="font-family:Menlo,Consolas,monospace;">{_h(r["src"] or "-")}</span>',
                     f'<span style="font-family:Menlo,Consolas,monospace;">{_h(r["tgt"] or "-")}</span>',
                     _h(r["rate"] or "-"), users, _h(r["count"] if r["count"] != "" else "-"),
                     f'<span style="color:{res_tone};font-weight:600;">{_h(r["result"])}</span>']
            trs.append(f'<tr style="background:{bg};">' + "".join(
                f'<td style="{FONT}font-size:12px;color:#111827;padding:7px 9px;border-bottom:1px solid #f0f0f0;'
                f'white-space:nowrap;">{c}</td>' for c in cells) + "</tr>")
        dass = {True: "已提交", False: "未提交"}.get(rep["dass_submitted"], "-")
        sections.append(f"""
  <tr><td style="padding:18px 26px 4px 26px;">
    <div style="{FONT}font-size:16px;font-weight:700;color:#111827;">{_h(rep["from"])} → {_h(rep["to"])}
      <span style="font-size:12px;font-weight:600;color:{st_tone};margin-left:8px;">{_h(STATUS_TEXT[rep["status"]])}
      · 放开 {n_ok}/{len(rep["ids"])} 条 · DASS {dass}</span></div>
    <div style="{FONT}font-size:12px;color:#6b7280;margin-top:3px;">Rule ID {_h(compress_ids(rep["ids"]))}</div>
  </td></tr>
  <tr><td style="padding:4px 26px 0 26px;">
    <div style="{FONT}font-size:13px;font-weight:600;color:#374151;margin:6px 0 2px 0;">灰度进度</div>{bars}
  </td></tr>
  <tr><td style="padding:0 26px;">
    <div style="overflow-x:auto;border:1px solid #e5e7eb;border-radius:8px;">
      <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
        <thead><tr>{th}</tr></thead><tbody>{"".join(trs)}</tbody></table></div>
  </td></tr>""")
    return f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>OTA 灰度放开汇总</title></head>
<body style="margin:0;padding:0;background:#eef2f7;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#eef2f7;padding:24px 0;">
<tr><td align="center">
<table role="presentation" width="820" cellpadding="0" cellspacing="0"
       style="max-width:820px;width:100%;background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 2px 8px rgba(15,23,42,0.08);">
  <tr><td style="background-color:{tone};padding:22px 26px;">
    <div style="{FONT}font-size:12px;letter-spacing:1px;color:#dbeafe;">HISENSE OVERSEAS OTA · GRAY RELEASE</div>
    <div style="{FONT}font-size:22px;font-weight:700;color:#ffffff;margin-top:6px;">📈 OTA 灰度放开汇总 · {_h(head_status)}</div>
    <div style="{FONT}font-size:13px;color:#e0e7ff;margin-top:6px;">
      {_h(" / ".join(devices) or "-")} &nbsp;·&nbsp; Rule ID {_h(compress_ids(all_ids))}（{len(all_ids)} 条）
      &nbsp;·&nbsp; {_h(when)}{(" &nbsp;·&nbsp; 发起人 " + _h(originator)) if originator else ""}</div>
  </td></tr>
  {"".join(sections)}
  <tr><td style="padding:18px 26px 0 26px;"></td></tr>
  <tr><td style="background:#f8fafc;border-top:1px solid #e5e7eb;padding:12px 26px;">
    <div style="{FONT}font-size:11px;color:#9ca3af;">此邮件由 deployota update 自动生成（阶段以 CMS『在线规则』为准）。请勿直接回复。</div>
  </td></tr>
</table></td></tr></table></body></html>"""


def subject_for(reports: list[dict]) -> str:
    from .rule_info import compress_ids
    devices = sorted({r["device"] for rep in reports for r in rep["rows"] if r["device"]})
    ids = [i for rep in reports for i in rep["ids"]]
    moves = " / ".join(dict.fromkeys(f"{rep['from']} → {rep['to']}" for rep in reports))
    dev = devices[0] if len(devices) == 1 else f"{len(devices)} 个设备类型"
    tag = ("【失败】" if all(r["status"] == "failed" for r in reports) else
           "【部分失败】" if any(r["status"] in ("partial", "failed") for r in reports) else "")
    return f"{tag}[OTA 灰度放开] {dev} {compress_ids(ids)} {moves}（{len(ids)} 条）"


def text_for(reports: list[dict]) -> str:
    lines = []
    for rep in reports:
        lines.append(f"{rep['from']} → {rep['to']}：{STATUS_TEXT[rep['status']]}，放开 {len(rep['done'])}/{len(rep['ids'])}")
        for r in rep["rows"]:
            lines.append(f"  {r['ruleid']}  {r['model']}  {r['src']} → {r['tgt']}  {r['rate'] or '-'}  {r['result']}")
    return "\n".join(lines)


# ----------------------------------------------------------------------
# 收尾：记录 + 存 HTML + 发邮件
# ----------------------------------------------------------------------

def finalize(reports: list[dict], send: bool = True) -> None:
    from . import depdb
    from .mailer import RECIPIENTS_ENV, fixed_bcc, fixed_cc, parse_recipients, send_html
    from .stage_db import append_run
    reports = [r for r in reports if r]
    if not reports or all(r.get("dry_run") for r in reports):
        return
    to_send = [r for r in reports if r["status"] != "none"]
    mail_note = ""
    html_path = ""
    if to_send:
        ids = {i for r in to_send for i in r["ids"]}
        who = sorted({str(b.get("发起人", "") or "").strip() for b in depdb.load_all()
                      if ids & {x.strip() for x in str(b.get("Rule ID列表", "")).split(",")}} - {""})
        html = render_html(to_send, " / ".join(who))
        d = os.path.join(depdb.db_dir(), "reports")
        os.makedirs(d, exist_ok=True)
        html_path = os.path.join(d, f"update_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html)
        print(color(f"✓ 放量汇总 HTML：{html_path}", "green"))
        if not send:
            mail_note = "--no-mail：未发送"
        else:
            cc_fixed, bcc = fixed_cc(), fixed_bcc()
            extra, bad = parse_recipients(os.environ.get(RECIPIENTS_ENV, ""))
            if bad:
                print(color(f"  ⚠ --mail-to 里这些看不懂，已忽略：{'; '.join(bad)}", "yellow"))
            skip = {a.lower() for _, a in cc_fixed + bcc}
            to = cc_fixed + [(n, a) for n, a in extra if a.lower() not in skip]
            ok, mail_note = send_html(subject_for(to_send), html, text_for(to_send), to=to, cc=[], bcc=bcc)
        print(color(("✓ 放量汇总邮件：" if mail_note.startswith("已发送") else "· 放量汇总邮件：") + mail_note,
                    "green" if mail_note.startswith("已发送") else "yellow"))
    else:
        mail_note = "无变化，不发邮件"
        print(color("· 放量汇总邮件：没有规则放开、也没有失败（或你取消了），不发。", "dim"))
    for r in reports:
        append_run({k: r[k] for k in ("time", "requested", "action", "from", "to", "ids", "done", "already",
                                      "failed", "mismatch", "status", "dass_submitted", "sentence")}
                   | {"rates": {x["ruleid"]: x["rate"] for x in r["rows"] if x["rate"]},
                      "users": {x["ruleid"]: x["users"] for x in r["rows"] if x["users"] is not None},
                      "mail": mail_note if r["status"] != "none" else "无变化，不发邮件", "report": html_path})


def _selftest() -> int:
    good = True

    def check(name, cond):
        nonlocal good
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        good = good and bool(cond)

    cells = [{"stage": k, "label": k, "state": s, "date": d, "rounds": n} for k, s, d, n in
             [("场测", "done", "11/05", 0), ("50", "done", "11/06", 0), ("200", "done", "", 0),
              ("1K", "done", "", 0), ("5K", "done", "11/14", 0), ("20K", "current", "11/18", 2), ("完全开放", "todo", "", 0)]]
    b = _bar_html(cells)
    check("进度条：7 格、已完成 ✓、本次 ●、未到 ○、20K ×2", b.count("✓") == 5 and b.count("●") == 1
          and b.count("○") == 1 and "×2" in b and "11/18" in b)
    rep = {"time": "2026-11-18 10:00:00", "from": "OTA5000", "to": "OTA20000", "ids": ["1", "2"], "done": ["1"],
           "already": [], "failed": [("2", "执行后累计检测数不对")], "mismatch": [], "status": "partial",
           "dass_submitted": True, "sentence": "同意（OTA20000，20261118，1）1:97%", "steps": [],
           "rows": [{"ruleid": "1", "device": "DEV", "model": "m", "ext": "e", "src": "0100AQ0101", "tgt": "0100BQ0202",
                     "rate": "97%", "users": 35289, "result": "已放开", "count": 26255, "ok": True},
                    {"ruleid": "2", "device": "DEV", "model": "m", "ext": "e2", "src": "0100AQ0101", "tgt": "0100BQ0202",
                     "rate": "", "users": None, "result": "失败：执行后累计检测数不对", "count": "", "ok": False}]}
    check("主题：部分失败、设备、阶段", subject_for([rep]) == "【部分失败】[OTA 灰度放开] DEV 1~2 OTA5000 → OTA20000（2 条）")
    check("主题：全部成功没有标记", subject_for([dict(rep, status="progress")]).startswith("[OTA 灰度放开] DEV"))
    import tempfile
    db = tempfile.mkdtemp()
    old = os.environ.get("OTA_DEPLOY_DB_DIR")
    os.environ["OTA_DEPLOY_DB_DIR"] = db
    try:
        h = render_html([rep], "康凯")
        check("HTML：标题、进度、明细、用户数", "OTA 灰度放开汇总" in h and "OTA5000 → OTA20000" in h
              and "35,289" in h and "失败：执行后累计检测数不对" in h and "康凯" in h)
        check("HTML：不含同意语句和操作步骤", "同意（" not in h and "<ul" not in h)
    finally:
        if old is None:
            os.environ.pop("OTA_DEPLOY_DB_DIR", None)
        else:
            os.environ["OTA_DEPLOY_DB_DIR"] = old
    print("\n自测结果：", "全部通过 ✅" if good else "存在失败 ❌")
    return 0 if good else 1


if __name__ == "__main__":
    import sys
    sys.exit(_selftest() if "--selftest" in sys.argv[1:] else 0)
