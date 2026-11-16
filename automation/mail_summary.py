"""deployota mail <Rule ID> —— 给已经部署的批次补发 HTML 部署汇总邮件。

  ./deployota mail 65391
  ./deployota mail 65391 --originator "康凯(Corki)"
  ./deployota mail 65391 --mail-to "王敏 <wangmin3@hisense.com>; 于洪 <yuhong@hisense.com>"

流程：按 Rule ID 在部署库找到批次和它的 Excel → 终端打印汇总、生成 HTML（存到 Excel 旁边 <Excel 名>.summary.html）
→ 收件人输入框（可粘贴 Outlook 收件人串；--mail-to 给了就不问）→ 显示草稿 → 输入 approve 才发送。
固定抄送『数据统计报告 <sys.datareports@hisense.com>』，密送王涛（devinw@hisense.com）；To / Cc 里不含 devinw。阶段写『场测 · OTA 部署已在 CMS 上线』。
不碰 CMS / DASS / S3，不改部署库。
"""
from __future__ import annotations

import argparse
import os
import sys

from ._log import color


def find_batch(rule_id: str) -> tuple[dict | None, str]:
    """→ (部署库里包含这个 Rule ID 的批次，问题说明)。有多个时取『有效』里最新的。"""
    from .depdb import STATUS_ACTIVE, _rid_in_record, load_all, sort_newest_first
    hits = [r for r in load_all() if _rid_in_record(r, rule_id)]
    if not hits:
        return None, f"部署库里没有 Rule ID {rule_id}（只能给通过 publish / dass init 入库的批次发汇总）"
    active = [r for r in hits if str(r.get("状态", STATUS_ACTIVE)) == STATUS_ACTIVE] or hits
    return sort_newest_first(active)[0], ""


def main() -> int:
    from .mailer import RECIPIENTS_ENV, SkipMail, approve_and_send, ask_recipients, plan_recipients
    p = argparse.ArgumentParser(prog="deployota mail", description="给已经部署的批次补发 HTML 部署汇总邮件。")
    p.add_argument("rule_id", help="这批里的任意一个 Rule ID，如 65391")
    p.add_argument("--originator", default=None, help="发起人姓名（不给则用 prep 时记的 / 部署库里的）")
    p.add_argument("--mail-to", dest="mail_to", default=None, help="收件人（Outlook 收件人串）；不给就弹出输入框")
    args = p.parse_args()

    rec, problem = find_batch(args.rule_id.strip())
    if not rec:
        print(color(f"✗ {problem}", "red"))
        return 2
    excel = str(rec.get("Excel路径", "") or "")
    if not excel or not os.path.isfile(excel):
        print(color(f"✗ 这批的 Excel 找不到了：{excel or '（部署库没记）'}", "red"))
        return 2
    from .last_prep import originator_for
    originator = (args.originator if args.originator is not None
                  else originator_for(excel) or str(rec.get("发起人", "") or "")).strip()
    print(f"批次：{rec.get('设备类型代码', '')}  Rule ID {rec.get('Rule ID范围', '')}  "
          f"部署时间 {rec.get('部署时间', '')}  发起人 {originator or '-'}")
    print(f"Excel：{excel}")

    from .dass_fill import FIRST_DEPLOY_STAGE
    from .summary import build_summary_html, print_summary
    print("")
    print_summary(excel, originator=originator)
    subject, html, text = build_summary_html(excel, originator=originator, stage=FIRST_DEPLOY_STAGE)
    html_path = os.path.splitext(os.path.abspath(excel))[0] + ".summary.html"

    if args.mail_to:
        os.environ[RECIPIENTS_ENV] = args.mail_to
    try:
        to, cc, bcc = plan_recipients(ask_recipients())
    except SkipMail:
        print(color("· 你选择不发邮件。", "dim"))
        return 0
    ok, why = approve_and_send(subject, html, text, to, cc, bcc, draft_path=html_path)
    print(color(("✓ 汇总邮件：" if ok else "· 汇总邮件：") + why, "green" if ok else "yellow"))
    return 0 if ok or "approve" in why else 1


if __name__ == "__main__":
    sys.exit(main())
