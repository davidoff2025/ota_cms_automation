"""模拟 publish --file：真实跑 prep（--no-upload），其余步骤模拟，最后真实发汇总邮件（需输入 approve）。

  ./.venv/bin/python -m automation.sim_publish --file /路径/source.txt [--originator 姓名]

  ⓪ prep：真实执行，但加 --no-upload（不碰 FTP / S3），输出到临时目录（不动你的文件夹）；
          MAC 组只读检查；会问发起人（没给 --originator 时）。
  ①~④：模拟（不建模板、不上线、不动 DASS、不写部署库）；Rule ID 用假的 99901 起，只写进临时 Excel。
  收尾：和真实 publish 一样弹出收件人输入框（可粘贴 Outlook 收件人串），
        显示草稿后，**输入 approve 才真的发送**；其它输入 = 不发。
        邮件主题前加『【模拟测试】』，标题栏写『模拟测试 · 未真实部署（Rule ID 为假）』。
  结束后恢复『最近 prep』记录，删除临时目录；HTML 草稿留在 .agents/mail_drafts/sim_publish.html。
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FAKE_START = 99901
STAGE = "模拟测试 · 未真实部署（Rule ID 为假）"


def main() -> int:
    p = argparse.ArgumentParser(prog="python -m automation.sim_publish",
                                description="模拟 publish --file，最后真实发送汇总邮件（需 approve）。")
    p.add_argument("--file", required=True, help="原始输入（文本或申请模板 Excel）")
    p.add_argument("--originator", default=None, help="发起人姓名（不给就在 prep 时询问）")
    args = p.parse_args()

    from ._log import color
    src = os.path.join(os.environ.get("OTA_CALLER_CWD") or os.getcwd(), os.path.expanduser(args.file))
    if not os.path.isfile(src):
        print(color(f"✗ 找不到文件：{src}", "red"))
        return 2

    lp = ROOT / ".last_prep.json"
    lp_backup = lp.read_bytes() if lp.exists() else None
    work = tempfile.mkdtemp(prefix="sim_publish_")
    try:
        print(color("\n▶ 检查点0：prep（真实执行，--no-upload：不碰 FTP / S3；输出到临时目录）", "bold"), flush=True)
        cmd = [sys.executable, "-m", "automation.prep", "--file", src, "--no-upload", "--out-dir", work]
        if args.originator is not None:
            cmd += ["--originator", args.originator]
        rc = subprocess.run(cmd, cwd=str(ROOT)).returncode
        if rc != 0:
            print(color(f"✗ 检查点0 失败：prep 退出码 {rc}。模拟到此为止。", "red"))
            return 10
        excel = next(str(x) for x in Path(work).glob("*.xlsx"))

        from .active_ota import _publish_mail, format_checkpoint_report, read_import_excel, write_rule_id_columns
        n = read_import_excel(excel)["count"]
        fake = [str(FAKE_START + i) for i in range(n)]
        write_rule_id_columns(excel, fake)
        print(color(f"\n▶ 检查点1~4：模拟（CMS / DASS / 部署库都没动）；假 Rule ID {fake[0]}~{fake[-1]}（{n} 个）", "bold"))
        print("\n" + format_checkpoint_report(
            [{"n": 0, "status": "pass", "detail": "真实执行（--no-upload）"},
             {"n": 1, "status": "pass", "detail": "模拟"},
             {"n": 2, "status": "pass", "detail": f"模拟 {fake[0]}~{fake[-1]}"},
             {"n": 3, "status": "pass", "detail": "模拟"},
             {"n": 4, "status": "pass", "detail": "模拟"}]))

        # —— 收尾：和真实 publish 一样问收件人；发送前多一道 approve ——
        from . import mailer as M
        from . import summary as S
        real_send, orig_build = M.send_html, S.build_summary_html
        S.build_summary_html = lambda *a, **k: (lambda r: ("【模拟测试】" + r[0], r[1], r[2]))(
            orig_build(*a, **{**k, "stage": STAGE}))

        def gated_send(subject, html, text="", to=None, cc=None, bcc=None):
            draft = ROOT / ".agents" / "mail_drafts" / "sim_publish.html"
            draft.parent.mkdir(parents=True, exist_ok=True)
            draft.write_text(html, encoding="utf-8")
            print(color("\n======== 邮件草稿（还没发送）========", "bold"))
            print(f"  主题：{subject}")
            print(f"  收件人：{'；'.join(f'{nm} <{a}>' if nm else a for nm, a in (to or []))}")
            print(f"  抄送：{'；'.join(a for _, a in (cc or [])) or '（无）'}")
            print(f"  密送：{'；'.join(a for _, a in (bcc or [])) or '（无）'}")
            print(f"  正文预览：{draft}（可以用浏览器打开看）")
            try:
                ans = input("输入 approve 真的发送；其它任意输入 = 不发：").strip().lower()
            except EOFError:
                ans = ""
            if ans != "approve":
                return False, "你没有输入 approve，未发送"
            return real_send(subject, html, text, to=to, cc=cc, bcc=bcc)
        M.send_html = gated_send
        _publish_mail(excel, True)
        return 0
    finally:
        if lp_backup is None:
            lp.unlink(missing_ok=True)
        else:
            lp.write_bytes(lp_backup)
        shutil.rmtree(work, ignore_errors=True)
        print(color("（模拟结束：已恢复『最近 prep』记录，临时目录已删除）", "dim"))


if __name__ == "__main__":
    sys.exit(main())
