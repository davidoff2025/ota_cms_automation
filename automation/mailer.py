"""把部署汇总（HTML）发邮件。SMTP 设置全部来自 config/.env，代码里不写任何账号密码。

config/.env 里需要：
  OTA_SMTP_HOST      邮件服务器，例如 smtp.hisense.com（问 IT 要）
  OTA_SMTP_PORT      端口：465（SSL）或 587（STARTTLS）或 25；默认 587
  OTA_SMTP_SECURITY  ssl / starttls / none；默认按端口推断（465→ssl，其余→starttls）
  OTA_SMTP_USER      登录账号（通常是邮箱地址）；服务器不需要登录时留空
  OTA_SMTP_PASS      密码 / 授权码
  OTA_MAIL_FROM      发件人；默认 = OTA_SMTP_USER
  OTA_MAIL_TO        收件人，逗号分隔；默认 devinw@hisense.com

没配 OTA_SMTP_HOST 时不发送，只提示；HTML 文件照样保存，可以手工打开或转发。
"""
from __future__ import annotations

import os
import re
import smtplib
import ssl
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid

from . import config  # noqa: F401  加载 config/.env

DEFAULT_TO = "devinw@hisense.com"
RECIPIENTS_ENV = "OTA_MAIL_RECIPIENTS"     # --mail-to 传进来的收件人（publish → dass init 子进程靠它继承）

_EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-']+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")


def parse_recipients(text: str) -> tuple[list[tuple[str, str]], list[str]]:
    """解析 Outlook 风格的收件人串 → ([(显示名, 邮箱), ...], [看不懂的片段])。

    支持：'康凯(Corki) <kangkai3@hisense.com>; 王涛/Devin Wang <wangtao14@hisense.com>'、
    '"Name" <a@x.com>'、裸地址 'a@x.com'；分隔符 ; ； , ， 换行都行。
    显示名原样保留（括号、斜杠都不丢）；按邮箱去重（不区分大小写），保持顺序。
    """
    out: list[tuple[str, str]] = []
    bad: list[str] = []
    seen: set[str] = set()

    def add(name: str, addr: str):
        addr = addr.strip().strip("<>").strip()
        if not _EMAIL_RE.fullmatch(addr):
            bad.append(f"{name} <{addr}>".strip() if name else addr)
            return
        if addr.lower() not in seen:
            seen.add(addr.lower())
            out.append((name.strip().strip('"').strip("'").strip(), addr))

    for part in re.split(r"[;；\n\r]+", str(text or "")):
        part = part.strip()
        if not part:
            continue
        if "<" in part:
            pos = 0
            for m in re.finditer(r"([^<>]*?)<([^<>]*)>", part):
                # 名字前面可能还夹着逗号分隔的裸地址：'a@x.com, "B Name" <b@x.com>'
                segs = re.split(r"[,，]", m.group(1))
                name_segs = []
                for seg in segs:
                    if "@" in seg and _EMAIL_RE.fullmatch(seg.strip()):
                        add("", seg.strip())
                    else:
                        name_segs.append(seg)
                name = ",".join(name_segs).strip().strip(",，").strip()
                add(name, m.group(2))
                pos = m.end()
            rest = part[pos:].strip(" ,，")
            for tok in [t for t in re.split(r"[,，\s]+", rest) if t]:
                (add("", tok) if "@" in tok else bad.append(tok))
        else:
            for tok in [t for t in re.split(r"[,，\s]+", part) if t]:
                (add("", tok) if "@" in tok else bad.append(tok))
    return out, bad


def _fmt(name: str, addr: str) -> str:
    return formataddr((name, addr)) if name else addr


def always_cc() -> list[tuple[str, str]]:
    """每封都抄送：发件人自己 + OTA_MAIL_TO（默认 devinw@hisense.com）。"""
    s = settings()
    got, _ = parse_recipients(", ".join(([s["from"]] if s["from"] else []) + s["to"]))
    return got


def ask_recipients() -> list[tuple[str, str]]:
    """部署完成后问收件人：--mail-to（环境变量）给了就用；否则终端里粘贴（回车 = 只发固定抄送的人）。"""
    preset = os.environ.get(RECIPIENTS_ENV, "").strip()
    if preset:
        got, bad = parse_recipients(preset)
        if bad:
            print(f"  ⚠ --mail-to 里这些看不懂，已忽略：{'; '.join(bad)}")
        return got
    import sys
    if not sys.stdin.isatty():
        return []
    cc = "; ".join(a for _, a in always_cc())
    print(f"\n======== 部署汇总邮件 ========\n固定抄送：{cc}")
    while True:
        print("请把所有收件人一次粘贴进来（可直接从 Outlook 复制，一行或多行都行，"
              "如『王敏 <wangmin3@hisense.com>; …』）。\n粘贴完再按一次回车（空行）结束；什么都不输入直接回车 = 只发固定抄送：")
        text = read_block()
        got, bad = parse_recipients(text)
        if got:
            print(f"  认出 {len(got)} 位收件人：")
            for n, a in got:
                print(f"    ✓ {n + ' ' if n else ''}<{a}>")
        if bad:
            print(f"  ⚠ 看不懂（不会发送）：{'; '.join(bad)}")
        try:
            ans = input("  回车 = 发送；r = 重新输入；n = 不发邮件：").strip().lower()
        except EOFError:
            ans = ""
        if ans == "r":
            continue
        if ans == "n":
            raise SkipMail()
        return got


class SkipMail(Exception):
    """你在收件人确认时选了 n：这次不发邮件。"""


def read_block() -> str:
    """读多行输入，直到空行（或 EOF）。第一行就是空行 → 返回空串。"""
    lines: list[str] = []
    while True:
        try:
            line = input("> " if not lines else "  ")
        except EOFError:
            break
        if not line.strip():
            break
        lines.append(line)
    return "\n".join(lines)


def pop_mail_to(argv: list[str]) -> list[str]:
    """从命令行取出 --mail-to "…"（或 --mail-to=…），放进环境变量；返回去掉它之后的 argv。"""
    out, i = [], 0
    while i < len(argv):
        a = argv[i]
        if a == "--mail-to" and i + 1 < len(argv):
            os.environ[RECIPIENTS_ENV] = argv[i + 1]
            i += 2
            continue
        if a.startswith("--mail-to="):
            os.environ[RECIPIENTS_ENV] = a.split("=", 1)[1]
            i += 1
            continue
        out.append(a)
        i += 1
    return out


def settings() -> dict:
    g = lambda k, d="": os.environ.get(k, d).strip()
    port = int(g("OTA_SMTP_PORT", "587") or 587)
    sec = g("OTA_SMTP_SECURITY").lower() or ("ssl" if port == 465 else "starttls")
    user = g("OTA_SMTP_USER")
    return {"host": g("OTA_SMTP_HOST"), "port": port, "security": sec, "user": user,
            "password": g("OTA_SMTP_PASS"), "from": g("OTA_MAIL_FROM") or user,
            "to": [x.strip() for x in (g("OTA_MAIL_TO") or DEFAULT_TO).replace(";", ",").split(",") if x.strip()]}


def build_message(subject: str, html: str, text: str, sender: str, to: list[str],
                  cc: list[str] | None = None) -> EmailMessage:
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=(sender.split("@")[-1] if "@" in sender else None))
    msg.set_content(text or "请用支持 HTML 的邮件客户端查看。")
    msg.add_alternative(html, subtype="html")
    return msg


def plan_recipients(extra: list[tuple[str, str]]) -> tuple[list[tuple[str, str]], list[tuple[str, str]]]:
    """→ (收件人 To, 抄送 Cc)。To = 你输入的人；Cc = 固定抄送（发件人 + devinw）中不在 To 里的。
    没输入任何人时，固定抄送的人直接当收件人。"""
    cc = always_cc()
    if not extra:
        return cc, []
    to_set = {a.lower() for _, a in extra}
    return list(extra), [(n, a) for n, a in cc if a.lower() not in to_set]


def send_html(subject: str, html: str, text: str = "",
              to: list[tuple[str, str]] | None = None, cc: list[tuple[str, str]] | None = None) -> tuple[bool, str]:
    """发送。to/cc 不给 = 发给固定抄送的人。返回 (是否成功, 说明)。失败不抛异常（不影响部署流程）。"""
    s = settings()
    if not s["host"]:
        return False, "没配邮件服务器（config/.env 的 OTA_SMTP_HOST），未发送"
    if not s["from"]:
        return False, "没配发件人（OTA_MAIL_FROM 或 OTA_SMTP_USER），未发送"
    if to is None:
        to, cc = plan_recipients([])
    to_s, cc_s = [_fmt(n, a) for n, a in to], [_fmt(n, a) for n, a in (cc or [])]
    if not to_s:
        return False, "没有收件人，未发送"
    msg = build_message(subject, html, text, s["from"], to_s, cc_s)
    try:
        ctx = ssl.create_default_context()
        if s["security"] == "ssl":
            server = smtplib.SMTP_SSL(s["host"], s["port"], timeout=30, context=ctx)
        else:
            server = smtplib.SMTP(s["host"], s["port"], timeout=30)
        with server:
            server.ehlo()
            if s["security"] == "starttls":
                server.starttls(context=ctx)
                server.ehlo()
            if s["user"] and s["password"]:
                server.login(s["user"], s["password"])
            server.send_message(msg)   # 收件人取自 To + Cc 头
        names = lambda lst: "、".join(n or a for n, a in lst)
        return True, f"已发送：收件人 {names(to)}" + (f"；抄送 {names(cc)}" if cc else "")
    except Exception as e:   # 不把密码等写进日志：只给异常类型和服务器返回
        return False, f"发送失败（{s['host']}:{s['port']} {s['security']}）：{type(e).__name__}: {e}"


def _selftest() -> int:
    good = True

    def check(name, cond):
        nonlocal good
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        good = good and bool(cond)

    saved = {k: os.environ.pop(k, None) for k in ("OTA_SMTP_HOST", "OTA_SMTP_PORT", "OTA_SMTP_SECURITY",
                                                   "OTA_MAIL_TO", "OTA_SMTP_USER", "OTA_MAIL_FROM")}
    try:
        check("默认收件人 devinw@hisense.com", settings()["to"] == [DEFAULT_TO])
        os.environ["OTA_SMTP_PORT"] = "465"
        check("465 端口默认 SSL", settings()["security"] == "ssl")
        os.environ["OTA_MAIL_TO"] = "a@x.com; b@x.com"
        check("多个收件人（; 或 , 分隔）", settings()["to"] == ["a@x.com", "b@x.com"])
        ok_, why = send_html("s", "<b>x</b>")
        check("没配服务器 → 不发送、给出原因", not ok_ and "OTA_SMTP_HOST" in why)
        m = build_message("主题", "<p>正文</p>", "正文", "ota@x.com", ["devinw@hisense.com"])
        parts = [p.get_content_type() for p in m.iter_parts()]
        check("邮件含纯文本 + HTML 两部分", parts == ["text/plain", "text/html"] and m["To"] == "devinw@hisense.com")

        sample = ("康凯(Corki) <kangkai3@hisense.com>; 王敏 <wangmin3@hisense.com>; 王涛/Devin Wang "
                  "<wangtao14@hisense.com>; 于洪 <yuhong@hisense.com>; 王泽纯 <wangzechun@hisense.com>; "
                  "黄金花 <huangjinhua@hisense.com>; 庄广海(Stevenson) <zhuangguanghai@hisense.com>")
        got, bad = parse_recipients(sample)
        check("Outlook 收件人串：7 人全部认出、无看不懂", len(got) == 7 and not bad)
        check("显示名原样保留（括号 / 斜杠）", got[0] == ("康凯(Corki)", "kangkai3@hisense.com")
              and got[2] == ("王涛/Devin Wang", "wangtao14@hisense.com")
              and got[6] == ("庄广海(Stevenson)", "zhuangguanghai@hisense.com"))
        g2, b2 = parse_recipients('a@x.com, "B Name" <b@x.com>；A@X.com\nc@x.com，坏地址 <oops>')
        check("裸地址 / 引号名 / 全角分隔 / 去重（不分大小写）",
              [a for _, a in g2] == ["a@x.com", "b@x.com", "c@x.com"] and g2[1][0] == "B Name")
        check("坏地址列进看不懂", b2 == ["坏地址 <oops>"])
        check("空输入", parse_recipients("  ") == ([], []))
        os.environ["OTA_MAIL_FROM"] = "sys.datareports@hisense.com"
        os.environ["OTA_MAIL_TO"] = "devinw@hisense.com"
        to, cc = plan_recipients(got[:2])
        check("To = 输入的人；Cc = 发件人 + devinw", [a for _, a in to] == ["kangkai3@hisense.com", "wangmin3@hisense.com"]
              and [a for _, a in cc] == ["sys.datareports@hisense.com", "devinw@hisense.com"])
        to2, cc2 = plan_recipients([("", "devinw@hisense.com")])
        check("输入里已有 devinw → 不重复抄送", [a for _, a in cc2] == ["sys.datareports@hisense.com"])
        to3, cc3 = plan_recipients([])
        check("没输入 → 固定的人当收件人", [a for _, a in to3] == ["sys.datareports@hisense.com", "devinw@hisense.com"]
              and cc3 == [])
        m2 = build_message("s", "<b>x</b>", "x", "sys@x.com", [_fmt(*got[0])], [_fmt(*cc[0])])
        check("中文显示名编码进 To / Cc 头", "kangkai3@hisense.com" in str(m2["To"]) and "Cc" in m2)
        check("pop_mail_to 取出参数", pop_mail_to(["init", "--mail-to", "a@x.com", "x.xlsx"]) == ["init", "x.xlsx"]
              and os.environ.get(RECIPIENTS_ENV) == "a@x.com")
        os.environ.pop(RECIPIENTS_ENV, None)

        # 输入框：多行粘贴 + 确认（模拟终端输入）
        import builtins
        import contextlib
        import io
        import sys as _sys
        real_input, real_tty = builtins.input, _sys.stdin.isatty

        def run(lines):
            feed = iter(lines)
            builtins.input = lambda *_a: next(feed)
            _sys.stdin.isatty = lambda: True
            try:
                with contextlib.redirect_stdout(io.StringIO()):
                    return ask_recipients()
            finally:
                builtins.input, _sys.stdin.isatty = real_input, real_tty
        r1 = run(["康凯(Corki) <kangkai3@hisense.com>; 王敏 <wangmin3@hisense.com>",
                  "于洪 <yuhong@hisense.com>", "", ""])
        check("多行粘贴 + 回车发送", [a for _, a in r1] == ["kangkai3@hisense.com", "wangmin3@hisense.com",
                                                            "yuhong@hisense.com"])
        check("直接回车（不输入）+ 回车 = 只发固定抄送", run(["", ""]) == [])
        check("r 重新输入", [a for _, a in run(["a@x.com", "", "r", "b@x.com", "", ""])] == ["b@x.com"])
        try:
            run(["a@x.com", "", "n"])
            check("n = 不发邮件", False)
        except SkipMail:
            check("n = 不发邮件", True)
    finally:
        for k, v in saved.items():
            os.environ.pop(k, None)
            if v is not None:
                os.environ[k] = v
    print("\n自测结果：", "全部通过 ✅" if good else "存在失败 ❌")
    return 0 if good else 1


if __name__ == "__main__":
    import sys
    sys.exit(_selftest() if "--selftest" in sys.argv[1:] else 0)
