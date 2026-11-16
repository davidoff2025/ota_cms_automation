"""prep 最终检查：生成的 Excel + CMS MAC 组 + S3 升级包，全部核对一遍，任一项 ✗ 则 prep 不算完成。

检查项（每项 pass / warn / fail / skip）：
  1) 场测 MAC 组  ：输入里的全部 MAC（后 8 位）都在 CMS『测试+设备类型代码』组里（重新只读回查 CMS）；
                    Excel 每行『MAC组』列 = 测试+设备类型代码。
  2) 升级提示语    ：表头 语种/描述；至少 1 行；第 1 行是 简体中文；语种都在标准列表里且不重复；描述不为空；
                    描述里没有制表符（制表符 = 另一个语种被并进来了，如『Indonesia<Tab>Fitur…』）。
                    简体中文 与 英语 描述完全相同（源数据没给中文、按规则复制了英语）→ ⚠ 提醒。
  3) 升级包 / S3   ：每行『升级文件地址』= https://ota-dl.intsmarthub.com/ota/<设备类型代码>/…/xxx.zip；
                    按地址算出 S3 key，head_object 存在且大小 = Excel『文件大小』；
                    S3 能直接给出 MD5（单段上传的 ETag）时再比 MD5；本次下载的文件已在上传前逐个比过 MD5。
  4) 规则行数      ：Excel 数据行数 = 解析出的行数。

纯逻辑函数不连网络（_selftest 覆盖）；S3 / CMS 部分由 prep 调用时注入。
"""
from __future__ import annotations

import hashlib
import os
import re
from urllib.parse import urlparse

from openpyxl import load_workbook

from ._log import color, err, ok, warn
from .otagen.legacy_engine import STANDARD_MESSAGE_LANGUAGES

ICON = {"pass": "✓", "warn": "⚠", "fail": "✗", "skip": "·"}


def item(name: str, status: str, detail: str, lines: list[str] | None = None) -> dict:
    return {"name": name, "status": status, "detail": detail, "lines": list(lines or [])}


# ----------------------------------------------------------------------
# 读 Excel
# ----------------------------------------------------------------------

def read_sheet(xlsx: str, title: str) -> tuple[list[str], list[list]]:
    """返回 (表头, 非空数据行)。没有这个 sheet 返回 ([], [])。"""
    wb = load_workbook(xlsx, data_only=True)
    try:
        if title not in wb.sheetnames:
            return [], []
        ws = wb[title]
        rows = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()
    if not rows:
        return [], []
    head = [str(v or "").strip() for v in rows[0]]
    data = [list(r) for r in rows[1:] if any(v not in (None, "") for v in r)]
    return head, data


def _col(head: list[str], name: str) -> int | None:
    return head.index(name) if name in head else None


def _s(v) -> str:
    return "" if v is None else str(v).strip()


# ----------------------------------------------------------------------
# 2) 升级提示语
# ----------------------------------------------------------------------

def check_messages(head: list[str], rows: list[list]) -> dict:
    name = "升级提示语"
    if head[:2] != ["语种", "描述"]:
        return item(name, "fail", f"表头应为『语种 / 描述』，实际：{head[:2]}")
    if not rows:
        return item(name, "fail", "没有任何提示语")
    msgs = [(_s(r[0]), _s(r[1]) if len(r) > 1 else "") for r in rows]
    langs = [l for l, _ in msgs]
    lines = [f"{i}. {l}：{' '.join(d.split())[:40]}{'…' if len(d) > 40 else ''}" for i, (l, d) in enumerate(msgs, 1)]
    fails, warns = [], []
    if langs[0] != "简体中文":
        fails.append(f"第 1 行是『{langs[0]}』，必须是『简体中文』")
    bad = [l for l in langs if l not in STANDARD_MESSAGE_LANGUAGES]
    if bad:
        fails.append(f"不认识的语种：{', '.join(bad)}")
    dup = sorted({l for l in langs if langs.count(l) > 1})
    if dup:
        fails.append(f"语种重复：{', '.join(dup)}")
    empty = [l for l, d in msgs if not d]
    if empty:
        fails.append(f"描述为空：{', '.join(empty)}")
    for l, d in msgs:
        if "\t" in d:
            seg = next((ln for ln in d.splitlines() if "\t" in ln), "")
            fails.append(f"『{l}』的描述里有制表符，疑似把另一个语种并进来了：『{' '.join(seg.split())[:40]}…』")
    by = dict(msgs)
    if by.get("简体中文") and by.get("简体中文") == by.get("英语"):
        warns.append("简体中文 与 英语 描述完全相同（源数据没给中文提示语，按规则复制了英语）")
    st = "fail" if fails else ("warn" if warns else "pass")
    detail = "；".join(fails + warns) or f"{len(msgs)} 种语言，简体中文在第一项"
    return item(name, st, detail, lines)


# ----------------------------------------------------------------------
# 1b) / 3) / 4) 升级规则表
# ----------------------------------------------------------------------

def rule_rows(head: list[str], rows: list[list]) -> list[dict]:
    want = {"设备类型代码": "code", "内部机型信息": "model", "MAC组": "mac_group", "升级文件地址": "url",
            "升级文件的MD5值": "md5", "文件大小": "size"}
    idx = {k: _col(head, h) for h, k in want.items()}
    out = []
    for i, r in enumerate(rows, 1):
        d = {k: (_s(r[j]) if j is not None and j < len(r) else "") for k, j in idx.items()}
        d["row"] = i
        out.append(d)
    return out


def check_mac_column(rr: list[dict], group: str) -> list[str]:
    """Excel 每行『MAC组』列必须 = 测试+设备类型代码。返回问题列表。"""
    return [f"第{r['row']}行 {r['model']}：MAC组『{r['mac_group'] or '（空）'}』≠『{group}』"
            for r in rr if r["mac_group"] != group]


def s3_key_from_url(url: str, aws_prefix: str) -> str:
    """'https://ota-dl.intsmarthub.com/ota/X/20261116/a.zip' → 'ota/X/20261116/a.zip'；前缀不对返回空串。"""
    if not url.lower().startswith(aws_prefix.lower().rstrip("/") + "/"):
        return ""
    return urlparse(url).path.lstrip("/")


def url_problems(rr: list[dict], device_code: str, aws_prefix: str) -> list[str]:
    out = []
    for r in rr:
        key = s3_key_from_url(r["url"], aws_prefix)
        if not key:
            out.append(f"第{r['row']}行：地址不是 {aws_prefix}/…：{r['url'] or '（空）'}")
        elif not key.startswith(f"ota/{device_code}/"):
            out.append(f"第{r['row']}行：地址不在 ota/{device_code}/ 下：{r['url']}")
        elif not key.lower().endswith(".zip"):
            out.append(f"第{r['row']}行：不是 .zip：{r['url']}")
        if not re.fullmatch(r"(?i)[0-9a-f]{32}", r["md5"]):
            out.append(f"第{r['row']}行：MD5 不是 32 位十六进制：{r['md5'] or '（空）'}")
        if not re.fullmatch(r"\d+", re.sub(r"[,\s]", "", r["size"])) or int(re.sub(r"[,\s]", "", r["size"])) <= 0:
            out.append(f"第{r['row']}行：文件大小不对：{r['size'] or '（空）'}")
    return out


def check_s3(rr: list[dict], bucket: str, aws_prefix: str, head_fn, local_md5_ok: set[str]) -> dict:
    """head_fn(bucket, key) → {"ContentLength", "ETag"} 或抛异常。local_md5_ok：本次下载后已比过 MD5 的文件名。"""
    name = "升级包 / S3"
    lines, fails, seen = [], [], {}
    for r in rr:
        key = s3_key_from_url(r["url"], aws_prefix)
        if not key or key in seen:
            continue
        seen[key] = True
        fname = os.path.basename(key)
        exp_size = int(re.sub(r"[^\d]", "", r["size"]) or 0)
        try:
            h = head_fn(bucket, key)
        except Exception as e:
            fails.append(f"S3 上没有：s3://{bucket}/{key}")
            lines.append(err(f"{fname}：S3 上没有（{str(e)[:60]}）"))
            continue
        size = h.get("ContentLength")
        etag = str(h.get("ETag", "")).strip('"').lower()
        if size != exp_size:
            fails.append(f"{fname} 大小不一致：S3={size}，Excel={exp_size}")
            lines.append(err(f"{fname}：大小不一致 S3={size} / Excel={exp_size}"))
            continue
        if re.fullmatch(r"[0-9a-f]{32}", etag):
            if etag != r["md5"].lower():
                fails.append(f"{fname} MD5 不一致：S3={etag}，Excel={r['md5'].lower()}")
                lines.append(err(f"{fname}：MD5 不一致"))
                continue
            how = "大小、MD5 都一致"
        elif fname in local_md5_ok:
            how = "大小一致；MD5 已在上传前比对一致（大文件分段上传，S3 没有整体 MD5）"
        else:
            how = "大小一致（S3 已有、本次未重传；大文件分段上传，S3 没有整体 MD5 可比）"
        lines.append(ok(f"{fname}：{how}"))
    if not seen:
        return item(name, "fail", "Excel 里没有可核对的升级文件地址", lines)
    st = "fail" if fails else "pass"
    return item(name, st, "；".join(fails) or f"{len(seen)} 个 zip 都在 s3://{bucket}/ 上", lines)


def file_md5(path: str) -> str:
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8 * 1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


# ----------------------------------------------------------------------
# 报告
# ----------------------------------------------------------------------

def render(items: list[dict]) -> tuple[list[str], bool]:
    """→ (输出行, 是否有 ✗)。"""
    out = [color("======== prep 最终检查 ========", "bold")]
    paint = {"pass": "green", "warn": "yellow", "fail": "red", "skip": "dim"}
    for i, it in enumerate(items, 1):
        out.append(color(f"  {ICON[it['status']]} {i}. {it['name']}：{it['detail']}", paint[it["status"]]))
        for ln in it["lines"]:
            out.append(f"       {ln}")
    failed = [it["name"] for it in items if it["status"] == "fail"]
    if failed:
        out.append(color(f"  ✗ 最终检查未通过：{'、'.join(failed)}。prep 不算完成，请处理后重跑。", "red"))
    elif any(it["status"] == "warn" for it in items):
        out.append(color("  ✓ 最终检查通过（有 ⚠ 提醒，请看一眼）。", "yellow"))
    else:
        out.append(color("  ★ 最终检查全部通过。", "green"))
    return out, bool(failed)


def _selftest() -> int:
    good = True

    def check(name, cond):
        nonlocal good
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        good = good and bool(cond)

    H = ["语种", "描述"]
    check("提示语：简体中文第一 → pass", check_messages(H, [["简体中文", "中"], ["英语", "en"]])["status"] == "pass")
    check("提示语：英语在第一 → fail", check_messages(H, [["英语", "en"], ["简体中文", "中"]])["status"] == "fail")
    check("提示语：中文复制英语 → warn", check_messages(H, [["简体中文", "en"], ["英语", "en"]])["status"] == "warn")
    m = check_messages(H, [["简体中文", "中"], ["英语", "New\nIndonesia\tFitur baru"]])
    check("提示语：描述里夹着另一种语言 → fail", m["status"] == "fail" and "Indonesia" in m["detail"])
    check("提示语：不认识的语种 / 重复 / 空描述 → fail",
          check_messages(H, [["简体中文", "中"], ["Indonesia", "x"]])["status"] == "fail"
          and check_messages(H, [["简体中文", "中"], ["英语", "a"], ["英语", "b"]])["status"] == "fail"
          and check_messages(H, [["简体中文", ""]])["status"] == "fail")
    check("提示语：表头不对 / 没有行 → fail", check_messages(["a", "b"], [["简体中文", "x"]])["status"] == "fail"
          and check_messages(H, [])["status"] == "fail")

    P = "https://ota-dl.intsmarthub.com"
    check("URL → S3 key", s3_key_from_url(f"{P}/ota/X/20261116/a.zip", P) == "ota/X/20261116/a.zip"
          and s3_key_from_url("http://x/a.zip", P) == "")
    rh = ["设备类型代码", "内部机型信息", "MAC组", "升级文件地址", "升级文件的MD5值", "文件大小"]
    md5 = "debfa7e56660272595ba7e257eb3ce44"
    rr = rule_rows(rh, [["X", "65", "测试X", f"{P}/ota/X/20261116/a.zip", md5, "100"],
                        ["X", "75", "测试Y", f"{P}/ota/Z/20261116/b.zip", "abc", "0"]])
    check("MAC组列核对", check_mac_column(rr, "测试X") == ["第2行 75：MAC组『测试Y』≠『测试X』"])
    up = url_problems(rr, "X", P)
    check("地址 / MD5 / 大小核对", len(up) == 3 and not any("第1行" in u for u in up))

    def head_ok(_b, k):
        if k.endswith("a.zip"):
            return {"ContentLength": 100, "ETag": f'"{md5}"'}
        raise RuntimeError("404")
    s = check_s3(rr[:1], "bkt", P, head_ok, set())
    check("S3：大小 + MD5 一致 → pass", s["status"] == "pass")
    s = check_s3(rr[:1], "bkt", P, lambda b, k: {"ContentLength": 99, "ETag": '"x-3"'}, set())
    check("S3：大小不一致 → fail", s["status"] == "fail")
    s = check_s3(rr[:1], "bkt", P, lambda b, k: {"ContentLength": 100, "ETag": '"0123-3"'}, {"a.zip"})
    check("S3：分段上传 + 本地 MD5 已比 → pass", s["status"] == "pass" and "上传前" in s["lines"][0])
    s = check_s3([dict(rr[0], url=f"{P}/ota/X/20261116/c.zip")], "bkt", P, head_ok, set())
    check("S3：不存在 → fail", s["status"] == "fail")
    lines, failed = render([item("a", "pass", "x"), item("b", "warn", "y")])
    check("报告：有 warn 无 fail → 通过", not failed and "通过" in lines[-1])
    lines, failed = render([item("a", "fail", "x")])
    check("报告：有 fail → 未通过", failed and "未通过" in lines[-1])
    print("\n自测结果：", "全部通过 ✅" if good else "存在失败 ❌")
    return 0 if good else 1


if __name__ == "__main__":
    import sys
    sys.exit(_selftest() if "--selftest" in sys.argv[1:] else 0)
