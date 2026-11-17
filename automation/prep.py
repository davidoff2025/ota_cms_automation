"""OTA 模板『prep CLI』：命令行粘贴/传入升级信息文本，三步串成一键流程——

  步骤1 生成：用自带的 OTAEngine 解析文本、生成 OTA 导入 Excel（升级规则 + 升级提示语两表）。
  步骤2 校验：对生成的工作簿做严格校验（表名/表头/每行升级文件地址与 OTA 类型等规则）。
  步骤3 上传：从源数据抽取 FTP .zip 地址，下载后上传到 S3（download_and_upload）。

用法：
    cd ota_cms_automation
    # 从 stdin 粘贴文本（推荐）：
    cat 文本.txt | ./.venv/bin/python -m automation.prep --device-code MT9655AMAndrGTV
    # 或从文件读：
    ./.venv/bin/python -m automation.prep --file 文本.txt --device-code MT9655AMAndrGTV
    # 只生成+校验，不碰 FTP/S3（打印 S3 计划与 ftp zip 清单）：
    ... | ./.venv/bin/python -m automation.prep --no-upload --device-code MT9655AMAndrGTV

文本来源：给了 --file 读该文件，否则读 stdin。CLI 的 --device-code/--device-name/
--feature-code/--mac-group 非空时覆盖从文本解析出的设备信息。

真实上传前置条件（仅在不带 --no-upload 时需要）：
  - FTP：config/.env 里的 OTA_FTP_PASS（及可选 OTA_FTP_USER）非占位值。
  - AWS：区域/密钥不在 .env 里配置，走 boto3 默认链（~/.aws），region=us-west。
『真实 S3 上传 / FTP 下载』是不可逆的外部 I/O；要安全演练请始终带 --no-upload。

返回码：
    0  成功（生成+校验通过；上传模式下全部上传成功，或 --no-upload 正常结束）
    1  业务失败（生成报错 / 校验不通过 / 无可上传地址 / 上传有失败项）
    2  参数或安全失败（未读到文本 / --file 不存在）
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import tempfile

from openpyxl import load_workbook

from . import config as cms_config  # noqa: F401  触发 ota_cms 的 config/.env 加载
from .otagen.engine import OTAEngine
from .otagen import transfer
from .otagen.legacy_engine import OUTPUT_HEADERS
# 终端着色（绿=成功/一致，黄=警告，红=错误）抽到共享模块 automation/_log.py；
# 用别名保留本文件原有调用名 _ok/_warn/_err/_c，其余调用点零改动。
from ._log import ok as _ok, warn as _warn, err as _err, color as _c


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="deployota prep",
        description="OTA 模板 prep CLI：生成 Excel → 严格校验 → 场测 MAC 组 → 上传 S3。",
    )
    p.add_argument("--file", dest="file", default=None,
                   help="输入文件：.txt（升级信息文本）或 .xlsx/.xlsm（OTA 申请模板 Excel，直接解析；旧版 .xls 需先另存为 .xlsx）；"
                        "省略则从 stdin 读文本。")
    p.add_argument("--device-code", dest="device_code", default="",
                   help="设备类型代码（方案号）；非空时覆盖文本解析值。")
    p.add_argument("--device-name", dest="device_name", default="",
                   help="设备类型名称；非空时覆盖文本解析值。")
    p.add_argument("--feature-code", dest="feature_code", default="",
                   help="特征码（24 位数字，可留空）；非空时覆盖文本解析值。")
    p.add_argument("--mac-group", dest="mac_group", default="",
                   help="MAC 组；非空时覆盖文本解析值。")
    p.add_argument("--out-dir", dest="out_dir", default=None,
                   help="工作目录（生成 Excel 与去重索引的存放处）。省略时：给了 --file 则用该文件所在目录"
                        "（相对路径也会先转成绝对路径），否则用 OTA_WORKSPACE。")
    p.add_argument("--no-upload", dest="no_upload", action="store_true", default=False,
                   help="只生成+校验并打印 S3 计划与 ftp zip 清单，绝不触碰 FTP/S3。")
    p.add_argument("--yes", "-y", dest="assume_yes", action="store_true", default=False,
                   help="设备信息与映射表不一致时自动确认继续（非交互/管道场景用）；"
                        "默认不一致会等待人工确认，无法交互时退出。")
    p.add_argument("--force", dest="force", action="store_true", default=False,
                   help="即使输出目录已存在相同内容的 Excel 也强制重新生成（默认复用已有）。")
    p.add_argument("--originator", dest="originator", default=None,
                   help="发起人姓名（这次 OTA 请求是谁提的）。不给时在终端询问，可直接回车跳过；"
                        "会显示在 dass init 后的部署汇总和邮件里，并写入部署库。")
    p.add_argument("--default-messages", dest="default_messages", action="store_true", default=False,
                   help="源文件没有提示语、确实要用内置默认文案时加（否则最终检查发现是默认文案会判失败）。")
    p.add_argument("--no-mac", dest="no_mac", action="store_true", default=False,
                   help="跳过 CMS『测试<设备类型代码>』MAC 组检查。")
    p.add_argument("--mac-dry-run", dest="mac_dry_run", action="store_true", default=False,
                   help="MAC 组测试模式：在 CMS 里填好后点『取消』，不保存（随后继续上传）。")
    return p


def _step_mac_group(args, sj, device_code: str, mac_group: str, raw_text: str = "",
                    out: dict | None = None) -> int:
    """步骤 2b：场测 MAC 组检查/加入。返回 0 = 继续；非 0 = 退出码（不上传）。

    raw_text：整份输入原文（文本输入 = 文本；Excel 输入 = 所有单元格文字），用来找 MAC 列之外写的 MAC。
    """
    from .mac_group import group_name, scan_rows, scan_text_macs
    print(_c("—— 步骤2b：场测 MAC 组 ——", "bold"))
    if args.no_mac:
        _no_mac_banner()
        return 0
    rows = sj.get("rows", []) if isinstance(sj, dict) else []
    scanned = scan_rows(rows)

    # 逐行列出全部 MAC（原文 → 后 8 位），看不懂的片段单独标出。
    print("  逐行场测 MAC（原文 → 后 8 位）：")
    for r in scanned:
        tail = ", ".join(r["macs"]) or "（无）"
        raw = " ".join(r["raw"].split()) or "（空）"
        line = f"    第{r['row']}行 {r['model'] or '?'}：{raw} → {tail}"
        print(_err(line + f"  ← 看不懂：{', '.join(r['bad'])}") if r["bad"] else line)
    macs: list[str] = []
    for r in scanned:
        for m in r["macs"]:
            if m not in macs:
                macs.append(m)
    extra = [m for m in scan_text_macs(raw_text) if m not in macs]
    if extra:
        print(_warn(f"MAC 列之外还写了 {len(extra)} 个 MAC，一并加入检查：{', '.join(extra)}"))
        macs += extra
    print(f"  全部场测 MAC（{len(macs)} 个）：{', '.join(macs) or '（无）'}")
    if out is not None:
        out["macs"] = list(macs)

    bad = [r for r in scanned if r["bad"]]
    if bad:
        _mac_failed_banner(group_name(device_code), "以下行的 MAC 写法看不懂（要么是完整 12 位 MAC 如 cc:12:28:cd:db:d8，要么是后 8 位如 b7fc8ce6）："
                           + "；".join(f"第{r['row']}行 {r['model']}『{', '.join(r['bad'])}』" for r in bad))
        return 1
    if not macs:
        _mac_failed_banner(group_name(device_code), "输入里没有找到任何场测 MAC 地址"
                           "（确实不需要 MAC 组，请加 --no-mac）")
        return 1
    without = [r for r in scanned if not r["macs"]]
    if without and len(without) < len(scanned):
        print(_warn(f"有 {len(without)} 行没写 MAC（"
                    + "、".join(f"第{r['row']}行 {r['model']}" for r in without)
                    + "）：这些机型用上面的 MAC 一起检查，请确认。"))
    if not device_code:
        _mac_failed_banner("测试<设备类型代码>", "没有设备类型代码，无法检查 MAC 组（可加 --device-code）")
        return 1
    want = group_name(device_code)
    if mac_group and mac_group.strip() != want:
        print(_warn(f"映射表/输入里的 MAC 组是『{mac_group}』，这里检查的是『{want}』。"))
    from playwright.sync_api import sync_playwright
    from .cms_mac import ensure_macs
    with sync_playwright() as p:
        res = ensure_macs(p, device_code, macs, assume_yes=args.assume_yes,
                          dry_run=args.mac_dry_run, read_only=args.no_upload)
    if out is not None:
        out["res"] = res   # 最终检查直接用这次结果，不再重复登录查 CMS
    st = res["status"]
    if st in ("ok", "added", "created", "tested"):
        print(_ok(f"MAC 组『{res['group']}』：{res['detail']}"))
        return 0
    if st in ("would_add", "would_create"):
        print(_warn(f"MAC 组『{res['group']}』：{res['detail']}；去掉 --no-upload 正式运行时会处理。"))
        return 0
    _mac_failed_banner(res["group"], res["detail"])
    return 1


def _ask_originator(args, output_path: str) -> str:
    """发起人姓名。--originator 给了就用；否则终端可交互时询问，回车沿用之前记的（复用的 Excel）或留空。"""
    from . import last_prep
    if args.originator is not None:
        return args.originator.strip()
    old = str(last_prep.load_meta(output_path).get("originator", "")).strip()
    if not sys.stdin.isatty():
        return old
    hint = f"回车沿用『{old}』" if old else "回车跳过"
    try:
        got = input(f"发起人姓名 / Originator（{hint}）：").strip()
    except EOFError:
        got = ""
    return got or old


def _no_mac_banner(where: str = "") -> None:
    """--no-mac 的红色提醒（开头、步骤2b、结尾都打）。"""
    bar = "!" * 64
    print(_c(bar, "red"))
    print(_c(f"!! --no-mac：本次 prep 未检查场测 MAC 组{where}。请确认场测 MAC 已在『测试+设备类型代码』组里。", "red"))
    print(_c(bar, "red"))


def _mac_failed_banner(group: str, detail: str) -> None:
    """MAC 组失败的醒目提醒（stderr）。"""
    bar = "=" * 64
    lines = [
        bar,
        f"✗✗ MAC 组未完成：{group}",
        f"   原因：{detail}",
        "   prep 已停止：未上传 S3，这份 Excel 也没有记为『最近 prep』",
        "   （publish / online / dass 不会默认使用它）。",
        "   处理好后重跑同一条 prep 命令（已生成的 Excel 会复用）；确定不需要 MAC 组可加 --no-mac。",
        bar,
    ]
    for ln in lines:
        print(_c(ln, "red"), file=sys.stderr)


def _validate_output_row_like(row: dict) -> list[str]:
    """单行规则校验（运行时与 _selftest 共用的唯一行规则）。返回问题串列表（空=通过）。

    规则：
      - 升级文件地址：非空；re.match(r'^https?://', v, re.I)；不含 3+ 连续斜杠
        （re.search(r'/{3,}', v)，用正则而非字面 '///'）；以 .zip 结尾（忽略大小写）。
      - OTA类型：str(v).strip() 属于 ('1','2','11')。
    """
    problems: list[str] = []

    url = str(row.get("升级文件地址", "") or "").strip()
    if not url:
        problems.append("升级文件地址为空")
    else:
        if not re.match(r"^https?://", url, re.I):
            problems.append(f"升级文件地址非 http(s)://：{url!r}")
        if re.search(r"/{3,}", url):
            problems.append(f"升级文件地址含 3+ 连续斜杠：{url!r}")
        if not url.lower().endswith(".zip"):
            problems.append(f"升级文件地址未以 .zip 结尾：{url!r}")

    ota_type = str(row.get("OTA类型", "") or "").strip()
    if ota_type not in ("1", "2", "11"):
        problems.append(f"OTA类型非法（应为 1/2/11）：{ota_type!r}")

    return problems


def _validate_excel(path: str) -> list[str]:
    """严格校验生成的工作簿，收集所有问题并返回（空列表=通过）。"""
    problems: list[str] = []
    wb = load_workbook(path, read_only=True)
    try:
        if wb.sheetnames != ["升级规则", "升级提示语"]:
            problems.append(f"表名应为 ['升级规则','升级提示语']，实际：{wb.sheetnames}")
            return problems

        ws = wb["升级规则"]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            problems.append("升级规则表为空（无任何行）。")
            return problems

        header = [("" if c is None else str(c)) for c in rows[0]]
        if header != list(OUTPUT_HEADERS):
            problems.append(f"升级规则表头与 OUTPUT_HEADERS 不一致，实际：{header}")
            return problems

        data_rows = rows[1:]
        if len(data_rows) < 1:
            problems.append("升级规则表无数据行（至少需要 1 行）。")
            return problems

        for idx, raw in enumerate(data_rows, start=2):
            rowdict = {h: ("" if v is None else str(v)) for h, v in zip(OUTPUT_HEADERS, raw)}

            for prob in _validate_output_row_like(rowdict):
                problems.append(f"第{idx}行：{prob}")

            device_code = str(rowdict.get("设备类型代码", "") or "").strip()
            if not device_code:
                problems.append(f"第{idx}行：设备类型代码为空")
            elif re.search(r"\s", device_code):
                problems.append(f"第{idx}行：设备类型代码含空白：{device_code!r}")
            elif not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_\-().]*", device_code):
                problems.append(f"第{idx}行：设备类型代码含非法字符：{device_code!r}")

            feature_code = str(rowdict.get("特征码", "") or "").strip()
            if feature_code and not re.fullmatch(r"\d{24}", feature_code):
                problems.append(f"第{idx}行：特征码非 24 位数字：{feature_code!r}")

            if not str(rowdict.get("MAC组", "") or "").strip():
                problems.append(f"第{idx}行：MAC组为空")

            internal = str(rowdict.get("内部机型信息", "") or "").strip()
            if not internal:
                problems.append(f"第{idx}行：内部机型信息为空")

            ext = str(rowdict.get("设备扩展信息", "") or "").strip()
            if not ext:
                problems.append(f"第{idx}行：设备扩展信息为空")
            else:
                if ext == internal:
                    problems.append(f"第{idx}行：设备扩展信息与内部机型信息相同：{ext!r}")
                if ext.isdigit():
                    problems.append(f"第{idx}行：设备扩展信息不应全为数字：{ext!r}")

        return problems
    finally:
        wb.close()


_DEVICE_COL_LABELS = ("设备类型代码", "设备类型代码(方案号)", "featurecode", "特征码", "featurecode(特征码)")


# 版本号写法：4 位数字 + 1~3 个字母 + 4 位数字（0100UQ1104、0100AWP0224），或 V1.2.3 这种。
# 复制来的引擎只认 1~2 个字母，0100AWP0224（MSD6886）认不出 → 整行按内容『重新对齐』时把列放错
# （源版本 / 目标版本变空，机型信息填成源版本）。这里替换引擎的版本判断，不改引擎文件本身。
_VERSION_RE = re.compile(r"^(?:\d{4}[A-Z]{1,3}\d{4}|V\d+(?:\.\d+)+)$", re.IGNORECASE)


# 引擎语种别名表里没有的写法（『X文』『繁體』…）。括号里的说明（如『英文（言語設定＝日本語以外）』）先去掉再认。
EXTRA_LANGUAGE_ALIASES = {
    "日文": "日语", "日本语": "日语", "日語": "日语",
    "韩文": "韩语", "韓文": "韩语", "韩国语": "韩语",
    "法文": "法语", "德文": "德语", "俄文": "俄语", "泰文": "泰语",
    "西文": "西班牙语", "西班牙文": "西班牙语", "葡文": "葡萄牙语", "葡萄牙文": "葡萄牙语",
    "意大利文": "意大利语", "意文": "意大利语", "荷兰文": "荷兰语", "波兰文": "波兰语",
    "阿拉伯文": "阿拉伯语", "印尼文": "印尼语", "马来文": "马来语", "越南文": "越南语",
    "繁体": "繁体中文", "繁體": "繁体中文", "繁體中文": "繁体中文", "繁中": "繁体中文",
    "简体": "简体中文", "簡體中文": "简体中文", "中文简体": "简体中文",
}
_PAREN_RE = re.compile(r"[（(][^（）()]*[)）]")


def resolve_language(eng, label) -> str:
    """语种标签 → 标准语种名（认不出返回空串）。先去掉括号说明，再查补充别名，最后交给引擎。"""
    raw = str(label or "").strip()
    if not raw or len(raw) > 40:
        return ""
    core = _PAREN_RE.sub("", raw).strip(" :：/")
    for cand in (core, raw):
        if cand in EXTRA_LANGUAGE_ALIASES:
            return EXTRA_LANGUAGE_ALIASES[cand]
    orig = getattr(eng.app, "_orig_normalize_message_language", eng.app.normalize_message_language)
    return orig(core) or orig(raw) or ""


def patch_engine(eng) -> None:
    """修正引擎：版本号识别（见 _VERSION_RE）、语种别名（见 EXTRA_LANGUAGE_ALIASES）。对 OTAEngine 实例调用一次即可。"""
    eng.app.looks_like_version = lambda value: bool(_VERSION_RE.match(str(value or "").strip()))
    if not hasattr(eng.app, "_orig_normalize_message_language"):
        eng.app._orig_normalize_message_language = eng.app.normalize_message_language
        eng.app.normalize_message_language = lambda language: resolve_language(eng, language)


def columnar_messages_to_rows(eng, text: str) -> tuple[str, list[str], str]:
    """横排提示语 → 引擎认的竖排『语种<Tab>描述』。

    横排：一行全是语种（Tab 分隔，如『日文<Tab>英文（言語設定＝日本語以外）』），下面每列一段内容，
    单元格内可以换行，列与列之间只有一个 Tab。条件：表头每格都是语种（≥2 个），后面内容里 Tab 个数正好 = 语种数 − 1。
    返回 (新文本, 认出的语种, 问题说明)。没有横排表头 → 原样返回、问题为空；有表头但对不上 → 原样返回、给出问题。
    """
    lines = str(text or "").split("\n")
    for i, ln in enumerate(lines):
        if "\t" not in ln:
            continue
        cells = [c.strip() for c in ln.split("\t")]
        if len(cells) < 2 or any(not c for c in cells):
            continue
        langs = [resolve_language(eng, c) for c in cells]
        if not all(langs):
            continue
        body = "\n".join(lines[i + 1:]).strip("\n")
        if body.count("\t") != len(cells) - 1:
            return text, langs, (f"提示语是横排的（{' / '.join(cells)}），但下面内容里有 {body.count(chr(9))} 个 Tab，"
                                 f"应为 {len(cells) - 1} 个，没法按列拆开")
        descs = []
        for part in body.split("\t"):
            part = part.strip()
            if len(part) >= 2 and part[0] == '"' and part[-1] == '"':      # Excel 复制多行单元格会加引号
                part = part[1:-1].replace('""', '"')
            descs.append(part.strip())
        if any(not d for d in descs):
            return text, langs, "横排提示语里有空的一列"
        block = ["语种\t描述"] + [f"{lg}\t{d}" for lg, d in zip(langs, descs)]
        return "\n".join(lines[:i] + block), langs, ""
    return text, [], ""


def lift_device_columns(eng, text: str) -> tuple[str, dict | None]:
    """竖排粘贴（每格一行）里，若表头最前面是『设备类型代码（方案号）』/『featurecode』这两列，
    把它们改写成表格前的『标签\n值』形式（引擎支持的写法），并从每行数据里去掉这几列的值。

    返回 (新文本, 说明或 None)。不是这种排法时原样返回 (text, None)。
    要求：同一份里设备类型代码 / featurecode 只有一个值（一个 Excel 对应一个设备类型，本来就是约定）。
    """
    norm = eng.app.normalize_header
    lines = str(text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    nb = [k for k, ln in enumerate(lines) if ln.strip()]
    if not nb:
        return text, None
    # 1) 最前面连续的『设备类型代码 / featurecode』表头
    lead = []
    for k in nb:
        if norm(lines[k]) in _DEVICE_COL_LABELS:
            lead.append(k)
        else:
            break
    if not lead or len(lead) >= len(nb):
        return text, None
    after = nb[len(lead)]
    # 紧跟着的必须是表格表头（机型等）；否则是『标签\n值』写法，引擎本来就支持，不动。
    if not eng._is_vertical_header_line(lines[after]):
        return text, None
    labels = [lines[k].strip() for k in lead]
    keys = [norm(lines[k]) for k in lead]
    rest = nb[len(lead):]
    # 2) 找第一行数据的起点：有 featurecode 列时用『feature code 值』定位，否则用映射表里的设备类型代码定位
    start = -1
    fi = next((i for i, k in enumerate(keys) if k.startswith("featurecode") or k == "特征码"), None)
    if fi is not None:
        p = next((i for i, k in enumerate(rest) if eng.app.looks_like_feature_code(lines[k].strip())), -1)
        start = p - fi if p >= 0 else -1
    else:
        known = {str(r.get("device_code", "")).lower() for r in eng.app.load_device_history_seed()}
        start = next((i for i, k in enumerate(rest) if lines[k].strip().lower() in known), -1)
    if start < 1 or start + len(lead) > len(rest):
        return text, None
    values = [lines[rest[start + i]].strip() for i in range(len(lead))]
    if any(not v or eng._is_vertical_header_line(v) for v in values):
        return text, None
    # 3) 去掉每行开头重复出现的这几列的值（逐行按『连续的同一组值』匹配）
    drop: set[int] = set()
    rows = 0
    i = start
    while i + len(lead) <= len(rest):
        if [lines[rest[i + j]].strip() for j in range(len(lead))] == values:
            drop.update(rest[i + j] for j in range(len(lead)))
            rows += 1
            i += len(lead)
        else:
            i += 1
    drop.update(lead)
    meta = []
    for lab, val in zip(labels, values):
        meta += [lab, val]
    new_lines = meta + [ln for k, ln in enumerate(lines) if k not in drop]
    return "\n".join(new_lines), {"labels": labels, "values": values, "rows": rows}


def normalize_message_lines(text: str) -> tuple[str, list[str]]:
    """把『英语 A Test version ...』这种『语种 + 空格 + 内容』的提示语行改成引擎认的『英语：A Test version ...』。

    只改不含 Tab、以已知语种名开头、语种名后是空白且后面还有内容的行。返回 (新文本, 改到的语种列表)。
    """
    from .otagen.legacy_engine import MESSAGE_LANGUAGE_NAMES
    names = sorted(MESSAGE_LANGUAGE_NAMES, key=len, reverse=True)
    pat = re.compile(r"^\s*(" + "|".join(re.escape(n) for n in names) + r")[ \u3000]+(\S.*)$", re.IGNORECASE)
    out, hit = [], []
    for ln in str(text or "").split("\n"):
        m = pat.match(ln) if "\t" not in ln else None
        if m:
            out.append(f"{m.group(1)}：{m.group(2).rstrip()}")
            hit.append(m.group(1))
        else:
            out.append(ln)
    return "\n".join(out), hit


def move_messages_before_table(eng, text: str) -> tuple[str, int]:
    """竖排粘贴里，表格后面的『语种：内容』提示语会被引擎并进最后一行数据然后截掉。
    把从第一条提示语开始到末尾的内容挪到表格表头前面（引擎在那里能正确识别）。

    只处理竖排（不含 Tab）的文本；不是竖排、或表格后没有提示语时原样返回。返回 (新文本, 挪动的行数)。
    """
    t = str(text or "")
    if "\t" in t or eng.reshape_vertical_paste(t) == t:
        return t, 0
    from .otagen.legacy_engine import MESSAGE_LANGUAGE_NAMES
    names = sorted(MESSAGE_LANGUAGE_NAMES, key=len, reverse=True)
    label = re.compile(r"^\s*(" + "|".join(re.escape(n) for n in names) + r")\s*[:：=]", re.IGNORECASE)
    lines = t.split("\n")
    hdr = next((i for i, ln in enumerate(lines) if ln.strip() and eng._is_vertical_header_line(ln)), -1)
    if hdr < 0:
        return t, 0
    first = next((i for i in range(hdr + 1, len(lines)) if label.match(lines[i])), -1)
    if first < 0:
        return t, 0
    block = [ln for ln in lines[first:]]
    while block and not block[-1].strip():
        block.pop()
    new = lines[:hdr] + block + lines[hdr:first]
    return "\n".join(new), len([b for b in block if b.strip()])


def _canonical_device_code(eng, code: str):
    """把设备码规范成映射表里的写法（以映射表为标准，大小写不敏感匹配）。

    返回 (规范后的设备码, 是否被改写)。先精确匹配；否则仅当大小写不敏感匹配【唯一】时才采用；
    没有匹配或不唯一则原样返回（交给后面的映射表校验去报“不在映射表中”）。
    """
    code = str(code or "").strip()
    if not code:
        return code, False
    table = [str(r.get("device_code", "")).strip() for r in eng.app.load_device_history_seed()]
    if code in table:
        return code, False
    ci = [t for t in table if t and t.lower() == code.lower()]
    if len(ci) == 1:
        return ci[0], True
    return code, False


def _check_device_against_mapping(eng, text, device_code, feature_code, device_name, mac_group,
                                  pasted_override=None):
    """把文本解析出的 设备类型代码/feature code 与映射表(seed)比对。

    返回 (ok, lines)：ok 为 True 表示一致（code 在表中且文本 feature==表中 feature，
    文本未显式给 feature 时只校验 code 存在）；lines 为要打印的日志行。
    一致时 lines 打印 设备类型代码/设备类型名称/feature code/测试MAC组 四项；
    不一致时 lines 说明具体差异（code 不在表 / feature 冲突 / code 缺失）。
    """
    lines = []
    # 重新抽取文本里【原始】解析值（normalize 竖排后），用于与表比对。
    if pasted_override is not None:
        # Excel 输入：设备码/feature 来自表格列（调用方已读好），不再从文本抽取。
        pasted = pasted_override
    else:
        norm = eng.reshape_vertical_paste(text) if str(text).strip() else text
        pasted = eng.app.extract_form_fields_from_text(norm) if str(norm).strip() else {}
    text_code = str(pasted.get("device_code", "") or "").strip()
    text_feat = str(pasted.get("feature_code", "") or "").strip()

    # —— 先打印从文本解析出的内容（透明可核对）——
    lines.append(_c("—— 从文本解析出的设备信息 ——", "bold"))
    lines.append(f"  设备类型代码（解析）：{text_code or _c('(未解析到)', 'yellow')}")
    lines.append(f"  feature code（解析）：{text_feat or _c('(未解析到)', 'yellow')}")

    code = str(device_code or "").strip()
    if not code:
        lines.append(_err("设备校验：未能确定设备类型代码（文本未解析到，且未用 --device-code 指定）。"))
        return False, lines

    preset = eng.device_preset(code)
    in_table = bool(preset.get("feature_code") or preset.get("device_name")) or \
        any(r.get("device_code") == code for r in eng.app.load_device_history_seed())

    if not in_table:
        lines.append(_err(f"设备校验：设备类型代码 {code!r} 不在映射表中（共 "
                          f"{len(eng.app.load_device_history_seed())} 个已知设备）。"))
        return False, lines

    table_feat = str(preset.get("feature_code", "") or "").strip()
    # feature 一致性：仅当文本【显式给了】feature 时才强比；否则以表为准（黄色提示）。
    if text_feat and table_feat and text_feat != table_feat:
        lines.append(_err(f"设备校验：feature code 不一致 —— 文本={text_feat!r}，"
                          f"映射表={table_feat!r}（设备码 {code!r}）。"))
        return False, lines
    if not text_feat:
        lines.append(_warn(f"文本未给出 feature code，将以映射表值为准：{table_feat or '(映射表也为空)'}"))

    # 一致：绿色打印四项（名称/feature/MAC 以映射表为权威来源）。
    name = preset.get("device_name", "") or ""
    mac = preset.get("mac_group", "") or ("测试" + code)
    lines.append(_ok("设备信息与映射表一致："))
    lines.append(_c(f"  设备类型代码：{code}", "green"))
    lines.append(_c(f"  设备类型名称：{name or '(映射表未填)'}", "green" if name else "yellow"))
    lines.append(_c(f"  feature code：{table_feat or '(映射表未填)'}", "green" if table_feat else "yellow"))
    lines.append(_c(f"  测试MAC组：{mac}", "green"))
    return True, lines


def _device_in_table(eng, code: str) -> bool:
    """设备类型代码是否已在映射表里（精确匹配）。"""
    code = str(code or "").strip()
    if not code:
        return False
    return any(str(r.get("device_code", "")).strip() == code
               for r in eng.app.load_device_history_seed())


def _save_device_to_mapping(eng, code, name, feature, mac) -> str:
    """把新设备追加进映射表（device_history_seed.json），并刷新引擎缓存；返回文件路径。

    已存在同代码（忽略大小写）时抛 ValueError，避免重复/覆盖。
    """
    import json
    path = eng.app.get_resource_path("device_history_seed.json")
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except FileNotFoundError:
        data = []
    if any(str(r.get("device_code", "")).lower() == code.lower() for r in data):
        raise ValueError(f"映射表已有设备类型代码 {code!r}（忽略大小写），不重复写入。")
    data.append({"device_code": code, "device_name": name,
                 "feature_code": feature, "mac_group": mac})
    with open(path, "w", encoding="utf-8") as f:
        f.write(json.dumps(data, ensure_ascii=False, indent=2) + "\n")
    if hasattr(eng.app, "_device_history_seed"):
        del eng.app._device_history_seed  # 下次查询重新读文件
    return path


def _prompt_new_device(eng, code, feature, name, mac, ask=input):
    """设备类型代码不在映射表时，交互录入 4 项：设备类型代码 / feature code / 设备类型名称 / MAC 组。

    括号里是默认值（来自文件解析或命令行），直接回车采用默认；输入 q 取消。
    返回 (code, feature, name, mac, saved_path_or_None)；取消返回 None。
    feature code 必填（纯数字，通常 24 位，位数不符会二次确认）；名称可留空；MAC 组默认『测试+设备类型代码』。
    """
    def read(label, default="", required=False, validator=None):
        while True:
            shown = f"［默认 {default}］" if default else ""
            ans = ask(f"  {label}{shown}（q 取消）：").strip()
            if ans.lower() == "q":
                return None
            val = ans or default
            if required and not val:
                print(_warn(f"  {label}不能为空。"))
                continue
            if validator:
                err = validator(val)
                if err:
                    print(_warn("  " + err))
                    continue
            return val

    def v_code(v):
        return None if re.fullmatch(r"[A-Za-z0-9_\-]+", v) else "设备类型代码只能含字母、数字、下划线、连字符。"

    def v_feat(v):
        return None if v.isdigit() else "feature code 应为纯数字。"

    print(_c("—— 映射表里没有这个设备类型代码，请录入设备信息 ——", "bold"))
    c = read("设备类型代码", code, True, v_code)
    if c is None:
        return None
    # 录入的代码若与映射表只差大小写，直接提示并采用映射表写法，避免重复条目。
    canon, changed = _canonical_device_code(eng, c)
    if changed:
        print(_warn(f"  映射表已有 {canon}（仅大小写不同），请直接用它，无需新增。"))
        return None
    f = read("feature code", feature, True, v_feat)
    if f is None:
        return None
    if len(f) != 24:
        a = ask(f"  ⚠ feature code 长度为 {len(f)}（映射表里通常是 24 位），确认无误？[y 确认，其它重输]：").strip().lower()
        if a not in ("y", "yes"):
            f = read("feature code", "", True, v_feat)
            if f is None:
                return None
    n = read("设备类型名称（可留空）", name)
    if n is None:
        return None
    m = read("MAC 组", mac or ("测试" + c), True)
    if m is None:
        return None
    print(_c("  请确认：", "bold"))
    print(f"    设备类型代码：{c}\n    feature code：{f}\n    设备类型名称：{n or '(空)'}\n    MAC 组：{m}")
    a = ask("  保存到映射表以后自动识别？[y 保存并继续 / n 仅本次使用 / 其它取消]：").strip().lower()
    if a in ("y", "yes"):
        try:
            path = _save_device_to_mapping(eng, c, n, f, m)
            print(_ok(f"已保存到映射表：{path}"))
            return c, f, n, m, path
        except (ValueError, OSError) as e:
            print(_err(f"保存映射表失败：{e}"), file=sys.stderr)
            return None
    if a in ("n", "no"):
        return c, f, n, m, None
    return None


def _excel_all_text(path: str) -> str:
    """Excel 所有 sheet 所有单元格的文字（找 MAC 列之外写的 MAC 用）。读不了返回空串。"""
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return ""
    parts = []
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                parts.extend(str(v) for v in row if v is not None)
    finally:
        wb.close()
    return "\n".join(parts)


def main() -> int:
    """--no-mac 时在开头和结尾各打一次红色提醒；其余见 _main。"""
    no_mac = "--no-mac" in sys.argv[1:]
    if no_mac:
        _no_mac_banner()
    try:
        return _main()
    finally:
        if no_mac:
            _no_mac_banner("（结束）")


def _main() -> int:
    args = _build_parser().parse_args()

    # 相对路径按「调用者所在目录」解析（deployota 会先 cd 到项目目录，不处理就会找不到文件）。
    caller_cwd = os.environ.get("OTA_CALLER_CWD")
    if caller_cwd:
        if args.file:
            args.file = os.path.join(caller_cwd, os.path.expanduser(args.file))
        if args.out_dir:
            args.out_dir = os.path.join(caller_cwd, os.path.expanduser(args.out_dir))

    # (1) 读取输入来源：.xlsx/.xlsm 直接按 Excel 解析；其余按文本（--file txt 或 stdin）。
    excel_input = False
    pasted_override = None
    if args.file:
        if not os.path.isfile(args.file):
            print(f"✗ 找不到 --file 指定的文件：{args.file}", file=sys.stderr)
            return 2
        base = os.path.basename(args.file)
        if base.startswith("~$"):
            print(_err(f"{base} 是 Excel 打开时产生的临时锁文件，不是真正的表格；请指定原文件。"),
                  file=sys.stderr)
            return 2
        if args.file.lower().endswith((".xls", ".xlsb")):
            print(_err(f"不支持旧版 Excel 格式（{os.path.splitext(args.file)[1]}）。"
                       "请在 Excel 里另存为 .xlsx 后再用 --file 传入。"), file=sys.stderr)
            return 2
        excel_input = _is_excel_input(args.file)
        if excel_input:
            text = ""  # Excel 输入没有文本；解析/指纹都走文件本身。
        else:
            with open(args.file, "r", encoding="utf-8") as f:
                text = f.read()
    else:
        text = sys.stdin.read()

    if not excel_input and text.strip() == "":
        print("✗ 未从 stdin/--file 读到任何文本", file=sys.stderr)
        return 2

    eng = OTAEngine()
    patch_engine(eng)

    if not excel_input:
        # 竖排粘贴里『设备类型代码 / featurecode』作为表格前两列时，引擎认不出（它只认这两项写在表格外）。
        # 先把这两列挪成表格前的『标签 / 值』，再交给引擎。
        text, lifted = lift_device_columns(eng, text)
        if lifted:
            print(_warn("竖排表格的前几列是 " + " / ".join(lifted["labels"])
                        + f"：已移到表格外（{', '.join(lifted['values'])}），从 {lifted['rows']} 行数据里去掉了这几列。"))
        text, col_langs, col_problem = columnar_messages_to_rows(eng, text)
        if col_langs and not col_problem:
            print(_warn(f"提示语是横排的（每个语种一列）：已转成竖排，语种：{', '.join(col_langs)}"))
        elif col_problem:
            print(_err(f"提示语没法解析：{col_problem}"), file=sys.stderr)
        text, langs = normalize_message_lines(text)
        if langs:
            print(_warn(f"提示语『语种 + 空格 + 内容』写法已识别：{', '.join(langs)}"))
        text, moved = move_messages_before_table(eng, text)
        if moved:
            print(_warn(f"竖排表格后面的提示语（{moved} 行）已移到表格前，避免被并进最后一行数据。"))

    if excel_input:
        print(_c("—— 输入：Excel 文件（直接解析）——", "bold"))
        print(f"  {os.path.abspath(args.file)}")
        # 设备码 / feature code 是表格里的列，引擎不会抽取，这里自己读。
        ex = _read_excel_device_fields(args.file)
        codes, feats = ex["device_codes"], ex["feature_codes"]
        if len(codes) > 1 and not args.device_code.strip():
            print(_err(f"Excel 内含多个设备类型代码 {codes}，无法确定唯一值。"
                       "请拆分文件，或用 --device-code 显式指定。"), file=sys.stderr)
            return 2
        if len(feats) > 1 and not args.feature_code.strip():
            print(_err(f"Excel 内含多个 Feature Code {feats}，无法确定唯一值。"
                       "请拆分文件，或用 --feature-code 显式指定。"), file=sys.stderr)
            return 2
        x_code = codes[0] if codes else ""
        x_feat = feats[0] if feats else ""
        pasted_override = {"device_code": x_code, "feature_code": x_feat}
        device_code = args.device_code.strip() or x_code
        feature_code = args.feature_code.strip() or x_feat
        preset = eng.device_preset(device_code) if device_code else {}
        device_name = args.device_name.strip() or preset.get("device_name", "")
        feature_code = feature_code or preset.get("feature_code", "")
        mac_group = args.mac_group.strip() or preset.get("mac_group", "")
    else:
        # 设备信息：先从文本解析，再用 CLI 非空值覆盖。
        info = eng.resolve_device_info_from_text(text)
        device_code = args.device_code.strip() or info.get("device_code", "")
        device_name = args.device_name.strip() or info.get("device_name", "")
        feature_code = args.feature_code.strip() or info.get("feature_code", "")
        mac_group = args.mac_group.strip() or info.get("mac_group", "")

    # 设备码以映射表写法为准（仅大小写不同时规范化），文件名 / S3 路径 / 生成的 Excel 都用规范写法。
    canon, changed = _canonical_device_code(eng, device_code)
    if changed:
        print(_warn(f"设备类型代码已按映射表写法规范化：{device_code} → {canon}（仅大小写不同）"))
        device_code = canon
        # 名称 / MAC 组 / feature code 若是按原始写法查不到而为空，则按规范后的设备码补齐。
        _p = eng.device_preset(device_code)
        device_name = args.device_name.strip() or _p.get("device_name", "") or device_name
        mac_group = args.mac_group.strip() or _p.get("mac_group", "") or mac_group
        feature_code = args.feature_code.strip() or feature_code or _p.get("feature_code", "")

    # 若设备码经 CLI 覆盖后才确定、而 MAC 组仍为空，则按设备预设补齐
    # （预设默认 '测试'+设备类型代码，与引擎 resolve 的建议值一致）。
    if not mac_group and device_code:
        mac_group = eng.device_preset(device_code).get("mac_group", "")

    manual_unsaved = False
    # (1a) 设备类型代码不在映射表（或没解析到）且是交互终端 → 引导录入 4 项并可保存进映射表。
    #      非交互 / --yes 保持原行为（报不一致，--yes 才继续）。
    if not args.assume_yes and sys.stdin.isatty() and not _device_in_table(eng, device_code):
        got = _prompt_new_device(eng, device_code, feature_code, device_name, mac_group)
        if got is None:
            print("已取消：未录入设备信息，未生成。", file=sys.stderr)
            return 2
        device_code, feature_code, device_name, mac_group, _saved = got
        manual_unsaved = _saved is None  # 仅本次使用：不在表里，跳过下面的映射表一致性校验
        if pasted_override is not None:
            pasted_override = {"device_code": device_code, "feature_code": feature_code}

    # (1b) 设备信息与映射表一致性校验（生成前）。
    #   从文本解析出的 设备类型代码/feature code 与 mapping table(seed) 比对：
    #   一致 -> 打印 4 项(代码/名称/feature/测试MAC组) 并继续；
    #   不一致 -> 打印差异，等待人工确认(yes/y 继续，其它退出)；--yes 自动继续；无法交互则退出。
    if manual_unsaved:
        ok, lines = True, [_warn("设备信息为本次手工录入（未保存到映射表），跳过映射表一致性校验。")]
    else:
        ok, lines = _check_device_against_mapping(eng, text, device_code, feature_code, device_name,
                                                  mac_group, pasted_override=pasted_override)
    for ln in lines:
        print(ln)
    if not ok:
        if args.assume_yes:
            print("已指定 --yes：设备信息与映射表不一致，仍继续。")
        elif not sys.stdin.isatty():
            print("✗ 设备信息与映射表不一致，且当前为非交互输入（无法读确认），已退出。"
                  "如确需继续请加 --yes。", file=sys.stderr)
            return 2
        else:
            try:
                ans = input("设备信息与映射表不一致，是否仍继续生成？[yes/y 继续，其它退出]：").strip().lower()
            except EOFError:
                ans = ""
            if ans not in ("yes", "y"):
                print("已取消：设备信息与映射表不一致，未生成。", file=sys.stderr)
                return 2

    try:
        sj = eng.parse_excel(args.file) if excel_input else eng.parse_pasted_text(text)
    except Exception as e:
        print(_err(f"解析输入失败：{e}"), file=sys.stderr)
        return 1

    # 解析内容概览（行数 / 提示语数 / FTP zip 数），透明可核对。
    parsed_rows = len(sj.get("rows", [])) if isinstance(sj, dict) else 0
    print(_c("—— 解析结果概览 ——", "bold"))
    print(f"  数据行数：{parsed_rows}")

    # 字段特征检查：值放错列（列错位）、格式不对、同一升级包 MD5 不一致、规则重复……生成 Excel 之前就拦住。
    from . import field_rules as FR
    issues = FR.check_rows(FR.from_source(sj.get("rows", [])), device_code or "", feature_code or "")
    fails = [t for s_, t in issues if s_ == "fail"]
    warns = [t for s_, t in issues if s_ == "warn"]
    print(_c("—— 字段特征检查 ——", "bold"))
    for t in warns[:20]:
        print(_warn(t))
    if fails:
        print(_err(f"字段特征检查未通过（{len(fails)} 处）：解析出的数据有字段放错列或格式不对，未生成 Excel、"
                   "没有动 CMS / S3："), file=sys.stderr)
        for t in fails[:30]:
            print(_c("  - " + t, "red"), file=sys.stderr)
        if len(fails) > 30:
            print(_c(f"  …… 还有 {len(fails) - 30} 处", "red"), file=sys.stderr)
        return 1
    print(_ok(f"字段特征检查通过（{parsed_rows} 行" + (f"，{len(warns)} 处 ⚠ 提醒" if warns else "") + "）。"))

    # 输出目录优先级：--out-dir > 输入 .txt 所在目录（--file 时）> OTA 工作目录 OTA_WORKSPACE。
    if args.out_dir:
        out_dir = args.out_dir
    elif args.file:
        # 默认把生成的 Excel 放到与输入 .txt 相同的目录，方便与源文件一起归档。
        out_dir = os.path.dirname(os.path.abspath(args.file))
    else:
        from .otagen import config as otagen_config
        out_dir = otagen_config.WORKSPACE

    # (2) 步骤1 生成（内容指纹去重：同一 source 内容已在该目录生成过则复用，不重复生成）。
    print(_c("—— 步骤1：生成 Excel ——", "bold"))
    content_hash = _file_fingerprint(args.file) if excel_input else _source_fingerprint(text)
    reused = None if args.force else _find_existing_excel(out_dir, content_hash)
    if reused:
        output_path = reused
        final_device_code = device_code or ""
        print(_ok(f"已存在相同内容的 Excel，跳过生成并复用：{output_path}"))
        print(_c("  （如需强制重新生成，请加 --force）", "dim"))
    else:
        try:
            res = eng.generate_excel(
                sj,
                device_code=device_code,
                device_name=device_name,
                feature_code=feature_code,
                mac_group=mac_group,
                output_dir=out_dir,
            )
        except ValueError as e:
            print(_err("步骤1 生成失败："), file=sys.stderr)
            print(str(e), file=sys.stderr)
            return 1

        output_path = os.path.abspath(res["output_path"])
        final_device_code = device_code or ""
        _record_excel_fingerprint(out_dir, content_hash, output_path)
        print(_ok(f"已生成：{output_path}"))
        print(f"  行数：{res.get('row_count')}  提示语数：{res.get('message_count')}")

    # (2b) 发起人 / Originator：--originator 优先；否则在终端询问（回车 = 沿用这份 Excel 之前记的，或留空）。
    originator = _ask_originator(args, output_path)

    # (3) 步骤2 严格校验。
    print(_c("—— 步骤2：严格校验 ——", "bold"))
    problems = _validate_excel(output_path)
    if problems:
        print(_err("步骤2 严格校验未通过："), file=sys.stderr)
        for prob in problems:
            print(_c("  - " + prob, "red"), file=sys.stderr)
        return 1
    print(_ok("严格校验通过（sheet / 18 列表头 / 每行地址与 OTA 类型 / 设备字段 均合格）。"))
    # (3b) 场测 MAC 组：确保 MAC（后 8 位）在 CMS『测试<设备类型代码>』组里。
    #      失败 → 醒目提醒并停止：不上传，也不记为『最近 prep』（publish / online / dass 不会默认用这份 Excel）。
    mac_out: dict = {}
    rc = _step_mac_group(args, sj, final_device_code, mac_group,
                         raw_text=_excel_all_text(args.file) if excel_input else text, out=mac_out)
    if rc:
        return _stop(output_path, rc)

    # 收尾：最终检查（MAC 组 / 提示语 / S3 / 行数）。通过才把这份 Excel 记为『最近 prep』。
    def finish(s3_check: bool, local_md5_ok: set) -> int:
        rc2 = _final_checks(args, output_path, final_device_code, mac_out.get("macs", []), mac_out.get("res"),
                            parsed_rows, s3_check=s3_check, local_md5_ok=local_md5_ok)
        if rc2:
            return _stop(output_path, rc2)
        from . import last_prep
        last_prep.save(output_path)   # 后续 publish / online / dass 不给文件时默认接着用它
        if originator:
            last_prep.save_meta(output_path, originator=originator)
            print(f"  发起人 / Originator：{originator}")
        return 0

    # (4) 步骤3 上传（或 --no-upload 计划）。
    urls = eng.extract_ftp_zip_urls(sj)
    bucket, prefix = transfer.build_s3_destination(final_device_code)

    if args.no_upload:
        print(_c("—— 步骤3：上传计划（--no-upload，不触碰 FTP/S3）——", "bold"))
        print(output_path)
        print(_warn(f"[--no-upload] 仅预览，未上传。S3 计划：s3://{bucket}/{prefix}"))
        print(f"[--no-upload] 待上传 FTP zip（{len(urls)} 个）：")
        for u in urls:
            print(f"  - {u}")
        if not urls:
            other = _non_ftp_source_urls(eng, sj)
            if other:
                print(_warn(f"源数据的 zip 地址不是 ftp://（共 {len(other)} 个，如 {other[0]}），"
                            "当前不支持下载，真实上传时会无文件可传。"))
        return finish(s3_check=False, local_md5_ok=set())

    if not urls:
        other = _non_ftp_source_urls(eng, sj)
        if other:
            print(_err(f"无可上传的 FTP zip：源数据的 zip 地址不是 ftp://（共 {len(other)} 个，如 {other[0]}），"
                       "当前只支持从 FTP 下载。"), file=sys.stderr)
        else:
            print(_err("无可上传的 FTP zip（源数据未含 ftp:// 地址）"), file=sys.stderr)
        return _stop(output_path, 1)

    # (4b) S3 预检：对每个 zip，若 S3 已存在同名文件且【大小一致】(与生成 Excel 里的「文件大小」比)，
    #      则跳过该文件的下载+上传（MD5 作辅助：能比就比，比上更放心）。--force 时不预检、全量重传。
    print(_c("—— 步骤3：S3 预检（大小+MD5 双重校验，一致则跳过）——", "bold"))
    size_md5_by_zip = _zip_size_md5_from_excel(output_path)
    skip_names = set()
    if not args.force:
        skip_names = _s3_precheck_skip(bucket, prefix, urls, size_md5_by_zip)
    todo_urls = [u for u in urls if _zip_name_from_url(u) not in skip_names]

    if not todo_urls:
        print(_ok(f"全部 {len(urls)} 个文件 S3 上已存在且大小一致，跳过下载和上传（complete）。"))
        print(output_path)
        print(f"S3 前缀：s3://{bucket}/{prefix}")
        return finish(s3_check=True, local_md5_ok=set())

    print(_c(f"—— 步骤3：FTP 下载 + 上传 S3（{len(todo_urls)}/{len(urls)} 个需处理）——", "bold"))

    # 先只下载，逐个比对 MD5 / 大小（与 Excel 一致才上传，坏包不进 S3），再上传。
    work_dir = tempfile.mkdtemp(prefix="otagen_")
    temp_dir = os.path.join(work_dir, "temp")
    progress = _make_progress_printer()
    result = transfer.download_and_upload(
        todo_urls, final_device_code, work_dir=work_dir,
        auto_upload=False, cleanup_after_upload=False, log=print, progress=progress,
    )
    print(f"已下载：{len(result.downloaded)}  失败：{len(result.failed)}")
    if result.failed:
        print(_err(f"下载失败项：{result.failed}"), file=sys.stderr)
        transfer.cleanup_zip_dir(temp_dir, log=print)
        return _stop(output_path, 1)

    print(_c("—— 步骤3：下载后 MD5 / 大小核对（与 Excel 一致才上传）——", "bold"))
    local_md5_ok, md5_bad = _verify_local_md5(result.downloaded, size_md5_by_zip)
    if md5_bad:
        print(_err("下载的升级包与 Excel 的 MD5 / 大小不一致，未上传："), file=sys.stderr)
        for b in md5_bad:
            print(_c("  - " + b, "red"), file=sys.stderr)
        transfer.cleanup_zip_dir(temp_dir, log=print)
        return _stop(output_path, 1)

    up = transfer.upload_dir_to_s3(temp_dir, final_device_code, log=print, progress=progress)
    print(output_path)
    print(f"S3 前缀：{up.s3_prefix}")
    if skip_names:
        print(_ok(f"S3 预检跳过（已存在且大小一致）：{len(skip_names)} 个"))
    print(f"已上传：{len(up.uploaded)}  失败：{len(up.failed)}")
    if up.failed:
        print(_err(f"上传失败项：{up.failed}"), file=sys.stderr)
        return _stop(output_path, 1)
    transfer.cleanup_zip_dir(temp_dir, log=print)

    # (5) 步骤3b 上传后校验：用 boto3 head_object 回查每个已上传文件确实在 S3 上。
    print(_c("—— 步骤3b：上传后 S3 回查 ——", "bold"))
    verify_bad = _verify_s3_uploads(bucket, prefix, up.uploaded)
    if verify_bad:
        print(_err("步骤3b 上传后 S3 校验未通过（以下文件上传后在 S3 未找到）："), file=sys.stderr)
        for b in verify_bad:
            print(_c("  - " + b, "red"), file=sys.stderr)
        return _stop(output_path, 1)
    return finish(s3_check=True, local_md5_ok=local_md5_ok)


_PREP_INDEX_FILE = ".prep_index.json"


def _source_fingerprint(text: str) -> str:
    """source 文本内容的 SHA-256 指纹（内容一变就变），用于判定是否已生成过。"""
    return hashlib.sha256(str(text).encode("utf-8")).hexdigest()


def _file_fingerprint(path: str) -> str:
    """文件字节内容的 SHA-256 指纹（Excel 输入用，不能按 utf-8 文本读）。"""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


_EXCEL_EXTS = (".xlsx", ".xlsm")


def _is_excel_input(path) -> bool:
    return bool(path) and str(path).lower().endswith(_EXCEL_EXTS)


def _norm_header(s) -> str:
    """表头归一：去掉不换行空格(\\xa0)/全角空格/所有空白并转小写，便于容错匹配。"""
    return re.sub(r"[\s\u00a0\u3000]+", "", str(s or "")).lower()


# 源 Excel（OTA 申请模板）里设备码 / feature code 的表头候选（均为 _norm_header 后的形式）。
_DEVICE_HEADERS = ("设备类型代码", "设备类型代码（方案号）", "设备类型代码(方案号)", "方案号", "devicecode")
_FEATURE_HEADERS = ("featurecode", "特征码", "feature")


def _read_excel_device_fields(path: str) -> dict:
    """从源 Excel 的列里读设备类型代码 / feature code（引擎对 Excel 不会抽取这两项）。

    返回 {"device_codes": [去重保序], "feature_codes": [去重保序]}。读取失败返回空列表。
    优先『升级规则』sheet，没有则取第一个 sheet；表头在第 1 行。
    """
    out = {"device_codes": [], "feature_codes": []}
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb["升级规则"] if "升级规则" in wb.sheetnames else wb.worksheets[0]
        rows = ws.iter_rows(values_only=True)
        header = next(rows, None) or ()
        norm = [_norm_header(h) for h in header]

        def col(cands):
            for i, h in enumerate(norm):
                if h in cands:
                    return i
            return None

        di, fi = col(_DEVICE_HEADERS), col(_FEATURE_HEADERS)
        for r in rows:
            for idx, key in ((di, "device_codes"), (fi, "feature_codes")):
                if idx is None or idx >= len(r) or r[idx] is None:
                    continue
                v = str(r[idx]).strip()
                if v and v not in out[key]:
                    out[key].append(v)
        wb.close()
    except Exception as e:
        print(_warn(f"读取 Excel 设备列失败：{e}"), file=sys.stderr)
    return out


def _non_ftp_source_urls(eng, sj) -> list:
    """源数据里存在、但不是 ftp:// 的 zip 地址（如 http://...），用于提示用户为何没有可下载的 FTP zip。"""
    seen = []
    for row in sj.get("rows", []) if isinstance(sj, dict) else []:
        try:
            u = str(eng.app.find_zip_source_url_in_row(row) or "").strip()
        except Exception:
            u = ""
        if u and not u.lower().startswith("ftp://") and u not in seen:
            seen.append(u)
    return seen


def _find_existing_excel(out_dir: str, content_hash: str) -> str | None:
    """在 out_dir 的指纹索引里查该内容是否已生成过 Excel，且该文件仍存在。

    返回已存在 xlsx 的绝对路径；没有记录、或记录的文件已被删则返回 None。
    """
    idx_path = os.path.join(out_dir, _PREP_INDEX_FILE)
    try:
        with open(idx_path, "r", encoding="utf-8") as f:
            idx = json.load(f)
    except Exception:
        return None
    rel = idx.get(content_hash)
    if not rel:
        return None
    path = rel if os.path.isabs(rel) else os.path.join(out_dir, rel)
    return os.path.abspath(path) if os.path.isfile(path) else None


def _record_excel_fingerprint(out_dir: str, content_hash: str, xlsx_path: str) -> None:
    """把 内容hash -> 生成的 xlsx 文件名 记进 out_dir 的指纹索引（存文件名，便于目录迁移）。"""
    idx_path = os.path.join(out_dir, _PREP_INDEX_FILE)
    try:
        with open(idx_path, "r", encoding="utf-8") as f:
            idx = json.load(f)
        if not isinstance(idx, dict):
            idx = {}
    except Exception:
        idx = {}
    idx[content_hash] = os.path.basename(xlsx_path)
    try:
        with open(idx_path, "w", encoding="utf-8") as f:
            json.dump(idx, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(_warn(f"写入指纹索引失败（不影响本次生成）：{e}"), file=sys.stderr)


def _zip_name_from_url(url: str) -> str:
    """从 ftp://.../xxx.zip 取文件名 xxx.zip（用于与 S3 key / Excel 行对应）。"""
    from urllib.parse import urlparse, unquote
    return os.path.basename(unquote(urlparse(str(url)).path or "")).strip()


def _zip_size_md5_from_excel(xlsx_path: str) -> dict:
    """从生成的 Excel 读 {zip文件名: {"size": 源文件大小(int或None), "md5": 源MD5小写或None}}。

    用『升级文件地址』列尾部的 zip 文件名作键，取同行『文件大小』『升级文件的MD5值』。
    这些值来自源数据（Excel），不现算、不下载——用于 S3 预检的大小/MD5 比对。
    """
    out = {}
    try:
        wb = load_workbook(xlsx_path, data_only=True)
        ws = wb["升级规则"] if "升级规则" in wb.sheetnames else wb.worksheets[0]
        hdr = [str(ws.cell(1, c).value or "").strip() for c in range(1, ws.max_column + 1)]
        iu = hdr.index("升级文件地址") + 1 if "升级文件地址" in hdr else None
        isz = hdr.index("文件大小") + 1 if "文件大小" in hdr else None
        im = hdr.index("升级文件的MD5值") + 1 if "升级文件的MD5值" in hdr else None
        if not iu:
            return out
        for r in range(2, ws.max_row + 1):
            url = str(ws.cell(r, iu).value or "").strip()
            if not url:
                continue
            name = os.path.basename(url)
            sz_raw = str(ws.cell(r, isz).value or "").strip() if isz else ""
            sz = int(re.sub(r"[^\d]", "", sz_raw)) if re.sub(r"[^\d]", "", sz_raw) else None
            md5 = (str(ws.cell(r, im).value or "").strip().lower() or None) if im else None
            if name:
                out[name] = {"size": sz, "md5": md5}
    except Exception as e:
        print(_warn(f"读取 Excel 大小/MD5 失败（将不跳过、照常上传）：{e}"), file=sys.stderr)
    return out


def _s3_precheck_skip(bucket: str, prefix: str, urls: list, size_md5_by_zip: dict) -> set:
    """对每个 zip 查 S3：已存在且【大小一致】则加入跳过集合（MD5 能比则作辅助确认）。

    返回可跳过（下载+上传都免）的 zip 文件名集合。查询异常/无法比对则不跳过（保守重传）。
    """
    import boto3
    from botocore.exceptions import ClientError

    s3 = boto3.client("s3")
    skip = set()
    for u in urls:
        name = _zip_name_from_url(u)
        if not name:
            continue
        key = f"{prefix}{name}"
        expect = size_md5_by_zip.get(name, {})
        exp_size = expect.get("size")
        exp_md5 = expect.get("md5")
        try:
            head = s3.head_object(Bucket=bucket, Key=key)
        except ClientError:
            print(f"  S3 无同名对象，需上传：{name}")
            continue
        except Exception as e:
            print(_warn(f"  S3 预检查询异常，保守重传：{name}（{e}）"))
            continue
        s3_size = head.get("ContentLength")
        if exp_size is None:
            print(_warn(f"  S3 有同名对象但 Excel 无『文件大小』可比，保守重传：{name}"))
            continue
        if s3_size != exp_size:
            print(_warn(f"  S3 大小不一致（S3={s3_size} vs 源={exp_size}），需重传：{name}"))
            continue
        # 大小一致 -> 跳过。MD5 作辅助：若对象 metadata 存了 md5 且与源不符则不跳过（更稳）。
        meta_md5 = str(head.get("Metadata", {}).get("md5", "")).strip().lower()
        if exp_md5 and meta_md5 and meta_md5 != exp_md5:
            print(_warn(f"  S3 大小一致但 MD5 metadata 不符（S3={meta_md5} vs 源={exp_md5}），需重传：{name}"))
            continue
        extra = ""
        if exp_md5 and meta_md5 == exp_md5:
            extra = "，MD5 亦一致"
        print("  " + _ok(f"已存在且大小一致{extra}，跳过：{name}"))
        skip.add(name)
    return skip


def _make_progress_printer():
    """返回一个 progress(pct, msg) 回调，用于显示 FTP 下载 / S3 上传进度百分比。

    节流：仅当百分比较上次增加 >=2%、或消息文案变化、或到达 100% 时才输出，避免分块回调刷屏。
    TTY 下用回车 `\\r` 原地刷新同一行（末尾达 100% 换行）；非 TTY（管道/重定向）下按节流打印整行，
    不写入 ANSI/回车控制符，保证被捕获输出干净。
    """
    state = {"last_pct": -100.0, "last_msg": ""}
    is_tty = sys.stdout.isatty()

    def progress(pct: float, msg: str) -> None:
        pct = max(0.0, min(100.0, float(pct)))
        changed_msg = (msg != state["last_msg"])
        if pct - state["last_pct"] < 2.0 and not changed_msg and pct < 100.0:
            return
        state["last_pct"] = pct
        state["last_msg"] = msg
        line = f"  [{pct:5.1f}%] {msg}"
        if is_tty:
            end = "\n" if pct >= 100.0 else ""
            print("\r" + line.ljust(70), end=end, flush=True)
        else:
            print(line, flush=True)

    return progress


def _verify_s3_uploads(bucket: str, prefix: str, filenames: list[str]) -> list[str]:
    """用 boto3 head_object 确认每个已上传文件真的在 S3 上（存在即成功），并打印 log。

    返回"校验失败（未找到/查询异常）"的 s3:// 地址列表；空列表表示全部确认存在。
    与上传同一套 boto3 凭据/region（~/.aws, us-west）。
    """
    if not filenames:
        return []
    import boto3

    s3 = boto3.client("s3")
    print(f"head_object 回查 {len(filenames)} 个文件 @ s3://{bucket}/{prefix}")
    bad: list[str] = []
    for name in filenames:
        key = f"{prefix}{name}"
        uri = f"s3://{bucket}/{key}"
        try:
            head = s3.head_object(Bucket=bucket, Key=key)
            size = head.get("ContentLength", "?")
            etag = str(head.get("ETag", "")).strip('"')
            print("  " + _ok(f"已确认存在：{uri}  (size={size}, etag={etag})"))
        except Exception as e:
            print("  " + _err(f"未找到/校验失败：{uri}  ({e})"), file=sys.stderr)
            bad.append(uri)
    if not bad:
        print(_ok(f"上传后校验通过：{len(filenames)} 个文件均已确认在 S3。"))
    return bad


def _stop(output_path: str, rc: int) -> int:
    """prep 没完成就停：若『最近 prep』记的正是这份 Excel，一并清掉（publish / online / dass 不会默认用它）。"""
    from . import last_prep
    if last_prep.forget(output_path):
        print(_c("  已清除『最近 prep』记录（它指向的正是这份 Excel）。", "red"), file=sys.stderr)
    return rc


def _verify_local_md5(paths: list[str], size_md5_by_zip: dict) -> tuple[set, list[str]]:
    """下载到本地的 zip 逐个算 MD5、量大小，与 Excel 比。返回 (一致的文件名集合, 问题列表)。"""
    from .prep_final import file_md5
    good, bad = set(), []
    for p in paths:
        name = os.path.basename(p)
        exp = size_md5_by_zip.get(name)
        if not exp:
            bad.append(f"{name}：Excel 里找不到这个文件对应的行")
            continue
        size = os.path.getsize(p)
        if exp.get("size") is not None and size != exp["size"]:
            bad.append(f"{name}：大小不一致 下载={size} / Excel={exp['size']}")
            continue
        md5 = file_md5(p)
        if not exp.get("md5"):
            bad.append(f"{name}：Excel 里没有 MD5，无法核对")
        elif md5 != exp["md5"].lower():
            bad.append(f"{name}：MD5 不一致 下载={md5} / Excel={exp['md5'].lower()}")
        else:
            print("  " + _ok(f"{name}：大小 {size}、MD5 {md5} 与 Excel 一致"))
            good.add(name)
    return good, bad


def _final_mac_item(F, args, group: str, macs: list[str], res: dict | None, col_bad: list[str]) -> dict:
    """最终检查第 1 项：由步骤2b 的 ensure_macs 结果 + Excel『MAC组』列问题得出。"""
    name = "场测 MAC 组"
    if args.no_mac:
        return F.item(name, "warn", "--no-mac：未检查 CMS MAC 组", col_bad)
    if not macs or not res:
        return F.item(name, "fail", "没有场测 MAC 或没有 MAC 组检查结果", col_bad)
    st = res.get("status")
    missing = res.get("missing", []) if st not in ("ok", "added", "created") else []
    lines = [f"{m}  " + ("✗ 不在组内" if m in missing else "✓ 在组内") for m in macs] + col_bad
    how = {"ok": "原本就在组内", "added": "本次已加入并回读确认", "created": "本次新建组并回读确认"}
    if col_bad:
        return F.item(name, "fail", "Excel『MAC组』列不对", lines)
    if st in how:
        return F.item(name, "pass", f"{len(macs)} 个 MAC 都在『{group}』（{how[st]}），Excel 每行 MAC组 正确", lines)
    if st in ("would_add", "would_create"):      # --no-upload 只查不改
        return F.item(name, "warn", res.get("detail", ""), lines)
    if st == "tested":                           # --mac-dry-run：填好后取消，没保存
        return F.item(name, "warn", f"--mac-dry-run 测试模式未保存，组里仍缺 {', '.join(missing)}", lines)
    return F.item(name, "fail", res.get("detail", "") or f"MAC 组检查未完成（{st}）", lines)


def _final_checks(args, output_path: str, device_code: str, macs: list[str], mac_res: dict | None,
                  parsed_rows: int,
                  s3_check: bool, local_md5_ok: set) -> int:
    """prep 最终检查：MAC 组 / 提示语 / 升级包 S3 / 规则行数。有 ✗ 返回 1。"""
    from . import prep_final as F
    from .mac_group import group_name
    from .otagen import config as otagen_config
    print(_c("—— 步骤4：最终检查 ——", "bold"))
    items = []
    group = group_name(device_code)
    rh, rrows = F.read_sheet(output_path, "升级规则")
    rr = F.rule_rows(rh, rrows)

    # 1) 场测 MAC 组：用步骤2b 的 CMS 结果（2b 已查过；加入/新建后 2b 也已回读确认），不再重复查 CMS；
    #    另核对 Excel『MAC组』列。
    items.append(_final_mac_item(F, args, group, macs, mac_res, F.check_mac_column(rr, group)))

    # 2) 字段特征（版本 / 机型 / 扩展码 / MD5 / 大小 / 品牌 …，列有没有错位）。
    items.append(F.check_fields(rr))

    # 3) 升级提示语。
    mh, mrows = F.read_sheet(output_path, "升级提示语")
    from .otagen.engine import OTAEngine
    defaults = [m["description"] for m in OTAEngine().app.get_default_ota_messages()]
    items.append(F.check_messages(mh, mrows, defaults=defaults, allow_default=args.default_messages))

    # 4) 升级包 / S3。
    aws_prefix = otagen_config.AWS_PREFIX
    up = F.url_problems(rr, device_code, aws_prefix)
    if up:
        items.append(F.item("升级包 / S3", "fail", f"{len(up)} 处地址 / MD5 / 大小不对", up))
    elif not s3_check:
        items.append(F.item("升级包 / S3", "skip", "--no-upload：未检查 S3"))
    else:
        import boto3
        s3 = boto3.client("s3")
        items.append(F.check_s3(rr, otagen_config.S3_UPLOAD_BUCKET, aws_prefix,
                                lambda b, k: s3.head_object(Bucket=b, Key=k), local_md5_ok))

    # 5) 规则行数。
    if len(rr) == parsed_rows and rr:
        items.append(F.item("规则行数", "pass", f"Excel {len(rr)} 行 = 输入 {parsed_rows} 行"))
    else:
        items.append(F.item("规则行数", "fail", f"Excel {len(rr)} 行，输入解析出 {parsed_rows} 行"))

    lines, failed = F.render(items)
    for ln in lines:
        print(ln)
    if not failed:
        print(_ok(f"prep 全部完成：{output_path}"))
    return 1 if failed else 0


# —— 不触碰 FTP/S3/~/.aws 的纯离线自测（python -m automation.prep --selftest）——
def _selftest() -> int:
    from datetime import datetime

    ok = True
    temp_dirs: list[str] = []

    def check(name, cond):
        nonlocal ok
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        if not cond:
            ok = False

    try:
        fixture = (
            "机型\t扩展信息\t源版本\t目标版本\tOTA类型\t升级包地址\n"
            "MT9655AMAndrGTV\tABCD1234\t0100AP0920\t0100AP1030\t1\t"
            "ftp://ftpuser@host/a/USBOTA_x.zip\n"
            "语种\t描述\n"
            "简体中文\t测试升级\n"
        )
        eng = OTAEngine()
        sj = eng.parse_pasted_text(fixture)

        out_dir = tempfile.mkdtemp(prefix="otagen_selftest_")
        temp_dirs.append(out_dir)
        res = eng.generate_excel(
            sj,
            device_code="MT9655AMAndrGTV",
            device_name="测试机",
            feature_code="",
            mac_group="测试MT9655AMAndrGTV",
            output_dir=out_dir,
        )
        gen = res["output_path"]

        # 设备码以映射表写法为准：大小写不敏感匹配后规范化；精确匹配/无匹配则不改写。
        check("规范化 MT9655andrWH80F -> MT9655AndrWH80F",
              _canonical_device_code(eng, "MT9655andrWH80F") == ("MT9655AndrWH80F", True))
        check("已是规范写法不改写",
              _canonical_device_code(eng, "MT9655AndrWH80F") == ("MT9655AndrWH80F", False))
        check("映射表外的设备码原样返回",
              _canonical_device_code(eng, "NOT_A_DEVICE") == ("NOT_A_DEVICE", False))
        check("MT9655AndrWH80F 映射表 feature code 已更正为 ...745",
              eng.device_preset("MT9655AndrWH80F").get("feature_code") == "861003009000005000000745")

        wb = load_workbook(gen, read_only=True)
        try:
            check("表名==['升级规则','升级提示语']", wb.sheetnames == ["升级规则", "升级提示语"])
            ws = wb["升级规则"]
            rows = list(ws.iter_rows(values_only=True))
            header = [("" if c is None else str(c)) for c in rows[0]]
            check("表头==OUTPUT_HEADERS", header == list(OUTPUT_HEADERS))
            first = {h: ("" if v is None else str(v)) for h, v in zip(OUTPUT_HEADERS, rows[1])}
            check("OTA类型 ∈ {1,2,11}", str(first.get("OTA类型", "")).strip() in ("1", "2", "11"))
            addr = str(first.get("升级文件地址", "") or "")
            check("升级文件地址 https:// 前缀", addr.startswith("https://"))
            check("升级文件地址 .zip 结尾", addr.lower().endswith(".zip"))
            check("升级文件地址 无 3+ 连续斜杠", not re.search(r"/{3,}", addr))
        finally:
            wb.close()

        check("_validate_excel 生成件无问题", _validate_excel(gen) == [])

        # 坏行：_validate_output_row_like 必须抓出各类问题（用正则，不用字面 '///'）。
        check("坏行 OTA类型=9 被抓",
              any("OTA类型" in p for p in _validate_output_row_like(
                  {"升级文件地址": "https://a/x.zip", "OTA类型": "9"})))
        check("坏行 地址非 http 被抓",
              any("升级文件地址" in p for p in _validate_output_row_like(
                  {"升级文件地址": "ftp://a/x.zip", "OTA类型": "1"})))
        check("坏行 地址含 3+ 斜杠被抓",
              any("连续斜杠" in p for p in _validate_output_row_like(
                  {"升级文件地址": "https://a///x.zip", "OTA类型": "1"})))
        # 特征码非 24 位：由 _validate_excel 层校验，这里单测 fullmatch 语义。
        check("特征码非 24 位判定", re.fullmatch(r"\d{24}", "12345") is None)

        # extract_ftp_zip_urls 去重。
        urls = eng.extract_ftp_zip_urls(sj)
        check("extract_ftp_zip_urls 去重为单条",
              urls == ["ftp://ftpuser@host/a/USBOTA_x.zip"])

        # build_s3_destination 目标。
        today = datetime.now().strftime("%Y%m%d")
        dest = transfer.build_s3_destination("CODE")
        check("build_s3_destination==('fam-media-andr','ota/CODE/<today>/')",
              dest == ("fam-media-andr", f"ota/CODE/{today}/"))

        # _build_parser 解析。
        ns = _build_parser().parse_args(
            ["--no-upload", "--file", "x.txt", "--device-code", "C"])
        check("解析 --no-upload/--file/--device-code", ns.no_upload and ns.file == "x.txt" and ns.device_code == "C")
        ns2 = _build_parser().parse_args([])
        check("--no-upload 默认 False", ns2.no_upload is False)
        check("--no-mac / --mac-dry-run 默认 False", ns2.no_mac is False and ns2.mac_dry_run is False)
        ns3 = _build_parser().parse_args(["--no-mac", "--mac-dry-run"])
        check("解析 --no-mac / --mac-dry-run", ns3.no_mac and ns3.mac_dry_run)
        # MAC 组步骤的离线分支（不打开浏览器）。
        import contextlib
        import io
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            r1 = _step_mac_group(ns3, {"rows": [{"MAC": "cc:12:28:cd:db:d8"}]}, "X", "")
            r2 = _step_mac_group(ns2, {"rows": [{"MAC": "空着为默认不设置"}]}, "X", "")
            r3 = _step_mac_group(ns2, {"rows": [{"MAC": "cc:12:28:cd:db:d8"}]}, "", "")
            buf = io.StringIO()
            with contextlib.redirect_stderr(buf):
                r4 = _step_mac_group(ns2, {"rows": [{"MAC": "cc:12:28:cd:db:d8", "机型信息": "65"},
                                                    {"MAC": "28cddbd", "机型信息": "75"}]}, "X", "")
            r4_msg = buf.getvalue()
        check("MAC 组：--no-mac 跳过", r1 == 0)
        check("MAC 组：没有 MAC → 停止", r2 == 1)
        check("MAC 组：有一行写法看不懂 → 停止并点出那一行", r4 == 1 and "第2行 75" in r4_msg and "『28cddbd』" in r4_msg)
        check(f"MAC 组：有 MAC 但没设备代码 → 退出（得到 {r3}）", r3 == 1)
        # 最终检查第 1 项：复用步骤2b 结果（不再查 CMS）。
        from . import prep_final as _F
        _m = ["b7fc8ce6", "b7fc8ce7"]
        _st = lambda a, r, cb=(): _final_mac_item(_F, a, "测试X", _m, r, list(cb))["status"]
        check("最终检查 MAC：2b ok / added / created → pass",
              all(_st(ns2, {"status": s, "missing": []}) == "pass" for s in ("ok", "added", "created")))
        check("最终检查 MAC：--no-upload 只查不改 / --mac-dry-run → warn",
              _st(ns2, {"status": "would_add", "missing": _m[:1]}) == "warn"
              and _st(ns2, {"status": "tested", "missing": _m[:1]}) == "warn")
        check("最终检查 MAC：failed / cancelled / Excel MAC组列错 → fail",
              _st(ns2, {"status": "failed"}) == "fail" and _st(ns2, {"status": "cancelled"}) == "fail"
              and _st(ns2, {"status": "ok", "missing": []}, ["第1行"]) == "fail")
        check("最终检查 MAC：--no-mac → warn", _st(ns3, None) == "warn")
        # 横排提示语（每个语种一列）+『日文 / 英文（…）』写法
        _ce = OTAEngine()
        patch_engine(_ce)
        check("语种别名：日文 / 英文（言語設定＝日本語以外）/ 繁體 / 西文",
              [resolve_language(_ce, x) for x in ("日文", "英文（言語設定＝日本語以外）", "繁體", "西文", "简体中文")]
              == ["日语", "英语", "繁体中文", "西班牙语", "简体中文"]
              and resolve_language(_ce, "描述") == "" and resolve_language(_ce, "重要なソフトウェア") == "")
        _col = ("规则行\n\n升级文案如下：\n\n日文\t英文（言語設定＝日本語以外）\n重要な更新。\n・安定化。\t"
                "An important update.\n- Stability.\n")
        _t, _l, _p = columnar_messages_to_rows(_ce, _col)
        check("横排提示语 → 竖排", _l == ["日语", "英语"] and not _p
              and "语种\t描述\n日语\t重要な更新。\n・安定化。\n英语\tAn important update.\n- Stability." in _t)
        _t2, _l2, _p2 = columnar_messages_to_rows(_ce, "日文\t英文\n只有一列没有 Tab")
        check("横排但 Tab 个数对不上 → 原样返回并报问题", _t2 == "日文\t英文\n只有一列没有 Tab" and "Tab" in _p2)
        check("竖排『语种<Tab>描述』不当横排处理", columnar_messages_to_rows(_ce, "语种\t描述\n英语\tx")[1] == [])
        _m = _ce.parse_pasted_text(
            "机型\t扩展信息\t源版本\t目标版本\t版本变化\t升级包地址\t升级文件MD5值\t升级包大小\tMAC\t品牌\tOTA类型\t区域\n"
            "32P1S(9A)\t32P1S9A888\t0001CQ0903\t0001CQ1116\tQ0903_to_Q1116\tftp://u@h/a/X.zip\t"
            "d82026de64b20398467b8c1423b19aaa\t118219077\t12122323\tREGZA\t非强制升级\t全球\n\n" + _t)["messages"]
        check("横排转换后引擎解析出日语、英语（多行保留）",
              [m["language"] for m in _m] == ["日语", "英语"] and "・安定化。" in _m[0]["description"])
        # 3 个字母的版本号（MSD6886：0100AWP0224）：列不能错位。
        _pe = OTAEngine()
        patch_engine(_pe)
        _awp = ("机型\t扩展信息\t源版本 \t目标版本 \t版本变化\t升级包地址\t升级文件MD5值\t升级包大小(Byte)\tMAC\t品牌 \tOTA类型\t区域\n"
                "43A53FUA\t43A53FUA0000784\t0100AWP0224\t0100AWQ1112\tP0224_to_Q1112\t"
                "ftp://u@h/a/X_OTA_P0224_to_Q1112.zip\t514b4e36f4bd0df7eed752d8528fe8ba\t227484372\t"
                "b7fc8ce6、b7fc8ce7\tHisense\t强制升级\t不限制\n")
        _r = _pe.parse_pasted_text(_awp)["rows"][0]
        check("3 个字母版本号：机型 / 扩展码 / 源版本 / 目标版本 不错位",
              (_r["机型信息"], _r["机器扩展码"], _r["源版本"], _r["目标版本"])
              == ("43A53FUA", "43A53FUA0000784", "0100AWP0224", "0100AWQ1112"))
        check("版本号识别：0100UQ1104 / 0100AWP0224 / V1.2.3 是，43A53FUA 不是",
              all(_pe.app.looks_like_version(x) for x in ("0100UQ1104", "0100AWP0224", "V1.2.3"))
              and not _pe.app.looks_like_version("43A53FUA"))
        # 下载后 MD5 / 大小核对。
        _d = tempfile.mkdtemp()
        temp_dirs.append(_d)
        _zp = os.path.join(_d, "a.zip")
        with open(_zp, "wb") as _f:
            _f.write(b"hello")
        _md5 = hashlib.md5(b"hello").hexdigest()
        with contextlib.redirect_stdout(io.StringIO()):
            g1, b1 = _verify_local_md5([_zp], {"a.zip": {"size": 5, "md5": _md5.upper()}})
            g2, b2 = _verify_local_md5([_zp], {"a.zip": {"size": 5, "md5": "0" * 32}})
            g3, b3 = _verify_local_md5([_zp], {"a.zip": {"size": 6, "md5": _md5}})
            g4, b4 = _verify_local_md5([_zp], {})
        check("下载后核对：MD5 一致（大小写无关）", g1 == {"a.zip"} and not b1)
        check("下载后核对：MD5 / 大小不一致、Excel 无此行 → 不上传",
              b2 and "MD5 不一致" in b2[0] and b3 and "大小不一致" in b3[0] and b4 and not (g2 or g3 or g4))

        # 自带子包可导入。
        import importlib
        for mod in ("automation.otagen.engine", "automation.otagen.transfer", "automation.otagen.config"):
            try:
                importlib.import_module(mod)
                check(f"import {mod} 成功", True)
            except Exception as e:
                check(f"import {mod} 成功（{e}）", False)

    finally:
        import shutil
        for d in temp_dirs:
            shutil.rmtree(d, ignore_errors=True)

    # —— 竖排粘贴：设备类型代码 / featurecode 作为表格前两列，提示语在表格后、用空格分隔 ——
    try:
        _e = OTAEngine()
        _vert = "\n".join([
            "设备类型代码（方案号）", "featurecode", "机型", "扩展信息", "品牌 ", "OTA类型", "区域", "源版本 ", "目标版本 ",
            "版本变化", "升级包地址", "升级包大小", "MD5", "场测MAC", "最大检测数",
            "A311D2AndrWE6GE", "861003009000010000000101", "65WE6GE(01)", "65WE6GE01888", "Hisense", "非强制升级", "全部",
            "0001UQ0815", "0001MQ0818", "Q0815_to_Q0818",
            "ftp://ftpuser@10.18.203.205/upload/A311D2/x/OTA_INC/65WE6GE_OTA_Q0815_to_Q0818.zip",
            "3073914", "debfa7e56660272595ba7e257eb3ce44", "cc:12:28:cd:db:d8", "空着为默认不设置", "",
            "英语 A Test version for EN16031 certification ."])
        _t, _lift = lift_device_columns(_e, _vert)
        _t, _langs = normalize_message_lines(_t)
        _t, _moved = move_messages_before_table(_e, _t)
        _info = _e.resolve_device_info_from_text(_t)
        _sj = _e.parse_pasted_text(_t)
        check("竖排：设备类型代码 / featurecode 作为表格列能解析",
              _lift is not None and _info.get("device_code") == "A311D2AndrWE6GE"
              and _info.get("feature_code") == "861003009000010000000101")
        check("竖排：只有 1 行数据，机型正确", len(_sj["rows"]) == 1 and _sj["rows"][0].get("机型信息") == "65WE6GE(01)")
        check("竖排：表格后『英语 + 空格 + 内容』提示语保留",
              _langs == ["英语"] and _moved == 1
              and any(m["description"].startswith("A Test version") for m in _sj.get("messages", [])))
        _sup = "\n".join(["设备类型代码（方案号）", "A311D2AndrWE6GE", "featurecode", "861003009000010000000101", "机型"])
        check("已支持的『标签/值』写法不被改动", lift_device_columns(_e, _sup)[1] is None)
        check("Tab 粘贴不挪提示语", move_messages_before_table(_e, "机型\t源版本\n英语：x")[1] == 0)
    except Exception as _ex:
        check(f"竖排粘贴自测异常：{_ex}", False)
    print("\n自测结果：", "全部通过 ✅" if ok else "存在失败 ❌")
    return 0 if ok else 1


if __name__ == "__main__":
    if "--selftest" in sys.argv[1:]:
        sys.exit(_selftest())
    sys.exit(main())
