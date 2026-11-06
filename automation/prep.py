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
        prog="python -m automation.prep",
        description="OTA 模板 prep CLI：生成 Excel → 严格校验 → 上传 S3（三步一键）。",
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
    return p


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


def main() -> int:
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

    # (1b) 设备信息与映射表一致性校验（生成前）。
    #   从文本解析出的 设备类型代码/feature code 与 mapping table(seed) 比对：
    #   一致 -> 打印 4 项(代码/名称/feature/测试MAC组) 并继续；
    #   不一致 -> 打印差异，等待人工确认(yes/y 继续，其它退出)；--yes 自动继续；无法交互则退出。
    ok, lines = _check_device_against_mapping(eng, text, device_code, feature_code, device_name, mac_group,
                                              pasted_override=pasted_override)
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

    # (3) 步骤2 严格校验。
    print(_c("—— 步骤2：严格校验 ——", "bold"))
    problems = _validate_excel(output_path)
    if problems:
        print(_err("步骤2 严格校验未通过："), file=sys.stderr)
        for prob in problems:
            print(_c("  - " + prob, "red"), file=sys.stderr)
        return 1
    print(_ok("严格校验通过（sheet / 18 列表头 / 每行地址与 OTA 类型 / 设备字段 均合格）。"))
    # 记下这份 Excel：后续 upload / online / dass 不给文件时默认接着用它。
    from . import last_prep
    last_prep.save(output_path)

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
        return 0

    if not urls:
        other = _non_ftp_source_urls(eng, sj)
        if other:
            print(_err(f"无可上传的 FTP zip：源数据的 zip 地址不是 ftp://（共 {len(other)} 个，如 {other[0]}），"
                       "当前只支持从 FTP 下载。"), file=sys.stderr)
        else:
            print(_err("无可上传的 FTP zip（源数据未含 ftp:// 地址）"), file=sys.stderr)
        return 1

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
        print(_ok("prep 全部完成：生成 → 校验 → S3 已是最新（免下载免上传）。"))
        return 0

    print(_c(f"—— 步骤3：FTP 下载 + 上传 S3（{len(todo_urls)}/{len(urls)} 个需处理）——", "bold"))

    work_dir = tempfile.mkdtemp(prefix="otagen_")
    result = transfer.download_and_upload(
        todo_urls,
        final_device_code,
        work_dir=work_dir,
        auto_upload=True,
        cleanup_after_upload=True,
        log=print,
        progress=_make_progress_printer(),
    )
    print(output_path)
    print(f"S3 前缀：{result.s3_prefix}")
    if skip_names:
        print(_ok(f"S3 预检跳过（已存在且大小一致）：{len(skip_names)} 个"))
    print(f"已下载：{len(result.downloaded)}  已上传：{len(result.uploaded)}  失败：{len(result.failed)}")
    if result.failed:
        print(_err(f"上传失败项：{result.failed}"), file=sys.stderr)
        return 1

    # (5) 步骤3b 上传后校验：用 boto3 head_object 回查每个已上传文件确实在 S3 上。
    print(_c("—— 步骤3b：上传后 S3 回查 ——", "bold"))
    verify_bad = _verify_s3_uploads(bucket, prefix, result.uploaded)
    if verify_bad:
        print(_err("步骤3b 上传后 S3 校验未通过（以下文件上传后在 S3 未找到）："), file=sys.stderr)
        for b in verify_bad:
            print(_c("  - " + b, "red"), file=sys.stderr)
        return 1
    print(_ok("prep 全部完成：生成 → 校验 → 上传 → S3 回查，均 correct。"))
    return 0


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

    print("\n自测结果：", "全部通过 ✅" if ok else "存在失败 ❌")
    return 0 if ok else 1


if __name__ == "__main__":
    if "--selftest" in sys.argv[1:]:
        sys.exit(_selftest())
    sys.exit(main())
