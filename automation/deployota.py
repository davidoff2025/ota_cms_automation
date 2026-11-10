"""deployota —— OTA 部署统一入口（薄分发壳）。

把分散的各功能模块归到一个命令下，便于记忆与将来编排成工作流。本模块【只做分发】：
把 `deployota <子命令> [参数...]` 的剩余参数【原样转发】给对应的现有模块执行，
不重写、不改动任何被转发模块的逻辑/参数/自测。旧入口（run_prep.sh / run_active_ota.sh /
run_dass.sh / run_login.sh）与本命令【并存】。

子命令映射：
  prep    -> automation.prep            生成 Excel → 严格校验 → 上传 S3（上传后 boto3 校验）
  upload  -> automation.active_ota collect   导入 OTA 模板建规则 + 收集 Rule ID 回写 Excel
  publish -> automation.active_ota publish   upload + online + dass init 合并，4 个检查点
  online  -> automation.active_ota online    在「规则上下线编排」逐条上线 + 冲突处理
  dass    -> automation.dass_fill       DASS 填报/更新（OPTION: init/ota200/.../ota-all，含成功率闸门）

用法：
  ./deployota <子命令> [该子命令的参数...]
  ./deployota --help                 列出子命令
  ./deployota <子命令> --help         查看该子命令自己的帮助（转发给对应模块）

示例：
  ./deployota prep --file x.txt --device-code MT9655AMAndrGTV
  ./deployota upload /路径/更新后的.xlsx
  ./deployota online /路径/更新后的.xlsx
  ./deployota dass ota200 /路径/更新后的.xlsx
"""
from __future__ import annotations

import runpy
import sys

# 子命令 -> (目标模块, 预置前缀参数)。
# 预置前缀用于把一个模块的两个子命令（active_ota 的 collect/online）拆成两个 deployota 子命令。
_SUBCOMMANDS: dict[str, tuple[str, list[str]]] = {
    "prep":   ("automation.prep",       []),
    "upload": ("automation.active_ota", ["collect"]),
    "online": ("automation.active_ota", ["online"]),
    "publish": ("automation.active_ota", ["publish"]),
    "dass":   ("automation.dass_fill",  []),
    "history": ("automation.depdb",     []),
    "backfill": ("automation.depdb_backfill", []),
}

_PROG = "deployota"

# 需要 Excel 的子命令：命令行没给 Excel 时，默认用 prep 最近一次生成的那份。
_EXCEL_SUBCOMMANDS = ("upload", "online", "publish", "dass")
_EXCEL_EXTS = (".xlsx", ".xlsm")


def _resolve_excel_args(rest: list[str]) -> list[str] | None:
    """处理 upload/online/dass 的 Excel 参数；失败返回 None（已打印原因）。

    - 命令行里已有 .xlsx/.xlsm 参数：相对路径按调用者所在目录（OTA_CALLER_CWD）转成绝对路径。
    - 没有：取 prep 最近一次生成的 Excel 追加进去；找不到记录或文件已不在则报错，不猜。
    - 带 -h/--help/--selftest 时原样转发，不注入。
    """
    import os
    if any(a in ("-h", "--help", "--selftest") for a in rest):
        return rest
    base = os.environ.get("OTA_CALLER_CWD") or os.getcwd()
    out, has_excel = [], False
    for a in rest:
        if not a.startswith("-") and a.lower().endswith(_EXCEL_EXTS):
            has_excel = True
            a = os.path.join(base, os.path.expanduser(a))
        out.append(a)
    if has_excel:
        return out
    from . import last_prep
    rec = last_prep.load()
    if not rec:
        print("✗ 没有指定 Excel，也没有 prep 的生成记录。请先运行 ./deployota prep，或在命令里给出 Excel 路径。",
              file=sys.stderr)
        return None
    path = rec["excel"]
    if not os.path.isfile(path):
        print(f"✗ prep 最近生成的 Excel 已不存在：{path}\n  请重新运行 prep，或在命令里给出 Excel 路径。",
              file=sys.stderr)
        return None
    print(f"\033[32m✓ 未指定 Excel，使用 prep 最近生成的文件（{rec.get('time', '')}）：\033[0m\n  {path}"
          if sys.stdout.isatty() and not os.environ.get("NO_COLOR") else
          f"✓ 未指定 Excel，使用 prep 最近生成的文件（{rec.get('time', '')}）：\n  {path}")
    return out + [path]


def _usage() -> str:
    lines = [
        "deployota —— OTA 部署统一入口（薄分发）。",
        "",
        "用法：",
        f"  ./{_PROG} <子命令> [该子命令的参数...]",
        f"  ./{_PROG} --help               列出子命令",
        f"  ./{_PROG} <子命令> --help       查看该子命令的帮助",
        "",
        "子命令：",
        "  prep     生成 Excel → 严格校验 → 上传 S3（上传后 boto3 回查确认）",
        "  upload   导入 OTA 模板建规则 + 收集 Rule ID 回写 Excel（CMS）",
        "  online   在「规则上下线编排」逐条上线 + 冲突处理（CMS）",
        "  publish  upload + online + dass init 一步到位，4 个检查点（导入成功 / 得到 Rule ID / 上线成功 / DASS 填报完成），任一失败醒目提醒（--no-dass 则不含 DASS）",
        "  dass     DASS 填报/更新（OPTION: init/ota200/ota1k/ota5k/ota20k/ota-all，含成功率闸门；完成后自动显示汇总并记录部署）",
        "  history  查看 OTA 部署历史（--device 过滤 / --limit N / --export 导出 CSV）",
        "  backfill CMS 部署历史回补（只读抓取 OTA升级模板 → 去重合并入库；默认 dry-run，--commit 才写）",
        "",
        "upload / online / publish / dass 不给 Excel 时，默认使用 prep 最近一次生成的那份（记录在 .last_prep.json）。",
        "",
        "示例：",
        "  ./deployota prep --file x.txt --device-code MT9655AMAndrGTV",
        "  ./deployota upload                    # 用 prep 刚生成的 Excel",
        "  ./deployota online                    # 同上（upload 已把 Rule ID 回写到这份 Excel）",
        "  ./deployota publish                   # upload + online + dass init 一步到位，带检查点",
        "  ./deployota dass ota200               # 同上",
        "  ./deployota upload /路径/更新后的.xlsx   # 也可显式指定",
    ]
    return "\n".join(lines)


def main() -> int:
    argv = sys.argv[1:]
    if not argv or argv[0] in ("-h", "--help"):
        print(_usage())
        return 0 if argv else 2

    sub = argv[0]
    if sub not in _SUBCOMMANDS:
        print(f"✗ 未知子命令：{sub!r}", file=sys.stderr)
        print(f"可用子命令：{', '.join(_SUBCOMMANDS)}。详见 ./{_PROG} --help。", file=sys.stderr)
        return 2

    module, prefix = _SUBCOMMANDS[sub]
    rest = argv[1:]
    if sub in _EXCEL_SUBCOMMANDS:
        rest = _resolve_excel_args(rest)
        if rest is None:
            return 2
    forwarded = prefix + rest

    # 以“被转发模块作为 __main__”运行：等价于 python -m <module> <forwarded>，
    # 该模块的 if __name__ == "__main__" 守卫（含 --selftest/ main()）照常生效。
    sys.argv = [module.replace(".", "/") + ".py"] + forwarded
    try:
        runpy.run_module(module, run_name="__main__", alter_sys=True)
    except SystemExit as e:  # 各模块用 sys.exit(main()) 返回码，透传出去
        return int(e.code) if isinstance(e.code, int) else (0 if e.code is None else 1)
    return 0


if __name__ == "__main__":
    sys.exit(main())
