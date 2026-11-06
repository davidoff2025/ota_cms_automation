"""deployota —— OTA 部署统一入口（薄分发壳）。

把分散的各功能模块归到一个命令下，便于记忆与将来编排成工作流。本模块【只做分发】：
把 `deployota <子命令> [参数...]` 的剩余参数【原样转发】给对应的现有模块执行，
不重写、不改动任何被转发模块的逻辑/参数/自测。旧入口（run_prep.sh / run_active_ota.sh /
run_dass.sh / run_login.sh）与本命令【并存】。

子命令映射：
  prep    -> automation.prep            生成 Excel → 严格校验 → 上传 S3（上传后 boto3 校验）
  upload  -> automation.active_ota collect   导入 OTA 模板建规则 + 收集 Rule ID 回写 Excel
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
    "dass":   ("automation.dass_fill",  []),
    "history": ("automation.depdb",     []),
    "backfill": ("automation.depdb_backfill", []),
}

_PROG = "deployota"


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
        "  dass     DASS 填报/更新（OPTION: init/ota200/ota1k/ota5k/ota20k/ota-all，含成功率闸门；完成后自动显示汇总并记录部署）",
        "  history  查看 OTA 部署历史（--device 过滤 / --limit N / --export 导出 CSV）",
        "  backfill CMS 部署历史回补（只读抓取 OTA升级模板 → 去重合并入库；默认 dry-run，--commit 才写）",
        "",
        "示例：",
        "  ./deployota prep --file x.txt --device-code MT9655AMAndrGTV",
        "  ./deployota upload /路径/更新后的.xlsx",
        "  ./deployota online /路径/更新后的.xlsx",
        "  ./deployota dass ota200 /路径/更新后的.xlsx",
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
    forwarded = prefix + argv[1:]

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
