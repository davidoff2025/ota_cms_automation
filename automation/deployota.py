"""deployota —— OTA 部署统一入口（薄分发）：把 `deployota <子命令> [参数...]` 原样转发给对应模块。

子命令 -> 模块：
  prep     -> automation.prep              生成 Excel → 校验 → FTP 下载 → 上传 S3
  publish  -> automation.active_ota publish  upload + online + dass init，4 个检查点
  online   -> automation.active_ota online   上线（Excel 或 Rule ID）
  dass     -> automation.dass_fill         DASS 填报（init）
  update   -> automation.dass_update       放量：成功率闸门 → CMS 灰度 → DASS 阶段
  offline  -> automation.active_offline    按 Rule ID 下线
  query    -> automation.rule_info         按 Rule ID / 月份查部署信息与灰度阶段
  history  -> automation.depdb             部署历史
菜单见 _usage()；各子命令的详细帮助：./deployota <子命令> --help。
"""
from __future__ import annotations

import re
import runpy
import sys

# 子命令 -> (目标模块, 预置前缀参数)。
# 预置前缀用于把一个模块的两个子命令（active_ota 的 collect/online）拆成两个 deployota 子命令。
_SUBCOMMANDS: dict[str, tuple[str, list[str]]] = {
    "prep":   ("automation.prep",       []),
    "online": ("automation.active_ota", ["online"]),
    "publish": ("automation.active_ota", ["publish"]),
    "offline": ("automation.active_offline", []),
    "dass":   ("automation.dass_fill",  []),
    "update": ("automation.dass_update", []),
    "history": ("automation.depdb",     []),
    "query":  ("automation.rule_info",  []),
}

_PROG = "deployota"

# 需要 Excel 的子命令：命令行没给 Excel 时，默认用 prep 最近一次生成的那份。
_EXCEL_SUBCOMMANDS = ("online", "publish", "dass")
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
    # online 直接给 Rule ID（如 50000-50005）时不需要 Excel，原样转发、不注入 prep 的 Excel。
    if any(re.match(r"^\d[\d,，\-~\s]*$", a.strip()) for a in rest):
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
        "deployota —— OTA 部署统一入口",
        "",
        "用法：./deployota <子命令> [参数...]      ./deployota <子命令> --help 看详细帮助",
        "",
        "一次完整部署：",
        "  ./deployota prep --file 申请模板.xlsx      ① 生成标准 Excel + 升级包上传 S3",
        "  ./deployota publish                       ② CMS 导入 → Rule ID → 上线 → DASS 场测（4 个检查点）",
        "  ./deployota update ota50 <Rule ID>        ③ 逐档放量：ota50 → ota200 → ota1k → ota5k → ota20k → ota-all",
        "",
        "准备",
        "  prep      生成标准 Excel → 严格校验 → FTP 下载 → 上传 S3",
        "部署",
        "  publish   CMS 导入建规则 → 回收 Rule ID → 上线 → DASS init，4 个检查点，任一失败醒目提醒",
        "  dass      DASS 填报：dass init 插入新行（场测）",
        "上线 / 下线（只动 CMS）",
        "  online  <Rule ID | Excel>   上线；不给参数 = 用 prep 最近生成的 Excel",
        "  offline <Rule ID>           下线：在线的直接下线，已下线或找不到的跳过（不改 DASS）",
        "放量",
        "  update  <阶段> <Rule ID> [选项]",
        "          阶段：ota50 | ota200 | ota1k | ota5k | ota20k | ota-all（只前进）",
        "          流程：成功率闸门（ota200 起）→ CMS 在线规则灰度放开 → DASS 更新阶段",
        "          选项：--dry-run  测试模式，只在 CMS 走一遍流程、什么都不改",
        "                --yes      跳过 CMS 执行前的确认",
        "                --no-cms   不动 CMS，只更新 DASS",
        "                --strict   闸门不达标直接中止，不询问人工批复",
        "查询",
        "  query   <Rule ID | 月份>    所属批次、版本、灰度阶段时间线（如 query 65370-65379、query 2026.04-2026.10）",
        "  history                     部署历史列表（--device / --owner / --limit / --export）",
        "",
        "Rule ID 写法：65370 / 65370-65379 / 65370,65375 / 混用；设备类型代码按 Rule ID 自动查找（--device 可选）。",
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
