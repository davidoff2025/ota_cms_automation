"""deployota —— OTA 部署统一入口（薄分发）：把 `deployota <子命令> [参数...]` 原样转发给对应模块。

子命令 -> 模块：
  prep     -> automation.prep              生成 Excel → 校验 → MAC 组 → FTP 下载 → 上传 S3
  publish  -> automation.active_ota publish  [--file → 先 prep] + upload + online + dass init，检查点 ⓪~④
  online   -> automation.active_ota online   上线（Excel 或 Rule ID）
  dass     -> automation.dass_fill         DASS 填报（init）
  update   -> automation.dass_update       放量：成功率闸门 → CMS 灰度 → DASS 阶段
  offline  -> automation.active_offline    按 Rule ID 下线
  query    -> automation.rule_info         按 Rule ID / 月份查部署信息与灰度阶段
  history  -> automation.depdb             部署历史
  mail     -> automation.mail_summary      给已部署的批次补发 HTML 汇总邮件（需 approve）
  help     -> 本菜单；help <子命令> = 该子命令的完整帮助
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
    "mail":   ("automation.mail_summary", []),
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
    # publish --file 原始输入：先跑 prep 生成 Excel，不注入旧的；相对路径由 prep 按 OTA_CALLER_CWD 解析。
    if any(a.split("=", 1)[0] == "--file" for a in rest):
        return rest
    # online 直接给 Rule ID（如 50000-50005）时不需要 Excel，原样转发、不注入 prep 的 Excel。
    if any(re.match(r"^[\d,，\-~～〜－–—\s]*\d[\d,，\-~～〜－–—\s]*$", a.strip()) for a in rest):
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
    return """\
deployota —— 海信海外安卓 OTA 部署统一入口

用法：./deployota <子命令> [参数...]
      ./deployota help <子命令>        看某个子命令的完整帮助（同 ./deployota <子命令> --help）

════════════════════════════════════════════════════════════════════════════════
常用流程
════════════════════════════════════════════════════════════════════════════════
  一条命令完成部署（prep → 导入 → Rule ID → 上线 → DASS 场测 → 汇总邮件）：
    ./deployota publish --file 申请模板.xlsx --originator 康凯(Corki)
  分两步：
    ./deployota prep --file 申请模板.xlsx            ① 生成 Excel + MAC 组 + 升级包上 S3
    ./deployota publish                              ② 用 prep 刚生成的 Excel 导入、上线、DASS
  逐档放量（每档先过成功率闸门）：
    ./deployota update ota50  65370-65379
    ./deployota update ota200 65370-65379   →  ota1k  →  ota5k  →  ota20k  →  ota-all
  演练（什么都不改）：
    ./deployota prep --file x.txt --no-upload        只生成 + 校验 + MAC 组只读检查
    ./deployota update ota200 65370 --dry-run        只在 CMS 走一遍流程
    ./.venv/bin/python -m automation.sim_publish --file x.txt    模拟整个 publish，最后 approve 才发邮件

════════════════════════════════════════════════════════════════════════════════
准备
════════════════════════════════════════════════════════════════════════════════
  prep  —— 生成标准 Excel → 字段特征检查 → 严格校验 → 场测 MAC 组 → FTP 下载 / MD5 核对 → 上传 S3 → 最终检查
    --file FILE            输入：.txt / .md（升级信息文本）或 .xlsx / .xlsm（申请模板）；省略则读 stdin
    --originator 姓名       发起人（不给就问，可回车跳过）；显示在汇总和邮件里
    --device-code CODE     设备类型代码（覆盖解析值）      --device-name NAME   设备类型名称
    --feature-code CODE    特征码（覆盖解析值）            --mac-group NAME     MAC 组名
    --out-dir DIR          输出目录（默认 = 输入文件所在目录）
    --no-upload            只生成 + 校验，不碰 FTP / S3（MAC 组只查不改）
    --yes, -y              设备信息与映射表不一致时自动继续；MAC 组保存前不再问 yes
    --force                强制重新生成 Excel、全量重传（默认复用相同内容的 Excel、跳过 S3 已有的包）
    --no-mac               跳过场测 MAC 组检查（开头、结尾红字提醒）
    --mac-dry-run          MAC 组填好后点『取消』不保存（然后继续上传）
    --default-messages     源文件确实没有提示语、要用内置默认文案时加（否则最终检查判失败）
    最终检查：MAC 组 / 字段特征（列错位、格式、同名升级包、重复行）/ 提示语（简体中文第一、
              不是默认文案）/ S3 大小 + MD5 / 规则行数。任一 ✗ → 不上传、不记为『最近 prep』。

════════════════════════════════════════════════════════════════════════════════
部署
════════════════════════════════════════════════════════════════════════════════
  publish [EXCEL]  —— ① 导入模板 → ② 回收 Rule ID 回写 Excel → ③ 逐条上线 → ④ DASS init → 汇总邮件
    --file FILE            先跑 prep 作为检查点 ⓪（prep 的参数都可以一起写，--no-upload 除外）
    --no-dass              只做 ①~③
    --auto-approve         DASS 全部填写成功时自动提交，不等 approve
    --mail-to "…"          汇总邮件收件人（可直接贴 Outlook 收件人串）；不给就在最后弹出输入框
    冲突中止不算失败：③ 标 ⚠，④ 只填已上线的规则，邮件里标『冲突未上线』。
    退出码：0 完成；1 运行出错；2 导入前检查不通过；10 prep；11 导入；12 Rule ID；13 上线；14 DASS。

  dass init [EXCEL]  —— DASS『安卓升级状态填报』插入新行（场测）；publish 的第 ④ 步就是它
    --only 65394-65399     只填这些 Rule ID（其余不填，邮件里标『冲突未上线』）
    --auto-approve         全部填写成功时自动提交
    --mail-to "…"          汇总邮件收件人
    （放量请用 update；旧用法 dass <otaXX> <Excel> 不经过 CMS 灰度，不推荐）

════════════════════════════════════════════════════════════════════════════════
上线 / 下线（只动 CMS）
════════════════════════════════════════════════════════════════════════════════
  online [EXCEL | Rule ID...]  —— 逐条上线；已在线的跳过。不给参数 = prep 最近生成的 Excel
    冲突时：旧规则 mac 组非空才替换；为空或读不到 → 中止该条（记录冲突）
    --device CODE          设备类型代码查不到时手工指定（一般自动查：部署库 → prep Excel → CMS）
    --force-replace        强制替换（只用于按 Rule ID）：冲突时旧规则 mac 组为空也替换。
                           开始前列出 Rule ID、你输入 yes；遇到冲突当场点『确定』，替换后回读确认
      --dry-run            配合 --force-replace：只读，列出会被替换的旧规则和它们的 mac 组
      --yes                配合 --force-replace：跳过开始前的确认
    ⚠ 旧规则 mac 组为空 = 已对所有人放开；被替换下线后撤销不了。

  offline <Rule ID...>  —— 在线的直接下线（确认框点确定、回读确认）；已下线 / 找不到的跳过。不改 DASS
    --yes                  跳过开始前的确认       --device CODE   手工指定设备类型代码

════════════════════════════════════════════════════════════════════════════════
放量
════════════════════════════════════════════════════════════════════════════════
  update [阶段] <Rule ID...>  —— 成功率闸门 → CMS 在线规则灰度放开 → DASS 更新『当前阶段』 → 放量汇总邮件
    ./deployota update 65370-65379          不写阶段：按 CMS『阶段名称』自动进到下一阶段
                                            （一批里阶段不同就分组；先列出计划，确认一次后逐组执行）
    阶段（只前进）：ota50 → ota200 → ota1k → ota5k → ota20k → ota-all
    两道闸门（ota200 起；ota50 不检查）：成功率每个 > 50%；间隔距进入上一阶段 > 48 小时
      都通过 → 全自动执行 CMS 灰度 + DASS 更新并提交，不再问确认（加 --confirm 恢复逐步确认）
      任一不通过 → 停下列出原因，输入 override 放行整组（--strict 不询问、直接中止）
    ota20k 再看『源版本用户数』（整批，> 20,000 才算大）：
       当前 5K：有大的 → 放 20K；全部 ≤ 20,000 → 直接完全开放
       当前 20K：有大的 → 再放一轮 20K（CMS 累计 +20000，DASS 不改）；全部 ≤ 20,000 → 完全开放
    --dry-run              测试模式：只在 CMS 走一遍（配置窗口点取消），不做闸门、不改 DASS
    --confirm              每一步都要你确认（CMS 执行前 yes、20K 判断 yes、DASS approve）
    --yes                  配合 --confirm：跳过 CMS 执行前、20K 判断结果的 yes 确认
    --auto-approve         DASS 全部写入成功时自动提交（不绕过成功率闸门）
    --no-cms               不动 CMS，只更新 DASS
    --strict               闸门不达标直接中止，不询问 override
    --skip-missing         部署库解析不出的 Rule ID 跳过（默认整体中止）
    --mail-to "…"          放量汇总邮件额外收件人      --no-mail   不发放量汇总邮件
    结束打印『同意语句』，如：同意（OTA200，20261113，50032~50033）50032:83%，50033:70%
    放量汇总邮件（带灰度进度条：场测 → OTA50 → … → 完全开放）：To 数据统计报告、Bcc devinw；
      有规则放开就发（失败的也写进去），执行了但一条都没放开发失败通知；测试模式 / 取消 / 无变化不发。
      HTML 存 OTADATA/reports/，每次记录到 OTADATA/update_runs.jsonl。

════════════════════════════════════════════════════════════════════════════════
查询 / 邮件（只读本地部署库；mail 只发邮件）
════════════════════════════════════════════════════════════════════════════════
  query <Rule ID... | 月份>  —— 所属批次、版本、发起人、灰度进度（[■■■■■□□] 5K + 每档日期）、最近一次 update、CMS 状态
    ./deployota query 65370-65379,65385      ./deployota query 2026.04-2026.10
  history  —— 部署历史，最新的在最上面；每批带灰度进度 [■■■■■□□] 5K（阶段不同时 5K×4/完全开放×1）
    --device CODE  --originator 姓名  --rule 65370  --limit N（默认 20；0 = 全部）  --export（导出 CSV）
  mail <Rule ID>  —— 给已部署的批次补发 HTML 部署汇总邮件：粘贴收件人 → 看草稿 → 输入 approve 才发送
    --originator 姓名     --mail-to "…"（给了就不弹输入框）
    收件人 = 你粘贴的人；固定抄送 数据统计报告 <sys.datareports@hisense.com>；固定密送 devinw@hisense.com

════════════════════════════════════════════════════════════════════════════════
通用
════════════════════════════════════════════════════════════════════════════════
  Rule ID 写法：65370 / 65370-65379 / 65370~65379 / 65370,65375，可以混用、可以带空格
  Excel：不给时用 prep 最近生成（并通过最终检查）的那份；相对路径按当前目录
  CMS_HEADED=true ./deployota …    这一次显示浏览器（config/.env 里 CMS_HEADED=false = 默认后台运行）
  配置：config/.env（CMS / DASS / FTP 账号、OTA_SMTP_* 邮件服务器、OTA_MAIL_CC / OTA_MAIL_BCC）
  导入、上线、替换、下线、灰度、DASS 提交都会改线上数据，撤销不了；不确定时先用 --dry-run / --no-upload。
  手册：README.md"""


def main() -> int:
    argv = sys.argv[1:]
    if not argv or argv[0] in ("-h", "--help"):
        print(_usage())
        return 0 if argv else 2

    sub = argv[0]
    if sub == "help":   # ./deployota help [子命令]
        if len(argv) < 2:
            print(_usage())
            return 0
        argv = [argv[1], "--help"]
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
