#!/usr/bin/env bash
# CMS active_ota：导入 OTA 模板建规则 → 逐条上线并处理冲突 → 回读状态汇总。
#
# 用法（两个独立子命令；一条龙后续再 consolidate）：
#   ./run_active_ota.sh collect ["/路径/模板.xlsx"]   # 导入模板建规则 + 收集本批 Rule ID（升序打印）
#   ./run_active_ota.sh online  ["/路径/模板.xlsx"]   # 从 Excel 的『Rule ID』列读 → 逐条上线 + 冲突 + 汇总
#   ./run_active_ota.sh --probe-online               # 只读 dump 编排视图选择器（不做任何写操作）
#   ./run_active_ota.sh --selftest                   # 纯逻辑自测（不连浏览器）
#   不给 Excel 则取 ota_v4 输出目录里最新的 .xlsx。
#
# collect：导入建规则（名称=设备类型代码_目标版本末5位；备注=今天YYYYMMDD_USA_去表头行数），
#          再去「OTA升级模板」按标题精确匹配本批行 →「查看规则」读『规则ID』列，升序打印。
# online ：读 Excel（sheet『升级规则』）里表头为『Rule ID』的列 → 在「升级事件→规则上下线」
#          编排弹窗逐条上线；遇『规则替换』冲突：查被冲突号的 mac组，非空才替换上线、为空则取消中止该条；
#          最后回读每个 Rule ID 的状态（online / 冲突中止 / 失败）并汇总。
#
# 安全：所有写操作（导入确定/上线确定/替换确定）均经二次确认 gate，条件不满足不点。
#       端到端写流程改线上生产数据，请在知情下运行本脚本。
set -euo pipefail
cd "$(dirname "$0")"

PY="./.venv/bin/python"

# 1. 确认虚拟环境存在。
if [[ ! -x "$PY" ]]; then
  echo "未找到虚拟环境，请先执行：" >&2
  echo "  python3 -m venv .venv" >&2
  echo "  ./.venv/bin/python -m pip install -r requirements.txt" >&2
  echo "  ./.venv/bin/python -m playwright install chromium" >&2
  exit 1
fi

# 2. 确认 config/.env 存在且 CMS 凭据已填写。
if [[ ! -f config/.env ]]; then
  echo "未找到 config/.env，正在从示例创建..." >&2
  cp config/.env.example config/.env
fi
if grep -qE "^CMS_(USERNAME|PASSWORD)=__set_me__" config/.env; then
  echo "请先在 config/.env 中填写 CMS_USERNAME / CMS_PASSWORD（当前仍是占位符 __set_me__）。" >&2
  exit 1
fi

# 3. 运行 active_ota（把所有参数透传给脚本）。
exec "$PY" -m automation.active_ota "$@"
