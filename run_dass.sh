#!/usr/bin/env bash
# 把更新后的 Excel 填报到 DASS（安卓升级状态填报）。
#
# 用法：
#   ./run_dass.sh                                # init 模式，取 ota_v4 输出目录里最新的 .xlsx
#   ./run_dass.sh "/路径/更新后的文件.xlsx"             # init 模式，指定 Excel
#   ./run_dass.sh --update ota50 "/路径/文件.xlsx"     # update 模式，匹配已存在行并更新字段
#
# --update <option> 取值：init / ota50 / ota200 / ota1k / ota5k / ota20k / ota-all
#   init（默认）：读 Excel -> 插入 N 行 -> 逐格填入。
#   otaXX / ota-all：数据行已存在，不插入行；按 方案号/设备扩展信息/机型/目标版本/ruleid
#     五字段精确匹配定位已存在行，更新『当前阶段』『当前阶段时间』（ota50 另写『是否展示』=1）。
#
# 机制：读 Excel -> 按模式填入/更新 -> 截图并暂停，
#       在终端输入 approve 才会执行『验证 -> 提交』，否则不提交。
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

# 2. 确认 config/.env 存在且 DASS 凭据已填写。
if [[ ! -f config/.env ]]; then
  echo "未找到 config/.env，正在从示例创建..." >&2
  cp config/.env.example config/.env
fi
if grep -qE "^DASS_(USERNAME|PASSWORD)=__set_me__" config/.env; then
  echo "请先在 config/.env 中填写 DASS_USERNAME / DASS_PASSWORD（当前仍是占位符 __set_me__）。" >&2
  exit 1
fi

# 3. 运行 DASS 填报（把所有参数透传给脚本，可选 Excel 路径）。
exec "$PY" -m automation.dass_fill "$@"
