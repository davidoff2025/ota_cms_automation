#!/usr/bin/env bash
# Run the M1 login automation (open CMS, log in, open the OTA page).
set -euo pipefail
cd "$(dirname "$0")"

PY="./.venv/bin/python"

# 1. Ensure the project venv exists.
if [[ ! -x "$PY" ]]; then
  echo "未找到虚拟环境，请先执行：" >&2
  echo "  python3 -m venv .venv" >&2
  echo "  ./.venv/bin/python -m pip install -r requirements.txt" >&2
  echo "  ./.venv/bin/python -m playwright install chromium" >&2
  exit 1
fi

# 2. Ensure config/.env exists and credentials are filled in.
if [[ ! -f config/.env ]]; then
  echo "未找到 config/.env，正在从示例创建..." >&2
  cp config/.env.example config/.env
fi
if grep -q "__set_me__" config/.env; then
  echo "请先在 config/.env 中填写 CMS_USERNAME / CMS_PASSWORD（当前仍是占位符 __set_me__）。" >&2
  exit 1
fi

# 3. Run the login flow.
exec "$PY" -m automation.login
