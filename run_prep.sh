#!/usr/bin/env bash
# OTA 模板 prep：从 stdin/--file 读升级信息文本 → 生成 Excel → 严格校验 → 上传 S3。
#
# 用法：
#   cat 文本.txt | ./run_prep.sh --device-code MT9655AMAndrGTV            # 生成+校验+上传
#   ./run_prep.sh --file 文本.txt --device-code MT9655AMAndrGTV            # 从文件读
#   cat 文本.txt | ./run_prep.sh --no-upload --device-code MT9655AMAndrGTV # 只生成+校验（不碰 FTP/S3）
#   ./.venv/bin/python -m automation.prep --selftest                       # 纯离线自测
#
# 真实上传（不带 --no-upload）前置条件：
#   - FTP：config/.env 里的 OTA_FTP_PASS 非占位值（见 config/.env.example）。
#   - AWS：区域/密钥不在 .env 配置，走 boto3 默认链（~/.aws），region=us-west。
# 『真实 S3 上传 / FTP 下载』不可逆，演练请加 --no-upload。
set -euo pipefail
cd "$(dirname "$0")"

PY="./.venv/bin/python"

# 1. 确认虚拟环境存在。
if [[ ! -x "$PY" ]]; then
  echo "未找到虚拟环境，请先执行：" >&2
  echo "  python3 -m venv .venv" >&2
  echo "  ./.venv/bin/python -m pip install -r requirements.txt" >&2
  exit 1
fi

# 2. 确认 config/.env 存在（缺则从示例创建）。
if [[ ! -f config/.env ]]; then
  echo "未找到 config/.env，正在从示例创建..." >&2
  cp config/.env.example config/.env
fi

# 3. 判断是否为 --no-upload；真实上传才要求 OTA_FTP_PASS 已填。
no_upload=false
for arg in "$@"; do
  if [[ "$arg" == "--no-upload" ]]; then
    no_upload=true
    break
  fi
done

if [[ "$no_upload" != true ]]; then
  ftp_pass="$(grep -E '^OTA_FTP_PASS=' config/.env | head -n1 | cut -d= -f2- || true)"
  if [[ -z "$ftp_pass" || "$ftp_pass" == "__set_me__" ]]; then
    echo "真实上传需要先在 config/.env 中填写 OTA_FTP_PASS（当前为空或仍是占位符 __set_me__）。" >&2
    echo "仅想生成+校验而不碰 FTP/S3，请加 --no-upload。" >&2
    exit 1
  fi
  echo "提示：AWS 区域/密钥不从 config/.env 读取，走 boto3 默认链（~/.aws），region=us-west。" >&2
fi

# 4. 运行 prep（参数透传；不消费 stdin，文本可经管道传入）。
exec "$PY" -m automation.prep "$@"
