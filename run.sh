#!/usr/bin/env bash
# 在项目根目录启动；使用 .env 中的 HOST / PORT
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT"
if [[ ! -d venv ]]; then
  echo "请先创建虚拟环境: python3 -m venv venv && source venv/bin/activate && pip install -r requirements.txt"
  exit 1
fi
# shellcheck source=/dev/null
source "$ROOT/venv/bin/activate"
export PYTHONPATH="$ROOT"
# 注意：.env 里含空格的值必须写成 KEY="value"（双引号），否则 source 会拆成多条命令
set -a
[[ -f "$ROOT/.env" ]] && . "$ROOT/.env"
set +a
HOST="${HOST:-0.0.0.0}"
PORT="${PORT:-8765}"
# 必须用 venv 里的 Python 执行，否则会用系统 uvicorn/全局环境，缺少 python-docx 等依赖
exec "$ROOT/venv/bin/python" -m uvicorn app.main:app --host "$HOST" --port "$PORT"
