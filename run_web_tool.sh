#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

if [ ! -d ".venv" ]; then
  python3 -m venv .venv
fi

source .venv/bin/activate
pip install -q -r requirements.txt

PORT="${PORT:-8502}"
python -m streamlit run app.py --server.address 0.0.0.0 --server.port "$PORT"
