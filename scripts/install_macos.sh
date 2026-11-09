#!/usr/bin/env bash
set -e
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$PROJECT_ROOT"

echo "=================================================="
echo "  Установка АСУ ПУТЁВКА v6.5.0 (macOS)"
echo "=================================================="

python3 -m venv venv
./venv/bin/pip install --upgrade pip
./venv/bin/pip install -r requirements.txt

if [ ! -f "data/putevka_db.json" ]; then
  cp -f data/app_data.json data/putevka_db.json
fi

lsof -ti:8080 | xargs kill -9 2>/dev/null || true
nohup ./venv/bin/python server.py > server.log 2>&1 &
sleep 1
echo ">>> Сервер ПУТЁВКА v6.5.0 запущен на http://localhost:8080 (PID: $(lsof -ti:8080 | head -n 1))"
