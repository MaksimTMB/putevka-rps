#!/usr/bin/env bash
set -e
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$PROJECT_ROOT"

lsof -ti:8080 | xargs kill -9 2>/dev/null || true
if [ ! -d "venv" ]; then
  python3 -m venv venv
fi
./venv/bin/pip install -r requirements.txt
cp -f data/app_data.json data/putevka_db.json
nohup ./venv/bin/python server.py > server.log 2>&1 &
sleep 1
echo ">>> Обновление до v6.5.0 завершено! Сервер слушает порт 8080."
