#!/usr/bin/env bash
set -e

TARGET_DIR="/opt/asu_waybills"
if [ ! -d "$TARGET_DIR" ]; then
  FOUND=$(find /opt /home -name "server.py" 2>/dev/null | grep -v "putevka_update" | head -n 1)
  if [ -n "$FOUND" ]; then
    TARGET_DIR=$(dirname "$FOUND")
  else
    TARGET_DIR="/opt/asu_waybills"
  fi
fi

echo "=========================================="
echo "  Установка ПУТЁВКА v6.3.0"
echo "  Целевая папка: $TARGET_DIR"
echo "=========================================="

mkdir -p "$TARGET_DIR/data" "$TARGET_DIR/static" "$TARGET_DIR/templates_xls"
CURRENT_SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Копируем все файлы сборки v6.3.0
cp -rf "$CURRENT_SCRIPT_DIR"/* "$TARGET_DIR/"

# Обновляем БД до структуры v6.3.0 (схема А4 на весь лист и первичная настройка админа)
cp -f "$TARGET_DIR/data/app_data.json" "$TARGET_DIR/data/putevka_db.json"
chmod -R 777 "$TARGET_DIR/data" 2>/dev/null || true

cd "$TARGET_DIR"

echo "Освобождение порта 8080 и перезапуск server.py..."
fuser -k 8080/tcp 2>/dev/null || true
pkill -9 -f "server.py" 2>/dev/null || true
sleep 1

if [ -f "./venv/bin/python" ]; then
  nohup ./venv/bin/python server.py > server.log 2>&1 &
elif command -v python3 >/dev/null 2>&1; then
  nohup python3 server.py > server.log 2>&1 &
else
  nohup python server.py > server.log 2>&1 &
fi

sleep 2

if pgrep -f "server.py" >/dev/null; then
  echo ">>> УСПЕШНО! Сервер ПУТЁВКА v6.3.0 запущен и слушает порт 8080."
  echo ">>> PID процесса: $(pgrep -f "server.py" | head -n 1)"
  echo ">>> Версия: $(cat "$TARGET_DIR/version.txt")"
  ls -lh "$TARGET_DIR/data/putevka_db.json"
else
  echo "Ошибка старта! Вывод лога $TARGET_DIR/server.log:"
  cat "$TARGET_DIR/server.log"
fi
