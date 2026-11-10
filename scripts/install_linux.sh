#!/usr/bin/env bash
set -e

TARGET_DIR="/opt/asu_waybills"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

echo "=================================================="
echo "  Чистая установка АСУ ПУТЁВКА v6.5.0 (Linux)"
echo "  Директория установки: $TARGET_DIR"
echo "=================================================="

mkdir -p "$TARGET_DIR/data" "$TARGET_DIR/static" "$TARGET_DIR/templates_xls" "$TARGET_DIR/scripts"
if [ "$PROJECT_ROOT" != "$TARGET_DIR" ]; then
  cp -rf "$PROJECT_ROOT"/* "$TARGET_DIR/"
fi

cd "$TARGET_DIR"

if ! command -v python3 >/dev/null 2>&1; then
  echo "Ошибка: python3 не найден. Установите пакет python3 и python3-venv."
  exit 1
fi

if [ ! -d "$TARGET_DIR/venv" ]; then
  python3 -m venv "$TARGET_DIR/venv"
fi

"$TARGET_DIR/venv/bin/pip" install --upgrade pip
"$TARGET_DIR/venv/bin/pip" install -r "$TARGET_DIR/requirements.txt"

if [ ! -f "$TARGET_DIR/data/putevka_db.json" ]; then
  cp -f "$TARGET_DIR/data/app_data.json" "$TARGET_DIR/data/putevka_db.json"
fi
chmod -R 777 "$TARGET_DIR/data" 2>/dev/null || true

if command -v systemctl >/dev/null 2>&1 && [ -d "/etc/systemd/system" ]; then
  cp -f "$TARGET_DIR/scripts/putevka.service" /etc/systemd/system/putevka.service
  systemctl daemon-reload
  systemctl enable putevka
  systemctl restart putevka
  echo ">>> Служба systemd 'putevka' установлена и запущена на порту 8080!"
else
  fuser -k 8080/tcp 2>/dev/null || true
  pkill -9 -f "server.py" 2>/dev/null || true
  nohup "$TARGET_DIR/venv/bin/python" "$TARGET_DIR/server.py" > "$TARGET_DIR/server.log" 2>&1 &
  echo ">>> Сервер ПУТЁВКА v6.5.0 запущен в фоновом режиме на порту 8080!"
fi
