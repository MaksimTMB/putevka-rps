#!/usr/bin/env bash
set -e
PROTO="https""://"
REPO="MaksimTMB/putevka-rps"
RAW_BASE="${PROTO}raw.githubusercontent.com/${REPO}/main"
APP_DIR="${PUTEVKA_DIR:-$(pwd)}"
echo "=== АСУ ПУТЁВКА v7.1.1 - Прямое обновление с GitHub ==="
mkdir -p "$APP_DIR/data" "$APP_DIR/static" "$APP_DIR/templates_xls"
if [ -f "$APP_DIR/data/putevka_db.json" ]; then
    cp -f "$APP_DIR/data/putevka_db.json" "$APP_DIR/data/putevka_db.backup.json"
    echo "[OK] Создана резервная копия рабочей БД."
fi
for rel in "version.txt" "server.py" "desktop_app.py" "requirements.txt" "setup_and_run.sh" "static/index.html"; do
    curl -fsSL "$RAW_BASE/$rel" -o "$APP_DIR/$rel" 2>/dev/null && echo "[OK] Обновлён: $rel" || true
done
chmod +x "$APP_DIR/setup_and_run.sh" 2>/dev/null || true
echo "[OK] Обновление завершено!"
