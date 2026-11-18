#!/usr/bin/env bash
set -e
PROTO="https""://"
REPO="MaksimTMB/putevka-rps"
RAW_BASE="${PROTO}raw.githubusercontent.com/${REPO}/main"
APP_DIR="${PUTEVKA_DIR:-$(pwd)}"
echo "=== АСУ ПУТЁВКА v8.5.0 - Прямое обновление с GitHub ==="
mkdir -p "$APP_DIR/data" "$APP_DIR/static" "$APP_DIR/templates_xls"
if [ -f "$APP_DIR/data/putevka_db.json" ]; then
    cp -f "$APP_DIR/data/putevka_db.json" "$APP_DIR/data/putevka_db.backup.json"
    echo "[OK] Создана резервная копия рабочей БД."
fi
for rel in "version.txt" "server.py" "desktop_app.py" "requirements.txt" "setup_and_run.sh" "static/index.html"; do
    curl -fsSL "$RAW_BASE/$rel" -o "$APP_DIR/$rel" 2>/dev/null && echo "[OK] Обновлён: $rel" || true
done
chmod +x "$APP_DIR/setup_and_run.sh" 2>/dev/null || true
# v8.5.0: модели чтения ВУ с фото (~10 МБ) — один раз, если их ещё нет
mkdir -p "$APP_DIR/static/ocr"
for rel in "static/ocr/ppocr_det_v3.onnx" "static/ocr/ppocr_eslav_rec_v5.onnx" "static/ocr/ppocr_eslav_dict.txt" "static/ocr/NOTICE.txt"; do
    [ -s "$APP_DIR/$rel" ] && continue
    curl -fsSL "$RAW_BASE/$rel" -o "$APP_DIR/$rel.part" 2>/dev/null && mv -f "$APP_DIR/$rel.part" "$APP_DIR/$rel" && echo "[OK] Скачан: $rel" || true
done
echo "[OK] Обновление завершено!"
