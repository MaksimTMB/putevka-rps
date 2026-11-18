#!/usr/bin/env bash
set -e
PROTO="https""://"
REPO="MaksimTMB/putevka-rps"
RAW_BASE="${PROTO}raw.githubusercontent.com/${REPO}/main"
APP_DIR="${PUTEVKA_DIR:-$(pwd)}"
echo "=== АСУ ПУТЁВКА v8.5.5 - Прямое обновление с GitHub ==="
mkdir -p "$APP_DIR/data" "$APP_DIR/static" "$APP_DIR/templates_xls"
if [ -f "$APP_DIR/data/putevka_db.json" ]; then
    cp -f "$APP_DIR/data/putevka_db.json" "$APP_DIR/data/putevka_db.backup.json"
    echo "[OK] Создана резервная копия рабочей БД."
fi
for rel in "version.txt" "server.py" "desktop_app.py" "requirements.txt" "setup_and_run.sh" "static/index.html"; do
    curl -fsSL "$RAW_BASE/$rel" -o "$APP_DIR/$rel" 2>/dev/null && echo "[OK] Обновлён: $rel" || true
done
chmod +x "$APP_DIR/setup_and_run.sh" 2>/dev/null || true
# v8.5.0/8.5.2: модели чтения ВУ и библиотеки страницы — один раз, если их ещё нет
mkdir -p "$APP_DIR/static/ocr" "$APP_DIR/static/vendor"
for rel in "static/ocr/ppocr_det_v3.onnx" "static/ocr/ppocr_eslav_rec_v5.onnx" "static/ocr/ppocr_eslav_dict.txt" "static/ocr/NOTICE.txt" "static/vendor/tailwind-3.4.17.js" "static/vendor/chart-4.4.1.umd.min.js" "static/vendor/xlsx-0.20.3.full.min.js" "static/vendor/pdf-3.11.174.min.js" "static/vendor/pdf.worker-3.11.174.min.js" "static/vendor/NOTICE.txt"; do
    [ -s "$APP_DIR/$rel" ] && continue
    curl -fsSL "$RAW_BASE/$rel" -o "$APP_DIR/$rel.part" 2>/dev/null && mv -f "$APP_DIR/$rel.part" "$APP_DIR/$rel" && echo "[OK] Скачан: $rel" || true
done
echo "[OK] Обновление завершено!"
