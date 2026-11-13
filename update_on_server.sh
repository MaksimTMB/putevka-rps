#!/usr/bin/env bash
# ==============================================================================
# АСУ «ПУТЁВКА» v7.0.3 — Единый скрипт проверки окружения, полной установки,
# обновления, самодиагностики и запуска сервера (Linux / macOS)
# Использование:
#   bash setup_and_run.sh              # Установка/обновление (сброс БД к чистому релизу)
#   bash setup_and_run.sh --keep-db    # Обновление с сохранением текущей рабочей БД
#   bash setup_and_run.sh --port 8080  # Запуск на указанном порту (по умолчанию 8080)
# ==============================================================================
set -e

APP_VERSION="v7.0.3"
PORT="${PORT:-8080}"
KEEP_DB=0
USE_SYSTEMD=1

while [[ $# -gt 0 ]]; do
  case "$1" in
    --keep-db)
      KEEP_DB=1
      shift
      ;;
    --reset-db)
      KEEP_DB=0
      shift
      ;;
    --no-systemd)
      USE_SYSTEMD=0
      shift
      ;;
    --port)
      PORT="$2"
      shift 2
      ;;
    *)
      shift
      ;;
  esac
done

GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
CYAN='\033[0;36m'
NC='\033[0m'

info()  { echo -e "${CYAN}[INFO]${NC}  $*"; }
ok()    { echo -e "${GREEN}[ OK ]${NC}  $*"; }
warn()  { echo -e "${YELLOW}[WARN]${NC}  $*"; }
err()   { echo -e "${RED}[ERR ]${NC}  $*" >&2; }

SUDO=""
if [ "$(id -u)" -ne 0 ] && command -v sudo >/dev/null 2>&1; then
  SUDO="sudo"
fi

echo -e "${CYAN}==================================================================${NC}"
echo -e "${CYAN}  АСУ «ПУТЁВКА» ${APP_VERSION} — Проверка окружения, установка и запуск${NC}"
echo -e "${CYAN}==================================================================${NC}"

# ------------------------------------------------------------------------------
# ШАГ 1. Определение ОС и целевой директории проекта
# ------------------------------------------------------------------------------
OS_TYPE="$(uname -s)"
DISTRO_ID="unknown"
if [ -f /etc/os-release ]; then
  . /etc/os-release
  DISTRO_ID="${ID:-linux}"
fi
info "Операционная система: ${OS_TYPE} (${DISTRO_ID}), архитектура: $(uname -m)"

CURRENT_SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
if [ ! -f "$CURRENT_SCRIPT_DIR/server.py" ] && [ -f "$CURRENT_SCRIPT_DIR/../server.py" ]; then
  CURRENT_SCRIPT_DIR="$(cd "$CURRENT_SCRIPT_DIR/.." && pwd)"
fi

if [ "$OS_TYPE" = "Darwin" ]; then
  TARGET_DIR="$CURRENT_SCRIPT_DIR"
else
  TARGET_DIR="/opt/asu_waybills"
  if [ "$CURRENT_SCRIPT_DIR" = "$TARGET_DIR" ]; then
    TARGET_DIR="$CURRENT_SCRIPT_DIR"
  elif [ ! -d "$TARGET_DIR" ]; then
    FOUND=$(find /opt /home -maxdepth 3 -name "server.py" 2>/dev/null | grep -v "putevka_update" | head -n 1 || true)
    if [ -n "$FOUND" ]; then
      TARGET_DIR="$(dirname "$FOUND")"
    fi
  fi
fi
info "Рабочая директория сервера: ${TARGET_DIR}"

# ------------------------------------------------------------------------------
# ШАГ 2. Проверка и автоустановка системных пакетов (Python 3, venv, pip, утилиты)
# ------------------------------------------------------------------------------
install_sys_packages_linux() {
  info "Установка недостающих системных пакетов (${DISTRO_ID})..."
  if command -v apt-get >/dev/null 2>&1; then
    $SUDO apt-get update -y
    $SUDO apt-get install -y python3 python3-venv python3-pip curl unzip psmisc lsof || true
  elif command -v dnf >/dev/null 2>&1; then
    $SUDO dnf install -y python3 python3-pip curl unzip psmisc lsof || true
  elif command -v yum >/dev/null 2>&1; then
    $SUDO yum install -y python3 python3-pip curl unzip psmisc lsof || true
  elif command -v pacman >/dev/null 2>&1; then
    $SUDO pacman -Sy --noconfirm python python-pip curl unzip psmisc lsof || true
  elif command -v zypper >/dev/null 2>&1; then
    $SUDO zypper install -y python3 python3-pip curl unzip psmisc lsof || true
  elif command -v apk >/dev/null 2>&1; then
    $SUDO apk add --no-cache python3 py3-pip curl unzip psmisc lsof || true
  fi
}

NEED_SYS_INSTALL=0
if ! command -v python3 >/dev/null 2>&1; then
  warn "Пакет python3 не найден в системе."
  NEED_SYS_INSTALL=1
elif ! python3 -c "import venv, ensurepip" >/dev/null 2>&1; then
  warn "Модуль python3-venv (ensurepip) отсутствует."
  NEED_SYS_INSTALL=1
fi

if [ "$OS_TYPE" = "Linux" ] && [ "$NEED_SYS_INSTALL" -eq 1 ]; then
  install_sys_packages_linux
elif [ "$OS_TYPE" = "Darwin" ] && ! command -v python3 >/dev/null 2>&1; then
  if command -v brew >/dev/null 2>&1; then
    brew install python
  else
    err "Python 3 не установлен. Установите Python 3 с https://www.python.org/ или через Homebrew."
    exit 1
  fi
fi

PY_VER=$(python3 -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}")')
PY_MAJ=$(python3 -c 'import sys; print(sys.version_info.major)')
PY_MIN=$(python3 -c 'import sys; print(sys.version_info.minor)')
if [ "$PY_MAJ" -lt 3 ] || { [ "$PY_MAJ" -eq 3 ] && [ "$PY_MIN" -lt 8 ]; }; then
  err "Требуется Python >= 3.8, обнаружена версия ${PY_VER}"
  exit 1
fi
ok "Python проверен: версия ${PY_VER} ($(command -v python3))"

# ------------------------------------------------------------------------------
# ШАГ 3. Копирование файлов релиза и проверка целостности структуры проекта
# ------------------------------------------------------------------------------
$SUDO mkdir -p "$TARGET_DIR/data" "$TARGET_DIR/static" "$TARGET_DIR/templates_xls" "$TARGET_DIR/scripts"
if [ "$CURRENT_SCRIPT_DIR" != "$TARGET_DIR" ]; then
  info "Копирование файлов сборки ${APP_VERSION} из ${CURRENT_SCRIPT_DIR} в ${TARGET_DIR}..."
  $SUDO cp -rf "$CURRENT_SCRIPT_DIR"/* "$TARGET_DIR/"
fi

if [ -n "$SUDO" ]; then
  $SUDO chown -R "$(id -u):$(id -g)" "$TARGET_DIR" 2>/dev/null || true
fi

cd "$TARGET_DIR"

for req_file in "server.py" "requirements.txt" "static/index.html" "templates_xls/template_0345001.xlsx" "data/app_data.json"; do
  if [ ! -s "$TARGET_DIR/$req_file" ]; then
    err "Отсутствует или поврежден обязательный файл проекта: $TARGET_DIR/$req_file"
    exit 1
  fi
done
ok "Все обязательные файлы проекта проверены и находятся на месте."

# Инициализация или обновление файла базы данных putevka_db.json
if [ "$KEEP_DB" -eq 1 ] && [ -s "$TARGET_DIR/data/putevka_db.json" ]; then
  if python3 -c 'import json; d=json.load(open("data/putevka_db.json", encoding="utf-8")); assert isinstance(d.get("profiles"), list)' >/dev/null 2>&1; then
    ok "Сохранена существующая рабочая база данных data/putevka_db.json (--keep-db)."
  else
    warn "Существующий файл data/putevka_db.json поврежден, восстанавливаем из эталона data/app_data.json..."
    cp -f "$TARGET_DIR/data/app_data.json" "$TARGET_DIR/data/putevka_db.json"
  fi
else
  cp -f "$TARGET_DIR/data/app_data.json" "$TARGET_DIR/data/putevka_db.json"
  ok "База данных инициализирована эталоном релиза ${APP_VERSION} (активирована форма первичной настройки Администратора)."
fi
chmod -R 777 "$TARGET_DIR/data" 2>/dev/null || true

# ------------------------------------------------------------------------------
# ШАГ 4. Проверка/создание виртуального окружения venv и установка зависимостей
# ------------------------------------------------------------------------------
if [ ! -x "$TARGET_DIR/venv/bin/python" ] || ! "$TARGET_DIR/venv/bin/python" -m pip --version >/dev/null 2>&1; then
  info "Создание чистого виртуального окружения Python ($TARGET_DIR/venv)..."
  rm -rf "$TARGET_DIR/venv"
  if ! python3 -m venv "$TARGET_DIR/venv" 2>/dev/null; then
    warn "Попытка доустановки python3-venv..."
    [ "$OS_TYPE" = "Linux" ] && install_sys_packages_linux
    python3 -m venv "$TARGET_DIR/venv"
  fi
fi
ok "Виртуальное окружение venv готово: $TARGET_DIR/venv/bin/python"

info "Проверка и установка Python-зависимостей (fastapi, uvicorn, openpyxl, pydantic)..."
if ! "$TARGET_DIR/venv/bin/python" -c "import fastapi, uvicorn, openpyxl, pydantic" >/dev/null 2>&1; then
  "$TARGET_DIR/venv/bin/pip" install --upgrade pip setuptools wheel
  "$TARGET_DIR/venv/bin/pip" install -r "$TARGET_DIR/requirements.txt"
else
  "$TARGET_DIR/venv/bin/pip" install -q -r "$TARGET_DIR/requirements.txt" 2>/dev/null || true
fi

# Предстартовая самодиагностика импорта модулей и синтаксиса server.py
"$TARGET_DIR/venv/bin/python" -c "
import fastapi, uvicorn, openpyxl, pydantic
import server
assert server.APP_VERSION == '${APP_VERSION}', f'Version mismatch: {server.APP_VERSION}'
print('  -> Предстартовая проверка модулей успешно пройдена (FastAPI ' + fastapi.__version__ + ', OpenPyXL ' + openpyxl.__version__ + ', версия сервера ' + server.APP_VERSION + ')')
"
ok "Все зависимости и код сервера прошли предстартовую валидацию."

# ------------------------------------------------------------------------------
# ШАГ 5. Освобождение порта, настройка firewall и перезапуск сервера
# ------------------------------------------------------------------------------
info "Остановка прежних экземпляров сервера на порту ${PORT}..."
if [ "$OS_TYPE" = "Linux" ] && command -v systemctl >/dev/null 2>&1; then
  $SUDO systemctl stop putevka 2>/dev/null || true
fi
$SUDO fuser -k "${PORT}/tcp" 2>/dev/null || true
if command -v lsof >/dev/null 2>&1; then
  lsof -ti:"${PORT}" | xargs -r kill -9 2>/dev/null || true
fi
$SUDO pkill -9 -f "server.py" 2>/dev/null || true
sleep 1

# Открываем порт в фаерволе Linux, если он активен
if [ "$OS_TYPE" = "Linux" ]; then
  if command -v ufw >/dev/null 2>&1 && $SUDO ufw status 2>/dev/null | grep -q "Status: active"; then
    $SUDO ufw allow "${PORT}/tcp" >/dev/null 2>&1 || true
    ok "Порт ${PORT}/tcp открыт в UFW."
  elif command -v firewall-cmd >/dev/null 2>&1 && $SUDO firewall-cmd --state >/dev/null 2>&1; then
    $SUDO firewall-cmd --add-port="${PORT}/tcp" --permanent >/dev/null 2>&1 || true
    $SUDO firewall-cmd --reload >/dev/null 2>&1 || true
    ok "Порт ${PORT}/tcp открыт в firewalld."
  fi
fi

STARTED_VIA="nohup"
if [ "$OS_TYPE" = "Linux" ] && [ "$USE_SYSTEMD" -eq 1 ] && command -v systemctl >/dev/null 2>&1 && [ -d /run/systemd/system ]; then
  info "Регистрация и запуск системной службы systemd (putevka.service)..."
  cat <<SERVICE_EOF | $SUDO tee /etc/systemd/system/putevka.service >/dev/null
[Unit]
Description=ASU Putevka ${APP_VERSION} Waybill Management Server (Form 0345001)
After=network.target

[Service]
Type=simple
WorkingDirectory=${TARGET_DIR}
ExecStart=${TARGET_DIR}/venv/bin/python ${TARGET_DIR}/server.py
Restart=always
RestartSec=3
StandardOutput=append:${TARGET_DIR}/server.log
StandardError=append:${TARGET_DIR}/server.log

[Install]
WantedBy=multi-user.target
SERVICE_EOF
  $SUDO systemctl daemon-reload
  $SUDO systemctl enable putevka >/dev/null 2>&1 || true
  $SUDO systemctl restart putevka
  STARTED_VIA="systemd (putevka.service)"
else
  info "Запуск сервера в фоновом режиме (nohup)..."
  nohup "$TARGET_DIR/venv/bin/python" "$TARGET_DIR/server.py" > "$TARGET_DIR/server.log" 2>&1 &
fi

# ------------------------------------------------------------------------------
# ШАГ 6. Контрольная проверка HTTP-ответа запущенного сервера (Health Check)
# ------------------------------------------------------------------------------
info "Ожидание готовности HTTP-сервера на порту ${PORT}..."
HEALTH_OK=0
for attempt in 1 2 3 4 5 6; do
  sleep 1
  if "$TARGET_DIR/venv/bin/python" -c "
import urllib.request, json
with urllib.request.urlopen('http://127.0.3.1:${PORT}/api/state', timeout=3) as resp:
    assert resp.status == 200
    data = json.loads(resp.read().decode('utf-8'))
    assert isinstance(data.get('profiles'), list)
" >/dev/null 2>&1; then
    HEALTH_OK=1
    break
  fi
done

if [ "$HEALTH_OK" -ne 1 ]; then
  err "Сервер не ответил на контрольный HTTP-запрос! Вывод лога $TARGET_DIR/server.log:"
  cat "$TARGET_DIR/server.log" || true
  exit 1
fi

PID_VAL=$(pgrep -f "server.py" | head -n 1 || echo "активен")
LAN_IP=$(hostname -I 2>/dev/null | awk '{print $1}' || ipconfig getifaddr en0 2>/dev/null || echo "127.0.3.1")

echo -e "${GREEN}==================================================================${NC}"
echo -e "${GREEN}  >>> УСПЕШНО! АСУ «ПУТЁВКА» ${APP_VERSION} установлена и запущена!${NC}"
echo -e "${GREEN}==================================================================${NC}"
echo -e "  Версия сборки:      ${GREEN}${APP_VERSION}${NC}"
echo -e "  Режим запуска:      ${STARTED_VIA} (PID: ${PID_VAL})"
echo -e "  Рабочая папка:      ${TARGET_DIR}"
echo -e "  Файл базы данных:   ${TARGET_DIR}/data/putevka_db.json"
echo -e "  Локальный адрес:    ${CYAN}http://127.0.3.1:${PORT}${NC}"
echo -e "  Адрес в сети (LAN): ${CYAN}http://${LAN_IP}:${PORT}${NC}"
echo -e "  Лог-файл сервера:   ${TARGET_DIR}/server.log"
echo -e "${GREEN}==================================================================${NC}"
