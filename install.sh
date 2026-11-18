#!/usr/bin/env bash
# АСУ «ПУТЁВКА» — сервер на Linux: установка, обновление и удаление одной командой.
#
# Установить или обновить до последней версии:
#   curl -fsSL https://raw.githubusercontent.com/MaksimTMB/putevka-rps/main/install.sh | sudo bash
#
# Ключи пишутся после «bash -s --», например:
#   curl -fsSL https://raw.githubusercontent.com/MaksimTMB/putevka-rps/main/install.sh | sudo bash -s -- --port 8090
#
#   --port N           порт (по умолчанию 8080; при обновлении остаётся прежний)
#   --dir ПАПКА        папка программы (по умолчанию /opt/putevka; при обновлении — прежняя)
#   --version vX.Y.Z   поставить эту версию, а не последнюю
#   --source ПУТЬ      ставить из скачанного архива исходников (.tar.gz) или папки, без GitHub
#   --no-firewall      не открывать порт в ufw / firewalld
#   --uninstall        удалить службу и программу; папка data (база и документы) остаётся
#   --purge            вместе с --uninstall: удалить и data (последняя копия — в /var/backups/putevka)
#
# Что делает: ставит Python 3 (если его нет), скачивает выпуск с GitHub, заводит пользователя
# putevka и службу systemd «putevka», запускает и проверяет, что сервер отвечает.
# Повторный запуск — обновление: база (data/) не трогается, перед заменой файлов делается копия,
# а если новая версия не запустилась, прежняя возвращается сама.

set -Eeuo pipefail
umask 022

REPO="MaksimTMB/putevka-rps"
RAW="https://raw.githubusercontent.com/${REPO}/main"
NAME="putevka"
UNIT="/etc/systemd/system/putevka.service"
ENV_FILE="/etc/putevka.env"
BACKUP_ROOT="/var/backups/putevka"
DEFAULT_DIR="/opt/putevka"
INSTALL_CMD="curl -fsSL ${RAW}/install.sh | sudo bash"

PORT_ARG="" DIR_ARG="" VERSION_ARG="" SOURCE_ARG="" FIREWALL=1 MODE=install PURGE=0
APP_DIR="" PORT="" TAG="" OLD_DIR="" OLD_PORT="" OLD_VER="" UPGRADE=0 BK="" WORK="" SRC="" PY="" SCHEME=http
OS_ID=""

if [ -t 1 ]; then B=$'\033[1m' G=$'\033[32m' Y=$'\033[33m' R=$'\033[31m' C=$'\033[36m' N=$'\033[0m'; else B="" G="" Y="" R="" C="" N=""; fi
step() { printf '\n%s==> %s%s\n' "$C$B" "$*" "$N"; }
ok()   { printf '  %s✓%s %s\n' "$G" "$N" "$*"; }
note() { printf '  %s\n' "$*"; }
warn() { printf '  %s!%s %s\n' "$Y" "$N" "$*" >&2; }
die()  { printf '\n%sОшибка:%s %s\n' "$R$B" "$N" "$*" >&2; exit 1; }
cleanup() { [ -n "$WORK" ] && rm -rf "$WORK"; return 0; }
on_err() { printf '\n%sУстановка прервана%s (строка %s). Повторите команду; если ошибка повторится — пришлите этот вывод.\n' "$R$B" "$N" "$1" >&2; }

usage() {
	cat <<EOF
АСУ «Путёвка» — сервер на Linux: установка, обновление, удаление.

  Установить или обновить:  $INSTALL_CMD
  С ключами:                $INSTALL_CMD -s -- --port 8090

  --port N           порт (по умолчанию 8080; при обновлении остаётся прежний)
  --dir ПАПКА        папка программы (по умолчанию $DEFAULT_DIR; при обновлении — прежняя)
  --version vX.Y.Z   поставить эту версию, а не последнюю
  --source ПУТЬ      ставить из скачанного архива исходников (.tar.gz) или папки, без GitHub
  --no-firewall      не открывать порт в ufw / firewalld
  --uninstall        удалить службу и программу; папка data (база и документы) остаётся
  --purge            вместе с --uninstall: удалить и data (последняя копия — в $BACKUP_ROOT)
EOF
}

parse_args() {
	while [ $# -gt 0 ]; do
		case "$1" in
			--port) [ $# -ge 2 ] || die "после --port нужен номер порта"; PORT_ARG="$2"; shift 2 ;;
			--port=*) PORT_ARG="${1#*=}"; shift ;;
			--dir) [ $# -ge 2 ] || die "после --dir нужна папка"; DIR_ARG="$2"; shift 2 ;;
			--dir=*) DIR_ARG="${1#*=}"; shift ;;
			--version) [ $# -ge 2 ] || die "после --version нужен номер версии"; VERSION_ARG="$2"; shift 2 ;;
			--version=*) VERSION_ARG="${1#*=}"; shift ;;
			--source) [ $# -ge 2 ] || die "после --source нужен архив или папка"; SOURCE_ARG="$2"; shift 2 ;;
			--source=*) SOURCE_ARG="${1#*=}"; shift ;;
			--no-firewall) FIREWALL=0; shift ;;
			--uninstall) MODE=uninstall; shift ;;
			--purge) PURGE=1; shift ;;
			-h|--help) usage; exit 0 ;;
			*) die "неизвестный ключ «$1». Список ключей: $INSTALL_CMD -s -- --help" ;;
		esac
	done
	if [ "$PURGE" -eq 1 ] && [ "$MODE" != uninstall ]; then die "--purge работает только вместе с --uninstall"; fi
	return 0
}

need_root() {
	[ "$(id -u)" -eq 0 ] && return 0
	# запуск файлом (bash install.sh) — перезапускаемся через sudo; из «curl | bash» так нельзя
	if [ -f "$0" ] && grep -q 'АСУ «ПУТЁВКА» — сервер на Linux' "$0" 2>/dev/null && command -v sudo >/dev/null 2>&1; then
		exec sudo bash "$0" "$@"
	fi
	die "нужны права администратора. Запустите так: $INSTALL_CMD"
}

check_system() {
	[ "$(uname -s)" = Linux ] || die "этот установщик — для Linux. Для Windows скачайте Putevka.exe: https://github.com/$REPO/releases/latest"
	[ -d /run/systemd/system ] || die "нужна система с systemd (Ubuntu, Debian, Astra, ALT, РЕД ОС, Rocky, AlmaLinux…)"
	local pretty="Linux"
	if [ -r /etc/os-release ]; then
		pretty="$(. /etc/os-release && printf '%s' "${PRETTY_NAME:-Linux}")"
		OS_ID="$(. /etc/os-release && printf '%s' "${ID:-}")"
	fi
	ok "$pretty, $(uname -m)"
}

# ---------- Python и системные пакеты ----------
venv_works() {
	local t r=0
	t="$(mktemp -d)"
	"$1" -m venv "$t/v" >/dev/null 2>&1 && [ -x "$t/v/bin/pip" ] || r=1
	rm -rf "$t"
	return $r
}

pick_python() {
	local c p
	for c in python3 python3.13 python3.12 python3.11 python3.10 python3.9 python3.8; do
		p="$(command -v "$c" 2>/dev/null)" || continue
		"$p" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 8) else 1)' 2>/dev/null || continue
		venv_works "$p" || continue
		PY="$p"
		return 0
	done
	return 1
}

install_packages() {
	note "ставлю недостающее: Python 3 с venv, curl…"
	if command -v apt-get >/dev/null 2>&1; then
		export DEBIAN_FRONTEND=noninteractive
		apt-get update -qq >/dev/null || warn "apt-get update не прошёл — ставлю из имеющихся списков"
		if [ "$OS_ID" = altlinux ]; then
			apt-get install -y python3 python3-module-pip curl ca-certificates tar >/dev/null || true
		else
			apt-get install -y -qq python3 python3-venv python3-pip curl ca-certificates tar >/dev/null || true
		fi
	elif command -v dnf >/dev/null 2>&1; then
		dnf install -y -q python3 python3-pip curl ca-certificates tar >/dev/null || true
	elif command -v yum >/dev/null 2>&1; then
		yum install -y -q python3 python3-pip curl ca-certificates tar >/dev/null || true
	elif command -v zypper >/dev/null 2>&1; then
		zypper --non-interactive --quiet install python3 python3-pip curl ca-certificates tar >/dev/null || true
	elif command -v pacman >/dev/null 2>&1; then
		pacman -Sy --noconfirm --needed python python-pip curl ca-certificates tar >/dev/null || true
	else
		warn "менеджер пакетов не распознан — поставьте Python 3.8+ (с модулем venv) и curl вручную"
	fi
	return 0
}

install_newer_python() {
	# системный Python старше 3.8 (RHEL 8, старые Astra) — пробуем более новый рядом
	note "системный Python старше 3.8 — ищу более новый в пакетах…"
	local v
	if command -v dnf >/dev/null 2>&1; then
		dnf install -y -q python3.11 >/dev/null 2>&1 || dnf install -y -q python39 >/dev/null 2>&1 || true
	elif command -v yum >/dev/null 2>&1; then
		yum install -y -q python3.11 >/dev/null 2>&1 || yum install -y -q python39 >/dev/null 2>&1 || true
	elif command -v apt-get >/dev/null 2>&1; then
		for v in 3.12 3.11 3.10 3.9 3.8; do
			apt-get install -y -qq "python$v" "python$v-venv" >/dev/null 2>&1 && break
		done
	fi
	return 0
}

ensure_prereqs() {
	local missing=0
	command -v curl >/dev/null 2>&1 || missing=1
	command -v tar >/dev/null 2>&1 || missing=1
	pick_python || missing=1
	if [ "$missing" -eq 1 ]; then
		install_packages
		command -v curl >/dev/null 2>&1 || die "не удалось поставить curl — поставьте его пакетом системы и запустите снова"
		command -v tar >/dev/null 2>&1 || die "не удалось поставить tar — поставьте его пакетом системы и запустите снова"
		if ! pick_python; then
			install_newer_python
			pick_python || die "нужен Python 3.8 или новее с модулем venv. Поставьте его пакетами системы и запустите установку снова"
		fi
	fi
	ok "Python $("$PY" -c 'import platform; print(platform.python_version())') ($PY)"
}

# ---------- выпуск ----------
fetch() { curl -fsSL --retry 3 --retry-delay 2 --connect-timeout 20 "$@"; }

resolve_version() {
	[ -n "$SOURCE_ARG" ] && return 0
	if [ -n "$VERSION_ARG" ]; then
		TAG="v${VERSION_ARG#v}"
	else
		TAG="$(fetch "https://api.github.com/repos/$REPO/releases/latest" 2>/dev/null | sed -n 's/^ *"tag_name": *"\([^"]*\)".*/\1/p' | sed -n 1p)" || TAG=""
		if [ -z "$TAG" ]; then
			TAG="v$(fetch "$RAW/version.txt" 2>/dev/null | tr -d '[:space:]')" || TAG=""
		fi
	fi
	[[ "$TAG" =~ ^v[0-9]+\.[0-9]+\.[0-9]+$ ]] || die "не удалось узнать версию на GitHub (нет связи с github.com?). Можно поставить из скачанного архива: --source"
	return 0
}

fetch_source() {
	WORK="$(mktemp -d /tmp/putevka-install.XXXXXX)"
	if [ -n "$SOURCE_ARG" ]; then
		if [ -d "$SOURCE_ARG" ]; then
			SRC="$(cd "$SOURCE_ARG" && pwd)"
		elif [ -f "$SOURCE_ARG" ]; then
			tar -xzf "$SOURCE_ARG" -C "$WORK" || die "не распаковался архив $SOURCE_ARG"
			SRC="$(find "$WORK" -mindepth 1 -maxdepth 1 -type d | sed -n 1p)"
		else
			die "нет такого файла или папки: $SOURCE_ARG"
		fi
	else
		note "скачиваю $TAG с GitHub…"
		fetch -o "$WORK/src.tar.gz" "https://github.com/$REPO/archive/refs/tags/$TAG.tar.gz" \
			|| die "не удалось скачать $TAG с GitHub (нет такой версии или нет связи с github.com)"
		tar -xzf "$WORK/src.tar.gz" -C "$WORK" || die "архив $TAG повреждён — запустите установку ещё раз"
		SRC="$(find "$WORK" -mindepth 1 -maxdepth 1 -type d | sed -n 1p)"
	fi
	local f v
	for f in server.py version.txt requirements.txt static/index.html templates_xls/template_0345001.xlsx data/app_data.json; do
		[ -s "$SRC/$f" ] || die "в выпуске нет файла $f — это не исходники «Путёвки»?"
	done
	v="v$(tr -d '[:space:]' < "$SRC/version.txt")"
	if [ -n "$SOURCE_ARG" ]; then TAG="$v"; elif [ "$v" != "$TAG" ]; then die "внутри архива $TAG оказалась версия $v"; fi
	"$PY" -c 'import ast, sys; ast.parse(open(sys.argv[1], encoding="utf-8").read())' "$SRC/server.py" 2>/dev/null \
		|| die "server.py версии $TAG не разбирается этим Python ($PY) — нужен более новый Python"
	ok "выпуск $TAG получен и проверен"
}

# ---------- куда ставим ----------
bad_dir() {
	case "$1" in
		/|/bin|/boot|/dev|/etc|/home|/lib|/lib64|/media|/mnt|/opt|/proc|/root|/run|/sbin|/srv|/sys|/tmp|/usr|/usr/*|/var|/var/lib|/var/log) return 0 ;;
	esac
	[[ "$1" =~ ^/[A-Za-z0-9._/-]+$ ]] || return 0
	return 1
}

detect_existing() {
	if [ -f "$UNIT" ]; then
		OLD_DIR="$(sed -n 's/^WorkingDirectory=//p' "$UNIT" | sed -n 1p)"
		OLD_PORT="$(sed -n 's/^Environment=PUTEVKA_PORT=//p' "$UNIT" | sed -n 1p)"
	fi
	if [ -n "$DIR_ARG" ]; then
		case "$DIR_ARG" in /*) ;; *) die "--dir: нужен полный путь, например /opt/putevka" ;; esac
		APP_DIR="${DIR_ARG%/}"
		if [ -n "$OLD_DIR" ] && [ "$OLD_DIR" != "$APP_DIR" ]; then
			die "«Путёвка» уже установлена в $OLD_DIR. Обновляйте без --dir или сначала удалите прежнюю: --uninstall"
		fi
	else
		APP_DIR="${OLD_DIR:-$DEFAULT_DIR}"
	fi
	bad_dir "$APP_DIR" && die "в $APP_DIR ставить нельзя — укажите отдельную папку (латиница, цифры, «-», «_», «.»), например --dir /opt/putevka"
	[ -f "$APP_DIR/server.py" ] && UPGRADE=1
	if [ "$UPGRADE" -eq 1 ] && [ -s "$APP_DIR/version.txt" ]; then OLD_VER="v$(tr -d '[:space:]' < "$APP_DIR/version.txt")"; fi
	if [ "$MODE" = install ] && [ "$UPGRADE" -eq 0 ] && [ -d "$APP_DIR" ] && [ ! -d "$APP_DIR/data" ] && [ -n "$(ls -A "$APP_DIR" 2>/dev/null)" ]; then
		die "папка $APP_DIR не пуста, и в ней нет «Путёвки». Укажите другую: --dir"
	fi
	PORT="${PORT_ARG:-${OLD_PORT:-8080}}"
	if ! [[ "$PORT" =~ ^[0-9]+$ ]] || [ "$PORT" -lt 1 ] || [ "$PORT" -gt 65535 ]; then die "неверный порт: $PORT"; fi
	return 0
}

check_port() {
	if [ "$PORT" = "$OLD_PORT" ] && systemctl is-active --quiet "$NAME"; then return 0; fi
	if ! command -v ss >/dev/null 2>&1; then warn "нет утилиты ss — не проверил, свободен ли порт $PORT"; return 0; fi
	local busy who
	busy="$(ss -Hltnp "sport = :$PORT" 2>/dev/null | sed -n 1p)" || busy=""
	[ -z "$busy" ] && return 0
	who="$(printf '%s' "$busy" | grep -o 'users:(("[^"]*"' | cut -d'"' -f2)" || who=""
	die "порт $PORT уже занят${who:+ (программа «$who»)}. Выберите другой: $INSTALL_CMD -s -- --port 8090"
}

# ---------- установка ----------
backup_existing() {
	[ "$UPGRADE" -eq 1 ] || return 0
	BK="$BACKUP_ROOT/$(date +%Y%m%d-%H%M%S)"
	install -d -m 700 "$BACKUP_ROOT" "$BK"
	tar -C "$APP_DIR" --exclude=./data --exclude=./venv --exclude=./server.log -czf "$BK/program.tar.gz" . \
		|| die "не удалось сохранить копию программы в $BK — обновление не начато"
	if [ -f "$UNIT" ]; then cp -p "$UNIT" "$BK/putevka.service"; fi
	if [ -d "$APP_DIR/data" ]; then
		local need avail
		need="$(du -sk "$APP_DIR/data" | cut -f1)"
		avail="$(df -Pk "$BACKUP_ROOT" | awk 'NR==2 {print $4}')"
		if [ "$need" -lt $((avail - 204800)) ]; then
			tar -C "$APP_DIR" -czf "$BK/data.tar.gz" data || die "не удалось сохранить копию данных в $BK — обновление не начато"
		else
			tar -C "$APP_DIR" --exclude=data/documents --exclude=data/backups -czf "$BK/data.tar.gz" data \
				|| die "не удалось сохранить копию базы в $BK — обновление не начато"
			warn "места мало: в копию вошли база и настройки, без папки документов"
		fi
	fi
	ok "копия прежней версии и данных: $BK"
}

setup_venv() {
	local venv="$APP_DIR/venv"
	if [ -x "$venv/bin/python" ] && "$venv/bin/python" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 8) else 1)' 2>/dev/null \
		&& "$venv/bin/python" -m pip --version >/dev/null 2>&1; then
		note "окружение Python уже есть — сверяю зависимости…"
	else
		rm -rf "$venv"
		install -d "$APP_DIR"
		"$PY" -m venv "$venv" || die "не удалось создать окружение Python в $venv"
		note "ставлю зависимости (FastAPI, Uvicorn, OpenPyXL)…"
	fi
	"$venv/bin/python" -m pip install --disable-pip-version-check --no-input -q -r "$SRC/requirements.txt" \
		|| die "не удалось скачать зависимости Python с pypi.org. Если он недоступен, укажите зеркало: curl -fsSL $RAW/install.sh | sudo PIP_INDEX_URL=https://<зеркало>/simple bash"
	"$venv/bin/python" -c 'import fastapi, uvicorn, openpyxl' 2>/dev/null || die "зависимости Python поставились с ошибкой"
	ok "окружение Python готово"
}

ensure_user() {
	if id -u "$NAME" >/dev/null 2>&1; then return 0; fi
	useradd --system --user-group --home-dir "$APP_DIR" --no-create-home \
		--shell "$(command -v nologin 2>/dev/null || echo /bin/false)" --comment "ASU Putevka server" "$NAME" \
		|| die "не удалось создать пользователя $NAME"
	ok "создан служебный пользователь $NAME (без входа в систему)"
}

replace_dir() {
	rm -rf "${APP_DIR:?}/$1.new"
	cp -R "$SRC/$1" "$APP_DIR/$1.new"
	rm -rf "${APP_DIR:?}/$1"
	mv "$APP_DIR/$1.new" "$APP_DIR/$1"
}

install_files() {
	install -d "$APP_DIR" "$APP_DIR/data"
	local f
	for f in "$SRC"/*.py; do install -m 644 "$f" "$APP_DIR/"; done
	install -m 644 "$SRC/version.txt" "$SRC/requirements.txt" "$APP_DIR/"
	replace_dir static
	replace_dir templates_xls
	install -m 644 "$SRC/data/app_data.json" "$APP_DIR/data/app_data.json"
	if [ ! -s "$APP_DIR/data/putevka_db.json" ]; then
		install -m 640 "$SRC/data/app_data.json" "$APP_DIR/data/putevka_db.json"
	fi
	find "$APP_DIR" -maxdepth 3 -name '*.updtmp' -type f -delete 2>/dev/null || true
	# права: программа и данные — пользователю putevka (кнопка «Обновить» пишет файлы программы),
	# окружение Python — только root; данные закрыты от остальных
	chown "$NAME:$NAME" "$APP_DIR"
	chmod 750 "$APP_DIR"
	find "$APP_DIR" -mindepth 1 -maxdepth 1 ! -name venv -exec chown -R "$NAME:$NAME" {} +
	chown -R root:root "$APP_DIR/venv"
	find "$APP_DIR/data" -type d -exec chmod 750 {} +
	find "$APP_DIR/data" -type f -exec chmod 640 {} +
	ok "файлы $TAG в $APP_DIR, данные в $APP_DIR/data"
}

write_env_template() {
	[ -f "$ENV_FILE" ] && return 0
	cat > "$ENV_FILE" <<'EOF'
# Настройки сервера «Путёвки». После правки: systemctl restart putevka
#
# HTTPS прямо в программе (если она не стоит за прокси с сертификатом).
# Файлы должны быть доступны на чтение пользователю putevka.
#PUTEVKA_SSL_CERTFILE=/etc/ssl/putevka/fullchain.pem
#PUTEVKA_SSL_KEYFILE=/etc/ssl/putevka/privkey.pem
EOF
	chown "root:$NAME" "$ENV_FILE"
	chmod 640 "$ENV_FILE"
}

write_unit() {
	local sandbox="$1"
	{
		echo "# Создано установщиком «Путёвки» (install.sh); при обновлении перезаписывается."
		echo "# Свои настройки — в $ENV_FILE."
		echo "[Unit]"
		echo "Description=АСУ «Путёвка» — сервер путевых листов"
		echo "Documentation=https://github.com/$REPO"
		echo "After=network-online.target"
		echo "Wants=network-online.target"
		echo
		echo "[Service]"
		echo "Type=simple"
		echo "User=$NAME"
		echo "Group=$NAME"
		echo "WorkingDirectory=$APP_DIR"
		echo "Environment=PUTEVKA_HOST=0.0.0.0"
		echo "Environment=PUTEVKA_PORT=$PORT"
		echo "EnvironmentFile=-$ENV_FILE"
		echo "ExecStart=$APP_DIR/venv/bin/python $APP_DIR/server.py"
		echo "Restart=always"
		echo "RestartSec=3"
		echo "UMask=0027"
		if [ "$sandbox" = 1 ]; then
			echo "NoNewPrivileges=yes"
			echo "PrivateTmp=yes"
			echo "ProtectSystem=strict"
			echo "ReadWritePaths=$APP_DIR"
			echo "ProtectHome=read-only"
			echo "ProtectKernelTunables=yes"
			echo "ProtectControlGroups=yes"
			echo "RestrictSUIDSGID=yes"
		fi
		echo
		echo "[Install]"
		echo "WantedBy=multi-user.target"
	} > "$UNIT"
	chmod 644 "$UNIT"
}

# ---------- запуск и проверка ----------
api_version() {
	curl -fsSk --max-time 3 "$SCHEME://127.0.0.1:$PORT/api/version" 2>/dev/null || true
}

wait_ready() {
	local want="$1" _ v
	for _ in $(seq 1 60); do
		v="$(api_version | sed -n 's/.*"version" *: *"\([^"]*\)".*/\1/p')"
		[ "$v" = "$want" ] && return 0
		sleep 1
	done
	return 1
}

start_service() {
	if grep -q '^[[:space:]]*PUTEVKA_SSL_CERTFILE=' "$ENV_FILE" 2>/dev/null; then SCHEME=https; fi
	systemctl daemon-reload
	systemctl enable "$NAME" >/dev/null 2>&1 || true
	systemctl restart "$NAME" || true
	sleep 2
	# в некоторых контейнерах (LXC без вложенности) systemd не может изолировать службу — код 226
	if [ "$(systemctl show -p ExecMainStatus --value "$NAME" 2>/dev/null)" = 226 ]; then
		warn "система не даёт изолировать службу (контейнер?) — запускаю без дополнительных ограничений"
		write_unit 0
		systemctl daemon-reload
		systemctl restart "$NAME" || true
	fi
}

rollback() {
	warn "новая версия не ответила — возвращаю прежнюю $OLD_VER"
	systemctl stop "$NAME" 2>/dev/null || true
	rm -rf "${APP_DIR:?}/static" "${APP_DIR:?}/templates_xls"
	tar -C "$APP_DIR" -xzf "$BK/program.tar.gz" || { warn "не удалось распаковать копию $BK/program.tar.gz"; return 0; }
	if [ -f "$BK/putevka.service" ]; then
		cp -p "$BK/putevka.service" "$UNIT"
		systemctl daemon-reload
		systemctl restart "$NAME" || true
		if wait_ready "$OLD_VER"; then warn "прежняя версия $OLD_VER снова работает"; else warn "прежняя версия тоже не отвечает: journalctl -u $NAME -n 50"; fi
	else
		rm -f "$UNIT"
		systemctl daemon-reload
		warn "файлы прежней версии возвращены; службы у неё не было — запустите её так же, как раньше"
	fi
	return 0
}

open_firewall() {
	[ "$FIREWALL" -eq 1 ] || return 0
	local st
	if command -v ufw >/dev/null 2>&1; then
		st="$(ufw status 2>/dev/null || true)"
		case "$st" in
			*"Status: active"*|*"Состояние: активен"*)
				if ufw allow "$PORT/tcp" comment 'Putevka' >/dev/null 2>&1; then ok "порт $PORT открыт в ufw"; fi
				return 0 ;;
		esac
	fi
	if command -v firewall-cmd >/dev/null 2>&1 && firewall-cmd --state >/dev/null 2>&1; then
		if firewall-cmd --permanent --add-port="$PORT/tcp" >/dev/null 2>&1 && firewall-cmd --reload >/dev/null 2>&1; then ok "порт $PORT открыт в firewalld"; fi
	fi
	return 0
}

close_firewall() {
	local p="$1"
	[ -n "$p" ] || return 0
	if command -v ufw >/dev/null 2>&1; then ufw delete allow "$p/tcp" >/dev/null 2>&1 || true; fi
	if command -v firewall-cmd >/dev/null 2>&1 && firewall-cmd --state >/dev/null 2>&1; then
		firewall-cmd --permanent --remove-port="$p/tcp" >/dev/null 2>&1 && firewall-cmd --reload >/dev/null 2>&1 || true
	fi
	return 0
}

summary() {
	local ips ip setup
	ips="$(hostname -I 2>/dev/null || true)"
	[ -n "$ips" ] || ips="$(ip -4 -o addr show scope global 2>/dev/null | awk '{print $4}' | cut -d/ -f1 | tr '\n' ' ' || true)"
	setup="$(api_version | sed -n 's/.*"adminSetupCompleted" *: *\(true\|false\).*/\1/p')"
	printf '\n%s«Путёвка» %s работает.%s\n\n' "$G$B" "$TAG" "$N"
	echo "  Откройте в браузере:"
	for ip in $ips; do
		case "$ip" in *:*) continue ;; esac
		echo "    $SCHEME://$ip:$PORT"
	done
	echo "    $SCHEME://127.0.0.1:$PORT   (на самом сервере)"
	if [ "$setup" = false ]; then
		echo
		printf '  %sСразу откройте адрес и создайте администратора:%s пока его нет,\n' "$Y$B" "$N"
		echo "  первую настройку может пройти любой, кто откроет страницу."
	fi
	echo
	echo "  Программа:   $APP_DIR"
	echo "  База и документы: $APP_DIR/data"
	if [ -n "$BK" ]; then echo "  Копия до обновления: $BK"; fi
	echo
	echo "  Обновить:    кнопка «Обновить» в программе или эта же команда ещё раз"
	echo "  Состояние:   systemctl status $NAME"
	echo "  Журнал:      journalctl -u $NAME -f"
	echo "  Настройки:   $ENV_FILE (HTTPS), порт и папка — ключи установщика (--help)"
	echo "  Удалить:     $INSTALL_CMD -s -- --uninstall"
	echo
}

install_or_update() {
	step "Система"
	check_system
	ensure_prereqs

	step "Выпуск"
	resolve_version
	fetch_source
	if [ "$UPGRADE" -eq 1 ]; then
		if [ "$OLD_VER" = "$TAG" ]; then note "переустанавливаю $TAG в $APP_DIR"; else note "обновляю ${OLD_VER:-прежнюю версию} → $TAG в $APP_DIR"; fi
	else
		note "ставлю $TAG в $APP_DIR, порт $PORT"
	fi

	step "Установка"
	check_port
	backup_existing
	setup_venv
	ensure_user
	if [ -f "$UNIT" ]; then systemctl stop "$NAME" 2>/dev/null || true; fi
	install_files
	write_env_template
	write_unit 1

	step "Запуск"
	start_service
	if ! wait_ready "$TAG"; then
		echo >&2
		journalctl -u "$NAME" -n 30 --no-pager >&2 2>/dev/null || true
		if [ -n "$BK" ]; then
			rollback
			die "обновление до $TAG не удалось. Копия прежней версии и данных: $BK"
		fi
		die "сервер не ответил на порту $PORT. Журнал: journalctl -u $NAME -n 50"
	fi
	ok "служба $NAME запущена, сервер отвечает: $TAG"
	open_firewall
	summary
}

uninstall() {
	step "Удаление"
	if [ -z "$DIR_ARG" ] && [ -z "$OLD_DIR" ]; then die "служба $NAME не найдена. Если программа стоит в своей папке, укажите её: --dir"; fi
	[ -f "$APP_DIR/server.py" ] || [ -d "$APP_DIR/data" ] || die "в $APP_DIR нет «Путёвки»"
	systemctl disable --now "$NAME" >/dev/null 2>&1 || true
	rm -f "$UNIT"
	systemctl daemon-reload
	systemctl reset-failed "$NAME" >/dev/null 2>&1 || true
	close_firewall "$OLD_PORT"
	ok "служба $NAME остановлена и удалена"
	if [ "$PURGE" -eq 1 ]; then
		if [ -d "$APP_DIR/data" ]; then
			BK="$BACKUP_ROOT/$(date +%Y%m%d-%H%M%S)-removed"
			install -d -m 700 "$BACKUP_ROOT" "$BK"
			if tar -C "$APP_DIR" -czf "$BK/data.tar.gz" data; then ok "последняя копия данных: $BK/data.tar.gz (не нужна — удалите её)"; fi
		fi
		rm -rf --one-file-system "${APP_DIR:?}"
		rm -f "$ENV_FILE"
		if id -u "$NAME" >/dev/null 2>&1; then userdel "$NAME" >/dev/null 2>&1 || true; fi
		if getent group "$NAME" >/dev/null 2>&1; then groupdel "$NAME" >/dev/null 2>&1 || true; fi
		ok "программа и данные удалены"
	else
		find "$APP_DIR" -mindepth 1 -maxdepth 1 ! -name data -exec rm -rf {} +
		ok "программа удалена; база и документы остались в $APP_DIR/data"
		if [ "$APP_DIR" = "$DEFAULT_DIR" ]; then
			note "поставить снова с этими данными: $INSTALL_CMD"
		else
			note "поставить снова с этими данными: $INSTALL_CMD -s -- --dir $APP_DIR"
		fi
	fi
	echo
	return 0
}

main() {
	parse_args "$@"
	need_root "$@"
	trap cleanup EXIT
	trap 'on_err $LINENO' ERR
	printf '%sАСУ «Путёвка» — сервер на Linux%s\n' "$B" "$N"
	detect_existing
	if [ "$MODE" = uninstall ]; then uninstall; else install_or_update; fi
}

# Вызов в самом конце: если скачивание оборвётся на середине, bash не выполнит недописанный скрипт.
main "$@"
