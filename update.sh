#!/usr/bin/env bash
# Обновление сервера «Путёвки» на Linux. С v8.5.6 ставит и обновляет один установщик:
#   curl -fsSL https://raw.githubusercontent.com/MaksimTMB/putevka-rps/main/install.sh | sudo bash
# Этот файл оставлен для старых инструкций: он скачивает и запускает тот же установщик.
set -euo pipefail
URL="https://raw.githubusercontent.com/MaksimTMB/putevka-rps/main/install.sh"
T="$(mktemp)"
trap 'rm -f "$T"' EXIT
curl -fsSL "$URL" -o "$T"
D="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# программа поставлена без службы (вручную) — обновляем её в той же папке
if [ -f "$D/server.py" ] && [ ! -f /etc/systemd/system/putevka.service ]; then set -- --dir "$D" "$@"; fi
if [ "$(id -u)" -eq 0 ]; then bash "$T" "$@"; else sudo bash "$T" "$@"; fi
