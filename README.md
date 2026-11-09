# АСУ «ПУТЁВКА» (v5.0.0 — v6.5.0)
**Автоматизированная система управления, расчета и печати путевых листов легкового автомобиля (Типовая межотраслевая форма № 3 по ОКУД 0345001)**

Репозиторий содержит полную историю разработки и релизов серверного комплекса **«ПУТЁВКА»** начиная с базовой версии **v5.0.0** и заканчивая актуальным релизом **v6.5.0**, включая исходный код сервера (`FastAPI` + `openpyxl`), клиентское одностраничное приложение, шаблоны строгой печатной формы ОКУД 0345001, предзагруженную гео-базу объектов и АЗС «Роснефть», скрипты автоматического развертывания и обновления для **Linux**, **Windows**, **macOS** и **Docker**, а также архивы всех релизов в каталоге `releases/`.

## 1. Архитектура и ключевые возможности

| Компонент / Модуль | Назначение и функциональность |
| :--- | :--- |
| **Первичная регистрация и 2 ролевых интерфейса** | При первом входе система открывает экран создания Главного Администратора (`#firstSetupScreen`). Интерфейс разделен на две среды: **Администратор** (полный доступ ко всем настройкам, профилям, чекам, картам и пользователям) и **Водитель** (доступ строго на чтение, просмотр и печать путевых листов своего автомобиля). |
| **Верхняя панель и навигация** | Слева направо: название системы, номер версии и код формы (`Форма № 3 / ОКУД 0345001`), по центру — текущий пользователь и его роль, правее — управление пользователями и ролями, справа — кнопка входа/выхода. Ниже расположен блок основных плиток: **Настройка шапки**, **Маршруты и ПЛ**, **Чеки ГСМ**, **Карты**. |
| **Шаблоны шапки ПЛ и профили водителей** | Создание, сохранение, редактирование и удаление шаблонов оформленной шапки путевого листа с мгновенным применением к любому существующему или новому профилю водителя. |
| **Точный предпросмотр А4 и двусторонняя печать** | Предпросмотр лицевой стороны (шапки) и оборотной стороны (маршрутов) в точных пропорциях листа А4 (`210 × 297 мм`, 1:1 с печатью). При пакетной выгрузке каждый путевой лист формируется строго парой страниц: **Лист 1 — Лицевая сторона (шапка)**, **Лист 2 (оборот) — Маршруты дня**. |
| **Полностраничный экспорт в Word (`.docx`) и Excel (`.xlsx`)** | Генерация документов в стандартном формате OpenXML (`application/vnd.openxmlformats-officedocument.wordprocessingml.document` и `spreadsheetml.sheet`), корректно открывающихся в Microsoft Word, Microsoft Excel и МойОфис/LibreOffice с масштабированием на весь лист А4 без обрезания границ. |
| **Физика бака и вариативный расчет пробегов (`v6.2`–`v6.5`)** | Автоматический пересчет межзаправочных окон и последующих дней месяца по чекам АЗС без отрицательных балансов и переливов бака (`minReserve <= остаток <= tankCap`). Исключены шаблонные повторы суммарного пробега между заправками и дневного километража. |
| **Реалистичный расчет времени движения и работы (`v6.5`)** | Время в пути рассчитывается индивидуально для каждого плеча от километража, типа трассы/района (город, пригород/СНТ, межгород Котовск/Тригуляй, трасса Р22) и трафика по часам пик. Время работы на объектах между переездами распределяется неравномерно без шаблонных 15-минутных шагов с точным возвращением на базу к концу смены (`18-00`). |
| **Импорт и навигация по чекам ГСМ (`v6.1`–`v6.4`)** | Двухуровневая навигация по реально существующим **Годам → Месяцам**, редактирование начального остатка топлива на 1-е число месяца, а также двойной парсер ведомостей транзакций «Роснефть» (`.xlsx`, `.xls`, `.csv`) с автосозданием периодов, учетом возвратов и переносом выходных чеков на рабочие дни. |
| **Календарь дней, рандом маршрутов и защита удалений** | Любое удаление в системе требует подтверждения в модальном окне. Отключенные дни сохраняются в календаре месяца с кнопкой **«+ Вернуть»**, доступно добавление любой календарной даты месяца (включая доп. смены) и генератор случайного числа маршрутов (`2–12` плеч) с полем ручного ввода. |

## 2. Структура репозитория

```text
putevka-rps/
├── server.py                        # Сервер FastAPI: API состояния, авторизации, парсинга чеков и генерации .docx/.xlsx
├── requirements.txt                 # Зависимости Python (fastapi, uvicorn, openpyxl, pydantic)
├── version.txt                      # Текущая версия сборки (v6.5.0)
├── static/
│   └── index.html                   # Клиентское приложение (Админ + Водитель, А4 предпросмотр, графики, калькулятор)
├── data/
│   ├── app_data.json                # Эталонная база данных релиза (схема ОКУД 0345001, локации, шаблоны, профили)
│   └── putevka_db.json              # Рабочая база данных сервера (автоматически синхронизируется при работе)
├── templates_xls/
│   └── template_0345001.xlsx        # Шаблон формы № 3 (ОКУД 0345001)
├── releases/                        # Архивы всех версий проекта с v5.0.0 по v6.5.0
│   ├── putevka_server_v5_0.zip
│   ├── putevka_server_v6_0.zip
│   ├── putevka_server_v6_1.zip
│   ├── putevka_server_v6_2.zip
│   ├── putevka_server_v6_3.zip
│   ├── putevka_server_v6_4.zip
│   └── putevka_server_v6_5.zip
├── scripts/
│   ├── install_linux.sh             # Чистая установка + создание службы systemd для Linux
│   ├── update_on_server.sh          # Скрипт быстрого обновления на Linux-сервере
│   ├── install_windows.bat          # Чистая установка для Windows (CMD)
│   ├── update_windows.bat           # Обновление для Windows (CMD)
│   ├── install_windows.ps1          # Чистая установка для Windows (PowerShell)
│   ├── update_windows.ps1           # Обновление для Windows (PowerShell)
│   ├── install_macos.sh             # Чистая установка для macOS
│   ├── update_macos.sh              # Обновление для macOS
│   └── putevka.service              # Шаблон unit-файла systemd для автозапуска в Linux
├── Dockerfile                       # Сборка контейнера Docker
├── docker-compose.yml               # Запуск в 1 команду через Docker Compose с сохранением тома data/
├── CHANGELOG.md                     # Подробная история изменений от v5.0.0 до v6.5.0
└── README.md                        # Документация по установке, обновлению и эксплуатации
```

## 3. История версий в Git-репозитории (v5.0.0 — v6.5.0)

Все этапы развития проекта зафиксированы в истории коммитов ветки `main` и отмечены соответствующими Git-тегами (а также сохранены в виде готовых архивов в папке `releases/`):

| Версия (Git tag) | Архив в `releases/` | Основные изменения релиза |
| :--- | :--- | :--- |
| **`v5.0.0`** | `putevka_server_v5_0.zip` | Базовая серверная версия: калькулятор маршрутов, ведомость ГСМ, гео-база Тамбов/Котовск/Пригород (13 АЗС «Роснефть»), базовая генерация листов. |
| **`v6.0.0`** | `putevka_server_v6_0.zip` | Полная смена концепции: разделение на 2 интерфейса (Администратор и Водитель), первичная регистрация логина и пароля Администратора при первом входе, перенос кнопок «Сохранить» строго вниз каждого окна, отключение стрелок-спиннеров в числовых полях, навигация по чекам (Года → Месяцы), точный предпросмотр листа А4 (лицевая + оборотная сторона) и полностраничный экспорт в Word (`.docx`) и Excel (`.xlsx`). |
| **`v6.1.0`** | `putevka_server_v6_1.zip` | Добавлены кнопки «Удалить» в пару к «Сохранить» во всех окнах, удалены пустые периоды в чеках, открыто прямое редактирование начального остатка топлива на 1-е число месяца, внедрена жесткая физика бака (`minReserve .. tankCap`) без отрицательных остатков, добавлен менеджер шаблонов шапки ПЛ. |
| **`v6.2.0`** | `putevka_server_v6_2.zip` | Универсальное модальное подтверждение для любого удаления, обратимое исключение и возврат/добавление любых календарных дат месяца, свободное добавление точек маршрута с автопересчетом последующих дней по чекам АЗС, генератор случайного числа маршрутов с полем ручного ввода (`2–12`), устранение одинаковых пробегов между заправками. |
| **`v6.3.0`** | `putevka_server_v6_3.zip` | Исправлен приоритет серверного флага `adminSetupCompleted: false` над старыми ключами `localStorage`, гарантирующий обязательный показ формы первичной настройки Администратора при первом входе после развертывания. |
| **`v6.4.0`** | `putevka_server_v6_4.zip` | Реализован двойной парсер ведомостей чеков ГСМ (`POST /api/fuel/parse-upload` на `openpyxl`/`csv` + клиентский `SheetJS`) с автоматическим созданием периодов, учетом возвратов, переносом заправок выходного дня и мгновенным отображением загруженных чеков в таблице. |
| **`v6.5.0`** | `putevka_server_v6_5.zip` | Внедрен реалистичный расчет времени движения от километража, типа дороги/района и дорожного трафика по времени суток (без шаблонных 15 минут и без одинакового времени на разные маршруты), а также неравномерное распределение времени работы на объектах между поездками. |

Просмотр истории коммитов и переключение между версиями в репозитории:

```bash
# Просмотр всей истории релизов и тегов
git log --oneline --decorate --graph

# Список всех версий (тегов)
git tag -l

# Сравнение изменений между версиями (например, v5.0.0 и v6.5.0)
git diff v5.0.0 v6.5.0 --stat
```

## 4. Публикация репозитория в GitHub / GitLab / Gitea

Архив репозитория уже содержит инициализированную папку `.git` со всеми коммитами (`v5.0.0` → `v6.5.0`) и тегами. Чтобы выгрузить его в ваш облачный или корпоративный Git-сервер:

```bash
# 1. Распакуйте архив репозитория и перейдите в папку
unzip putevka_git_repository_v5_0_to_v6_5.zip -d putevka-rps
cd putevka-rps

# 2. Привяжите ваш удаленный репозиторий (GitHub / GitLab / Gitea)
git remote add origin https://github.com/MaksimTMB/putevka-rps.git

# 3. Отправьте ветку main и все теги версий (v5.0.0 .. v6.5.0)
git branch -M main
git push -u origin main --tags
```

## 5. Установка и обновление для различных ОС

Ниже приведены полные инструкции по чистой установке, настройке автозапуска и обновлению системы для **Linux**, **Windows**, **macOS** и **Docker**. По умолчанию сервер запускается на порту **`8080`** (`http://<IP-адрес-сервера>:8080`).

### 5.1. Linux (Ubuntu, Debian, Astra Linux, ALT Linux, CentOS / RHEL / Rocky)

#### Чистая установка на Linux (автоматическим скриптом)

```bash
# 1. Установка системных пакетов Python 3 и venv
# Для Ubuntu / Debian / Astra Linux:
sudo apt-get update && sudo apt-get install -y python3 python3-venv python3-pip unzip psmisc

# Для CentOS / RHEL / Rocky / AlmaLinux:
# sudo dnf install -y python3 python3-pip unzip psmisc

# Для ALT Linux:
# sudo apt-get update && sudo apt-get install -y python3 python3-module-venv python3-module-pip unzip psmisc

# 2. Клонирование репозитория (или распаковка архива) в /opt/asu_waybills
sudo mkdir -p /opt/asu_waybills
sudo unzip -o putevka_git_repository_v5_0_to_v6_5.zip -d /opt/asu_waybills
cd /opt/asu_waybills

# 3. Запуск автоматического установщика Linux (создает venv, ставит зависимости и службу systemd)
sudo bash scripts/install_linux.sh
```

#### Ручная установка и настройка службы `systemd` на Linux

```bash
cd /opt/asu_waybills
python3 -m venv venv
./venv/bin/pip install --upgrade pip
./venv/bin/pip install -r requirements.txt

# Инициализация рабочей базы данных
cp -n data/app_data.json data/putevka_db.json
chmod -R 775 data/

# Установка и запуск фоновой службы systemd
sudo cp scripts/putevka.service /etc/systemd/system/putevka.service
sudo systemctl daemon-reload
sudo systemctl enable putevka
sudo systemctl restart putevka

# Проверка статуса службы
sudo systemctl status putevka
```

#### Обновление существующего сервера на Linux

Если сервер уже установлен (например, в `/opt/asu_waybills`), обновление выполняется одной командой:

```bash
# Вариант А: Обновление через скрипт из репозитория или архива релиза
cd /opt/asu_waybills
sudo bash update_on_server.sh

# Вариант Б: Обновление через Git (если проект склонирован из Git-репозитория)
cd /opt/asu_waybills
sudo git pull origin main --tags
sudo ./venv/bin/pip install -r requirements.txt
sudo cp -f data/app_data.json data/putevka_db.json   # Выполнять только если требуется сброс к чистой базе релиза и первичной настройке Админа
sudo systemctl restart putevka || sudo bash update_on_server.sh
```

> **Примечание при обновлении:** Если вы хотите **сохранить текущих пользователей, созданные профили и загруженные чеки** при обновлении через Git, не перезаписывайте `data/putevka_db.json` файлом `data/app_data.json` — сервер автоматически обновит структуру и тайминги существующей базы при старте. Если же требуется чистый первый вход с заданием нового логина и пароля Администратора, используйте `update_on_server.sh`.

### 5.2. Windows (Windows 10 / 11 / Windows Server 2016 / 2019 / 2022)

Предварительное требование: установленный **Python 3.10+** с официального сайта [python.org](https://www.python.org/downloads/windows/) (при установке обязательно отметьте галочку **«Add python.exe to PATH»**).

#### Чистая установка в Windows через PowerShell (`install_windows.ps1`)

Откройте **PowerShell** от имени Администратора и выполните:

```powershell
# 1. Разрешите выполнение локальных скриптов для текущей сессии
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# 2. Распакуйте архив в C:\asu_waybills
Expand-Archive -Path ".\putevka_git_repository_v5_0_to_v6_5.zip" -DestinationPath "C:\asu_waybills" -Force
cd C:\asu_waybills

# 3. Запустите скрипт установки (создаст venv, установит пакеты, откроет порт 8080 в брандмауэре и запустит сервер)
.\scripts\install_windows.ps1
```

#### Чистая установка в Windows через Командную строку CMD (`install_windows.bat`)

```cmd
cd /d C:\asu_waybills
call scripts\install_windows.bat
```

Или пошагово вручную в CMD / PowerShell:

```cmd
cd /d C:\asu_waybills
python -m venv venv
call venv\Scripts\activate.bat
pip install --upgrade pip
pip install -r requirements.txt
if not exist data\putevka_db.json copy /Y data\app_data.json data\putevka_db.json
python server.py
```

#### Регистрация в качестве фоновой службы Windows (через NSSM или Планировщик заданий)

```powershell
# Вариант 1: Автозапуск при старте Windows через встроенный Планировщик заданий (Task Scheduler)
$Action = New-ScheduledTaskAction -Execute "C:\asu_waybills\venv\Scripts\python.exe" -Argument "C:\asu_waybills\server.py" -WorkingDirectory "C:\asu_waybills"
$Trigger = New-ScheduledTaskTrigger -AtStartup
$Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask -TaskName "PutevkaServer" -Action $Action -Trigger $Trigger -Principal $Principal -Force
Start-ScheduledTask -TaskName "PutevkaServer"

# Вариант 2: Через утилиту NSSM (если установлена)
nssm install PutevkaServer "C:\asu_waybills\venv\Scripts\python.exe" "C:\asu_waybills\server.py"
nssm set PutevkaServer AppDirectory "C:\asu_waybills"
nssm start PutevkaServer
```

#### Обновление в Windows (`update_windows.ps1` или `update_windows.bat`)

```powershell
# В PowerShell:
cd C:\asu_waybills
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\scripts\update_windows.ps1

# Или в CMD:
cd /d C:\asu_waybills
call scripts\update_windows.bat
```

### 5.3. macOS (macOS Ventura / Sonoma / Sequoia — Intel & Apple Silicon M1/M2/M3/M4)

#### Чистая установка на macOS

```bash
# 1. Установка Python 3 через Homebrew (если не установлен)
brew install python

# 2. Распаковка проекта в ~/asu_waybills
mkdir -p ~/asu_waybills
unzip -o putevka_git_repository_v5_0_to_v6_5.zip -d ~/asu_waybills
cd ~/asu_waybills

# 3. Запуск скрипта установки для macOS
bash scripts/install_macos.sh
```

#### Обновление на macOS

```bash
cd ~/asu_waybills
bash scripts/update_macos.sh
```

### 5.4. Docker и Docker Compose (Универсальный запуск для Linux, Windows и macOS)

В корне репозитория подготовлены `Dockerfile` и `docker-compose.yml`. База данных `data/` монтируется как постоянный том (`volume`), поэтому все данные сохраняются при перезапуске контейнера.

```bash
# Первичная сборка и запуск контейнера в фоновом режиме
docker compose up -d --build

# Просмотр логов работы сервера
docker compose logs -f

# Обновление после получения новой версии файлов
docker compose down
docker compose up -d --build
```

## 6. Проверка работоспособности после установки или обновления

1. Откройте в браузере адрес **`http://<IP-вашего-сервера>:8080`** (или `http://localhost:8080` при локальном запуске) и нажмите **`Ctrl+F5`** (`Cmd+Shift+R` на macOS) для сброса кэша браузера.
2. При первом входе заполните форму **«Первичная настройка Администратора»** (ФИО, логин и пароль Главного Администратора) и нажмите **«Сохранить Администратора и войти в систему»**.
3. Для создания учетных записей водителей перейдите в плитку **«Пользователи и роли»** в правой верхней части шапки, назначьте водителю логин, пароль и привязанный профиль автомобиля, после чего нажмите **«Сохранить»** в нижней части окна.
