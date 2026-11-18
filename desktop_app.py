import os, sys, time, socket, threading, subprocess, urllib.request, json, traceback

_appdata = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
_log_dir = os.path.join(_appdata, "PutevkaApp")
try:
    os.makedirs(_log_dir, exist_ok=True)
    _log_file = open(os.path.join(_log_dir, "putevka_desktop.log"), "a", encoding="utf-8", buffering=1)
except Exception:
    _log_file = open(os.devnull, "w", encoding="utf-8")

if sys.stdout is None:
    sys.stdout = _log_file
if sys.stderr is None:
    sys.stderr = _log_file


# ---------------------------------------------------------------------------------------
# v8.1.2: «программа обновилась, а окно старое». Старые версии после замены exe запускают
# новый exe из СВОЕГО процесса — новый наследует служебные переменные PyInstaller и считает
# себя «дочерним»: не распаковывает себя, а работает из распакованной папки СТАРОЙ версии
# (код новый, страница и данные комплекта — старые; F5 не помогает). Воспроизведено на
# Windows 17.11.2026: 7.3.5 → 8.1.1 = сервер v8.1.1, страница v7.3.5.
# Признак: в нормальном запуске exe «из одного файла» наш родитель — загрузчик этого же exe.
# Если родитель другой (батник обновления) — перезапускаемся с чистым окружением.
# Проверка — ДО загрузки тяжёлых библиотек (из чужой папки они могут не совпасть по версии).
# ---------------------------------------------------------------------------------------
def _parent_is_own_bootloader():
    try:
        import ctypes
        from ctypes import wintypes
        k32 = ctypes.windll.kernel32
        k32.OpenProcess.restype = wintypes.HANDLE
        h = k32.OpenProcess(0x1000, False, os.getppid())  # PROCESS_QUERY_LIMITED_INFORMATION
        if not h:
            return False  # родителя уже нет (батник завершился) — значит, это не наш загрузчик
        try:
            buf = ctypes.create_unicode_buffer(32768)
            size = wintypes.DWORD(32768)
            if not k32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
                return False
            return os.path.samefile(buf.value, sys.executable)
        finally:
            k32.CloseHandle(h)
    except Exception:
        return True  # проверить не удалось — не мешаем обычному запуску


def _relaunch_clean_if_foreign_bundle():
    if not getattr(sys, "frozen", False) or os.name != "nt":
        return
    if os.environ.pop("PUTEVKA_RELAUNCHED", None) == "1":
        return  # уже перезапущены — второй раз не перезапускаемся ни при каком исходе
    if not any(k.startswith("_PYI_") or k.startswith("_MEIPASS") for k in os.environ):
        return
    if _parent_is_own_bootloader():
        return
    env = {k: v for k, v in os.environ.items() if not (k.startswith("_PYI_") or k.startswith("_MEIPASS"))}
    env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    env["PUTEVKA_RELAUNCHED"] = "1"
    try:
        print("[%s] запуск из чужой распакованной папки %s — перезапуск с чистым окружением"
              % (time.strftime("%Y-%m-%d %H:%M:%S"), getattr(sys, "_MEIPASS", "?")))
        subprocess.Popen([sys.executable] + sys.argv[1:], env=env, close_fds=True,
                         cwd=os.path.dirname(sys.executable),
                         creationflags=0x00000008 | 0x00000200)  # DETACHED | NEW_PROCESS_GROUP
        os._exit(0)
    except Exception:
        pass  # не вышло — работаем как есть (как до v8.1.2)


_relaunch_clean_if_foreign_bundle()


# ---------------------------------------------------------------------------------------
# v8.2.1: ОДНО МЕСТО УСТАНОВКИ, ЗНАЧОК НА РАБОЧЕМ СТОЛЕ И ОБНОВЛЕНИЕ БЕЗ СЮРПРИЗОВ.
# Владелец 17.11: «версия для виндовс не обновляется!». Причины (все три — из кода до 8.2.0):
#  1) старый процесс завершали по имени «Putevka.exe»: файл «Putevka (1).exe» или
#     «Putevka_Windows_v7.5.0.exe» не закрывался, замена не проходила, запускалась прежняя версия;
#  2) после закрытия окна программа остаётся работать в фоне; новый exe, увидев работающую,
#     просто открывал окно К НЕЙ — на экране опять старая версия;
#  3) данные лежат рядом с exe: новый exe из «Загрузок» стартовал с пустой базой.
# Теперь: программа живёт в %LOCALAPPDATA%\Programs\Putevka (права администратора не нужны),
# запуск из любого места переносит её туда вместе с базой, ставит ярлык «Путёвка» с фирменным
# значком на рабочий стол и в «Пуск», старые ярлыки ведут туда же; старая версия закрывается
# по номеру процесса; новый exe сам заменяет старый (без батника — русские пути не ломаются).
# Портативный режим (флешка): файл portable.txt рядом с exe или PUTEVKA_PORTABLE=1.
# ---------------------------------------------------------------------------------------
INSTALL_DIR = os.path.join(_appdata, "Programs", "Putevka")
INSTALL_EXE = os.path.join(INSTALL_DIR, "Putevka.exe")
INSTALL_MARK = os.path.join(INSTALL_DIR, "installed_exe_version.txt")
_NOWIN = 0x08000000  # CREATE_NO_WINDOW


def _log(msg):
    line = "[%s] %s" % (time.strftime("%Y-%m-%d %H:%M:%S"), msg)
    try:
        _log_file.write(line + "\n")  # журнал — всегда в файл (у exe есть консоль, print уходит туда)
        _log_file.flush()
    except Exception:
        pass


def _own_version():
    for base in (getattr(sys, "_MEIPASS", None), os.path.dirname(os.path.abspath(__file__))):
        try:
            with open(os.path.join(base, "version.txt"), encoding="utf-8") as f:
                return "v" + f.read().strip().lstrip("v")
        except Exception:
            continue
    return "v0.0.0"


def _vt(v):
    out = []
    for x in str(v or "").strip().lstrip("v").split(".")[:3]:
        try:
            out.append(int(x))
        except Exception:
            out.append(0)
    while len(out) < 3:
        out.append(0)
    return tuple(out)


def _same_path(a, b):
    try:
        return bool(a and b) and os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))
    except Exception:
        return False


def _clean_env(extra=None):
    env = {k: v for k, v in os.environ.items()
           if not (k.startswith("_PYI_") or k.startswith("_MEIPASS") or k == "PUTEVKA_RELAUNCHED")}
    env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    if extra:
        env.update(extra)
    return env


def _ps(script, timeout=40, env=None):
    """PowerShell через -EncodedCommand (UTF-16): русские пути и имена файлов не ломаются."""
    import base64
    enc = base64.b64encode(("[Console]::OutputEncoding=[Text.Encoding]::UTF8\n" + script).encode("utf-16-le")).decode("ascii")
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                            "-EncodedCommand", enc], capture_output=True, timeout=timeout,
                           creationflags=_NOWIN, env=env or os.environ.copy())
        return r.stdout.decode("utf-8", "replace").strip()
    except Exception as e:
        _log("PowerShell: %s" % e)
        return ""


def _read_port_early():
    try:
        with open(os.path.join(_log_dir, "putevka_port.txt"), encoding="utf-8") as f:
            p = int(f.read().strip())
            return p if 1024 <= p <= 65535 else 18080
    except Exception:
        return 18080


def _running_server_info(port):
    """Кто сейчас работает на порту программы: версия, процесс и его загрузчик (exe «из одного
    файла» — это два процесса: загрузчик и сама программа)."""
    try:
        with urllib.request.urlopen("http://127.0.0.1:%d/api/version" % port, timeout=1.5) as r:
            ver = json.loads(r.read().decode("utf-8")).get("version", "")
    except Exception:
        return None
    info = {"version": ver, "port": port}
    out = _ps(r'''
$o = @{}
$c = Get-NetTCPConnection -LocalPort __PORT__ -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
if ($c) {
  $p = Get-CimInstance Win32_Process -Filter "ProcessId=$($c.OwningProcess)"
  $o.pid = [int]$c.OwningProcess; $o.path = $p.ExecutablePath; $o.ppid = [int]$p.ParentProcessId
  $pp = Get-CimInstance Win32_Process -Filter "ProcessId=$($p.ParentProcessId)"
  if ($pp) { $o.ppath = $pp.ExecutablePath }
}
$o | ConvertTo-Json -Compress
'''.replace("__PORT__", str(int(port))))
    try:
        info.update(json.loads(out or "{}"))
    except Exception:
        pass
    return info


def _kill_running(info, wait_port=True):
    """Закрыть работающую программу по НОМЕРУ процесса (не по имени файла)."""
    pids = []
    if info.get("pid"):
        pids.append(int(info["pid"]))
    if info.get("ppid") and _same_path(info.get("ppath"), info.get("path")):
        pids.append(int(info["ppid"]))  # загрузчик того же exe
    for pid in pids:
        if pid and pid != os.getpid():
            subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True, creationflags=_NOWIN)
    _log("остановлена работающая версия %s (%s), процессы %s" % (info.get("version"), info.get("path"), pids))
    if wait_port and info.get("port"):
        for _ in range(40):
            try:
                with urllib.request.urlopen("http://127.0.0.1:%d/api/version" % info["port"], timeout=0.5):
                    pass
                time.sleep(0.25)
            except Exception:
                break
        time.sleep(0.5)


def _db_has_content(db_path):
    try:
        if os.path.getsize(db_path) < 50:
            return False
        with open(db_path, encoding="utf-8") as f:
            d = json.load(f)
        return bool(isinstance(d, dict) and (d.get("users") or d.get("profiles")))
    except Exception:
        return False


def _db_summary(data_dir):
    db = os.path.join(data_dir, "putevka_db.json")
    try:
        with open(db, encoding="utf-8") as f:
            d = json.load(f)
        return {"profiles": len(d.get("profiles") or []), "users": len(d.get("users") or []),
                "docs": len(d.get("savedDocuments") or []),
                "when": time.strftime("%d.%m.%Y %H:%M", time.localtime(os.path.getmtime(db)))}
    except Exception:
        return {"profiles": 0, "users": 0, "docs": 0, "when": "?"}


def _pick_data_source(running):
    """v8.3.0: базу берём ТОЛЬКО у работающей прежней версии или рядом с запущенным exe (обновление на месте).
    Раньше ещё и искали по «Рабочему столу», «Загрузкам», «Документам» — так в «свежую» установку владельца
    молча попали его тестовые ведомости (17.11). Поиск по папкам убран."""
    cands = []
    if running and running.get("path"):
        cands.append(os.path.join(os.path.dirname(running["path"]), "data"))
    cands.append(os.path.join(os.path.dirname(sys.executable), "data"))
    for d in cands:
        if not _same_path(d, os.path.join(INSTALL_DIR, "data")) and _db_has_content(os.path.join(d, "putevka_db.json")):
            return d
    return None


def _ask_migrate(src):
    """Спросить, переносить ли найденную базу. Без окна (проверки) — по PUTEVKA_MIGRATE=yes/no."""
    forced = (os.environ.get("PUTEVKA_MIGRATE") or "").strip().lower()
    if forced in ("yes", "no"):
        _log("перенос базы — по PUTEVKA_MIGRATE=%s" % forced)
        return forced == "yes"
    i = _db_summary(src)
    text = ("Найдена рабочая база прежней версии программы:\n%s\n\n"
            "Водителей: %d, пользователей: %d, документов: %d\nПоследнее изменение: %s\n\n"
            "Перенести эту базу в программу?\n\n"
            "«Да» — перенести (обычное обновление, всё останется как было).\n"
            "«Нет» — начать с чистой базы. Найденная база останется на месте; её можно загрузить позже "
            "(«Документы» → «Резервная копия базы»).") % (src, i["profiles"], i["users"], i["docs"], i["when"])
    try:
        import ctypes
        # MB_YESNO | MB_ICONQUESTION | MB_SETFOREGROUND | MB_TOPMOST
        ans = ctypes.windll.user32.MessageBoxW(None, text, "Путёвка — перенос рабочей базы", 0x4 | 0x20 | 0x10000 | 0x40000)
        _log("перенос базы: ответ %s" % ("да" if ans == 6 else "нет"))
        return ans == 6
    except Exception as e:
        _log("окно вопроса не открылось (%s) — базу НЕ переношу, она остаётся на месте" % e)
        return False


def _copy_data(src, dst):
    import shutil
    tmp = dst + ".copying"
    shutil.rmtree(tmp, ignore_errors=True)
    shutil.copytree(src, tmp)
    if os.path.exists(dst):
        os.replace(dst, dst + ".before-" + time.strftime("%Y%m%d-%H%M%S"))
    os.replace(tmp, dst)
    try:
        with open(os.path.join(src, "ПЕРЕНЕСЕНО — данные теперь в Programs-Putevka.txt"), "w", encoding="utf-8") as f:
            f.write("Рабочая база программы «Путёвка» скопирована %s в папку\n%s\n"
                    "Программа теперь работает оттуда (ярлык «Путёвка» на рабочем столе). "
                    "Эта папка — резервная копия на момент переноса, её можно удалить.\n" % (time.strftime("%d.%m.%Y %H:%M"), dst))
    except Exception:
        pass


def _make_shortcuts(old_targets):
    """Ярлык «Путёвка» с фирменным значком на рабочий стол и в «Пуск»; старые ярлыки на прежние
    exe программы — перенаправить на установленную."""
    env = os.environ.copy()
    env["PUTEVKA_SC_EXE"] = INSTALL_EXE
    env["PUTEVKA_SC_OLD"] = "|".join(t for t in old_targets if t)
    # WScript.Shell не сохраняет ярлык, если в ИМЕНИ файла буквы не из «ANSI»-кодировки системы
    # (английская Windows и «Путёвка.lnk» — проверено 17.11): создаём под латинским именем,
    # путь и значок пишем через Shell.Application (Юникод), потом переименовываем.
    out = _ps(r'''
$exe = $env:PUTEVKA_SC_EXE; $dir = Split-Path $exe
$old = @($env:PUTEVKA_SC_OLD -split '\|' | Where-Object { $_ })
$ws = New-Object -ComObject WScript.Shell; $sh = New-Object -ComObject Shell.Application
$desk = [Environment]::GetFolderPath('Desktop'); $menu = [Environment]::GetFolderPath('Programs')
function Set-PutevkaLink($folder, $name) {
  $final = Join-Path $folder $name
  $tmpName = 'putevka-' + [guid]::NewGuid().ToString('N') + '.lnk'; $tmp = Join-Path $folder $tmpName
  $s = $ws.CreateShortcut($tmp); $s.TargetPath = (Join-Path $env:SystemRoot 'explorer.exe'); $s.Save()
  $l = $sh.Namespace($folder).ParseName($tmpName).GetLink
  $l.Path = $exe; $l.WorkingDirectory = $dir; $l.Description = 'Путёвка — путевые листы'; $l.SetIconLocation($exe, 0); $l.Save()
  if (Test-Path -LiteralPath $final) { Remove-Item -LiteralPath $final -Force }
  Move-Item -LiteralPath $tmp -Destination $final -Force
  "ярлык: $final"
}
foreach ($f in @($desk, $menu)) {
  try { Set-PutevkaLink $f 'Путёвка.lnk' } catch {
    "ярлык не создан в ${f}: $($_.Exception.Message)"
    Get-ChildItem -LiteralPath $f -Filter 'putevka-*.lnk' -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue }
}
foreach ($lnk in @(Get-ChildItem -LiteralPath $desk -Filter *.lnk -ErrorAction SilentlyContinue)) {
  try { $l = $sh.Namespace($desk).ParseName($lnk.Name).GetLink
        if ($l.Path -and ($old -contains $l.Path) -and ($l.Path -ne $exe)) {
          $l.Path = $exe; $l.WorkingDirectory = $dir; $l.SetIconLocation($exe, 0); $l.Save(); "старый ярлык перенаправлен: $($lnk.Name)" } } catch {}
}
''', env=env)
    for line in out.splitlines():
        _log(line)


def _copy_exe_to(dst):
    import shutil
    tmp = dst + ".tmp"
    for _ in range(40):
        try:
            shutil.copyfile(sys.executable, tmp)
            os.replace(tmp, dst)
            return os.path.getsize(dst) == os.path.getsize(sys.executable)
        except Exception:
            time.sleep(1.0)  # файл ещё занят (процесс закрывается, проверяет антивирус)
    return False


def _launch(exe, args=(), extra_env=None):
    subprocess.Popen([exe] + list(args), env=_clean_env(extra_env), close_fds=True, cwd=os.path.dirname(exe),
                     creationflags=0x00000008 | 0x00000200)  # DETACHED | NEW_PROCESS_GROUP


def _finish_update_if_asked():
    """Новый exe, скачанный обновлением, ставит себя на место старого: «--finish-update <куда> <pid…>».
    Этот договор не меняется между версиями — его вызывает сервер ПРЕЖНЕЙ версии."""
    if "--finish-update" not in sys.argv:
        return
    i = sys.argv.index("--finish-update")
    target = sys.argv[i + 1] if len(sys.argv) > i + 1 else ""
    pids = [a for a in sys.argv[i + 2:] if a.isdigit()]
    log_path = os.path.join(os.path.dirname(target) or _log_dir, "update.log")

    def ulog(m):
        _log("обновление: " + m)
        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write("[%s] %s\n" % (time.strftime("%Y-%m-%d %H:%M:%S"), m))
        except Exception:
            pass
    try:
        ulog("начало: %s → %s (%s)" % (sys.executable, target, _own_version()))
        time.sleep(3.0)  # ответ сервера и сообщение успевают дойти до окна
        for pid in pids:
            if int(pid) != os.getpid():
                subprocess.run(["taskkill", "/F", "/PID", pid], capture_output=True, creationflags=_NOWIN)
        ulog("прежняя версия остановлена (процессы %s)" % ",".join(pids))
        ok = _copy_exe_to(target)
        ulog("новый exe установлен" if ok else "ВНИМАНИЕ: файл программы занят — новый exe НЕ установлен")
        if ok and _same_path(target, INSTALL_EXE):
            try:
                with open(INSTALL_MARK, "w", encoding="utf-8") as f:
                    f.write(_own_version())
            except Exception:
                pass
        _launch(target, ["--after-update"])
        ulog("запуск %s" % target)
    except Exception:
        ulog("ошибка: " + traceback.format_exc())
    os._exit(0)


def _install_and_handoff():
    """Запуск НЕ из места установки: перенести программу и базу туда, поставить ярлыки, запустить."""
    if not getattr(sys, "frozen", False) or os.name != "nt":
        return
    if _same_path(sys.executable, INSTALL_EXE):
        return
    if os.environ.get("PUTEVKA_PORTABLE") == "1" or os.path.exists(os.path.join(os.path.dirname(sys.executable), "portable.txt")):
        _log("портативный режим: работаю из %s" % os.path.dirname(sys.executable))
        return
    mine = _own_version()
    try:
        os.makedirs(INSTALL_DIR, exist_ok=True)
        running = _running_server_info(_read_port_early())
        installed_ver = "v0.0.0"
        if os.path.exists(INSTALL_EXE):
            try:
                with open(INSTALL_MARK, encoding="utf-8") as f:
                    installed_ver = f.read().strip() or "v0.0.1"
            except Exception:
                installed_ver = "v0.0.1"
        _log("запуск из %s (%s); установлена %s; работает %s" % (sys.executable, mine,
             installed_ver if os.path.exists(INSTALL_EXE) else "—", (running or {}).get("version", "—")))
        replace_installed = (not os.path.exists(INSTALL_EXE)) or _vt(installed_ver) < _vt(mine)
        if running:
            if _same_path(running.get("path"), INSTALL_EXE) and not replace_installed:
                pass  # работает установленная той же или новее версии — просто откроем её окно
            else:
                _kill_running(running)
        if replace_installed:
            _ps(r'''Get-Process | Where-Object { $_.Path -and ($_.Path -eq $env:PUTEVKA_SC_EXE) } | Stop-Process -Force -ErrorAction SilentlyContinue''',
                env=dict(os.environ, PUTEVKA_SC_EXE=INSTALL_EXE), timeout=20)
        inst_db = os.path.join(INSTALL_DIR, "data", "putevka_db.json")
        if not _db_has_content(inst_db):
            src = _pick_data_source(running)
            if src and _ask_migrate(src):
                _copy_data(src, os.path.join(INSTALL_DIR, "data"))
                _log("рабочая база перенесена: %s → %s" % (src, os.path.join(INSTALL_DIR, "data")))
            elif src:
                _log("рабочая база НЕ перенесена (ответ «нет»): %s — программа начнёт с чистой базы" % src)
        if replace_installed:
            if not _copy_exe_to(INSTALL_EXE):
                _log("не удалось записать %s — работаю из текущего места" % INSTALL_EXE)
                return
            with open(INSTALL_MARK, "w", encoding="utf-8") as f:
                f.write(mine)
            _log("программа установлена: %s (%s)" % (INSTALL_EXE, mine))
        old_targets = [sys.executable]
        if running and running.get("path") and not _same_path(running["path"], INSTALL_EXE):
            old_targets.append(running["path"])
            if not _same_path(running["path"], sys.executable):
                try:  # прежний exe тоже становится новым: его старый ярлык или двойной щелчок приведут сюда же
                    import shutil
                    shutil.copyfile(sys.executable, running["path"])
                except Exception:
                    pass
        _make_shortcuts(old_targets)
        _launch(INSTALL_EXE, [a for a in sys.argv[1:] if a == "--after-update"])
        _log("передал запуск установленной программе")
        os._exit(0)
    except SystemExit:
        raise
    except Exception:
        _log("установка не удалась, работаю из текущего места:\n" + traceback.format_exc())


def _cleanup_update_leftovers():
    for attempt in range(6):
        time.sleep(10)
        left = [sys.executable + x for x in (".new", ".update.bat", ".tmp") if os.path.exists(sys.executable + x)]
        for f in left:
            try:
                os.remove(f)
                _log("убран временный файл обновления %s" % f)
            except Exception as e:
                if attempt == 5:
                    _log("не удалось убрать %s: %s" % (f, e))
        if not left:
            return


if getattr(sys, "frozen", False) and os.name == "nt":
    _finish_update_if_asked()
    _install_and_handoff()


def _cleanup_orphan_bundles():
    """Убрать распакованные папки прошлых запусков программы, брошенные при обновлении
    (старый процесс завершается принудительно и не успевает убрать за собой ~20 МБ).
    Удаляем только папку с признаками нашей программы, старше 10 минут и НИКЕМ не занятую:
    занятую папку (в ней загружена библиотека Python работающего процесса) Windows не даст
    переименовать — такую пропускаем."""
    try:
        import tempfile, shutil
        own = os.path.normcase(os.path.abspath(getattr(sys, "_MEIPASS", "") or ""))
        tmp = tempfile.gettempdir()
        for name in os.listdir(tmp):
            if not name.startswith("_MEI"):
                continue
            path = os.path.join(tmp, name)
            if os.path.normcase(os.path.abspath(path)) == own or not os.path.isdir(path):
                continue
            if not (os.path.exists(os.path.join(path, "static", "index.html"))
                    and os.path.exists(os.path.join(path, "data", "app_data.json"))
                    and os.path.exists(os.path.join(path, "version.txt"))):
                continue
            if time.time() - os.path.getmtime(path) < 600:
                continue
            trash = path + "_putevka_old"
            try:
                os.rename(path, trash)
            except Exception:
                continue  # занята работающей программой
            shutil.rmtree(trash, ignore_errors=True)
    except Exception:
        pass


# v8.2.0: программа на этом компьютере слушает только 127.0.0.1 — принимаем запросы только с этим адресом
# (чужой сайт в браузере не сможет обратиться к программе через подменённое имя — «DNS rebinding»)
os.environ.setdefault("PUTEVKA_LOCAL_ONLY", "1")

import uvicorn
import server

DEFAULT_PORT = 18080
PROTO = "http" + "://"
PORT_FILE = os.path.join(_log_dir, "putevka_port.txt")


def _read_saved_port():
    try:
        with open(PORT_FILE, encoding="utf-8") as f:
            p = int(str(f.read()).strip())
            if 1024 <= p <= 65535:
                return p
    except Exception:
        pass
    return DEFAULT_PORT


def _save_port(p):
    try:
        with open(PORT_FILE, "w", encoding="utf-8") as f:
            f.write(str(p))
    except Exception:
        pass


def _port_is_free(port):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", port))
        return True
    except Exception:
        return False
    finally:
        s.close()


def _write_error(text):
    paths = [_log_dir]
    try:
        paths.append(os.path.dirname(sys.executable))
    except Exception:
        pass
    for d in paths:
        try:
            with open(os.path.join(d, "startup_error.txt"), "w", encoding="utf-8") as f:
                f.write(text)
        except Exception:
            pass


def _find_free_port():
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]
    finally:
        s.close()


def is_server_responding(port=DEFAULT_PORT):
    try:
        with urllib.request.urlopen(f"{PROTO}127.0.0.1:{port}/api/version", timeout=1.0) as r:
            return r.status == 200
    except Exception:
        return False


def wait_for_server(url, timeout=15.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(url + "/api/version", timeout=1.0) as r:
                if r.status == 200:
                    return True
        except Exception:
            time.sleep(0.15)
    return False


def run_backend(port):
    try:
        config = uvicorn.Config(server.app, host="127.0.0.1", port=port, log_level="warning", log_config=None)
        srv = uvicorn.Server(config)
        srv.run()
    except Exception:
        tb = traceback.format_exc()
        _write_error(tb)
        print("Backend error:\n" + tb)


def open_app_window(app_url):
    if os.environ.get("PUTEVKA_HEADLESS_TEST") == "1":
        return
    candidates = [
        os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"),
        os.path.expandvars(r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"),
        os.path.expandvars(r"%LocalAppData%\Microsoft\Edge\Application\msedge.exe"),
        os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
    ]
    profile_dir = os.path.join(_log_dir, "window_profile")
    os.makedirs(profile_dir, exist_ok=True)
    for exe in candidates:
        if os.path.exists(exe):
            try:
                subprocess.Popen([
                    exe,
                    f"--app={app_url}",
                    f"--user-data-dir={profile_dir}",
                    "--window-size=1460,920",
                    "--no-first-run",
                    "--no-default-browser-check",
                    "--disable-extensions",
                ])
                return
            except Exception:
                continue
    import webbrowser
    webbrowser.open(app_url)


def _open_error_page(err_text):
    err_path = os.path.join(_log_dir, "startup_error.html")
    try:
        body = (
            "<html><head><meta charset='utf-8'><title>Ошибка запуска ПУТЁВКА</title></head>"
            "<body style='font-family:Arial,sans-serif;padding:40px;max-width:900px'>"
            "<h2 style='color:#b91c1c'>Не удалось запустить программу</h2>"
            "<p>Скопируйте текст ниже и отправьте разработчику, либо пришлите файл "
            "<b>startup_error.txt</b> из папки:</p>"
            "<pre style='background:#f1f5f9;padding:10px;border-radius:8px'>%s</pre>"
            "<pre style='background:#fff1f2;padding:14px;border:1px solid #fecaca;border-radius:8px;white-space:pre-wrap'>%s</pre>"
            "</body></html>"
        ) % (_log_dir, err_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
        with open(err_path, "w", encoding="utf-8") as f:
            f.write(body)
    except Exception:
        pass
    open_app_window("file:///" + err_path.replace(os.sep, "/"))


def main():
    # Стабильный порт: новое окно после обновления попадает на тот же адрес,
    # а повторный запуск не плодит второй сервер (просто открывает окно к работающему).
    started_at = time.time()
    port = _read_saved_port()
    after_update = "--after-update" in sys.argv
    if is_server_responding(port):
        # v8.2.1: раньше здесь открывалось окно к ЛЮБОЙ работающей программе — даже к старой версии,
        # оставшейся в фоне после закрытия окна («обновил — а версия прежняя»)
        info = _running_server_info(port) if os.name == "nt" else None
        if not info or _vt(info.get("version")) >= _vt(_own_version()):
            open_app_window(f"{PROTO}127.0.0.1:{port}")
            return
        _kill_running(info)

    try:
        subprocess.run([
            "powershell", "-NoProfile", "-Command",
            "Get-NetTCPConnection -LocalPort 18080 -ErrorAction SilentlyContinue | ForEach-Object { if ($_.OwningProcess -gt 0) { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue } }"
        ], creationflags=0x08000000, check=False)
    except Exception:
        pass

    if not _port_is_free(port):
        port = _find_free_port()
    _save_port(port)
    app_url = f"{PROTO}127.0.0.1:{port}"

    t = threading.Thread(target=run_backend, args=(port,), daemon=True)
    t.start()

    ok = wait_for_server(app_url, timeout=15.0)
    if not ok:
        err = "Сервер не поднялся на %s\n\nСм. файлы putevka_desktop.log и startup_error.txt в папке:\n%s" % (app_url, _log_dir)
        try:
            p = os.path.join(_log_dir, "startup_error.txt")
            if os.path.exists(p):
                err += "\n\n" + open(p, encoding="utf-8").read()
        except Exception:
            pass
        _write_error(err)
        _open_error_page(err)
        return

    # v8.2.1: после обновления старое окно само перезагрузится на новую версию — второе не открываем
    if after_update:
        while time.time() - started_at < 25 and getattr(server, "LAST_CLIENT_SEEN", 0) < started_at:
            time.sleep(0.5)
        if getattr(server, "LAST_CLIENT_SEEN", 0) >= started_at:
            _log("после обновления: окно программы уже открыто — второе не открываю")
        else:
            open_app_window(app_url)
    else:
        open_app_window(app_url)
    threading.Thread(target=_cleanup_orphan_bundles, daemon=True).start()
    if getattr(sys, "frozen", False):
        threading.Thread(target=_cleanup_update_leftovers, daemon=True).start()
        try:
            if _same_path(sys.executable, INSTALL_EXE):
                with open(INSTALL_MARK, "w", encoding="utf-8") as f:
                    f.write(_own_version())
        except Exception:
            pass

    while True:
        time.sleep(2)


if __name__ == "__main__":
    main()
