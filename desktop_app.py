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
    port = _read_saved_port()
    if is_server_responding(port):
        open_app_window(f"{PROTO}127.0.0.1:{port}")
        return

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

    open_app_window(app_url)
    threading.Thread(target=_cleanup_orphan_bundles, daemon=True).start()

    while True:
        time.sleep(2)


if __name__ == "__main__":
    main()
