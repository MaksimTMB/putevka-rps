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

import uvicorn
import server

DEFAULT_PORT = 18080
PROTO = "http" + "://"


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
    port = _find_free_port()
    app_url = f"{PROTO}127.0.0.1:{port}"

    try:
        subprocess.run([
            "powershell", "-NoProfile", "-Command",
            "Get-NetTCPConnection -LocalPort 18080 -ErrorAction SilentlyContinue | ForEach-Object { if ($_.OwningProcess -gt 0) { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue } }"
        ], creationflags=0x08000000, check=False)
    except Exception:
        pass

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

    while True:
        time.sleep(2)


if __name__ == "__main__":
    main()
