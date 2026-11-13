import os, sys, time, socket, threading, subprocess, urllib.request, json

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

def is_server_responding(port=DEFAULT_PORT):
    try:
        with urllib.request.urlopen(f"{PROTO}127.0.3.1:{port}/api/version", timeout=1.0) as r:
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
        config = uvicorn.Config(server.app, host="127.0.3.1", port=port, log_level="warning", log_config=None)
        srv = uvicorn.Server(config)
        srv.run()
    except Exception as e:
        print("Backend error:", e)

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

def main():
    app_url = f"{PROTO}127.0.3.1:{DEFAULT_PORT}"
    if is_server_responding(DEFAULT_PORT):
        open_app_window(app_url)
        return

    t = threading.Thread(target=run_backend, args=(DEFAULT_PORT,), daemon=True)
    t.start()
    wait_for_server(app_url, timeout=15.0)
    open_app_window(app_url)

    while True:
        time.sleep(2)

if __name__ == "__main__":
    main()
