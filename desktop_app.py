import os
import sys
import time
import socket
import threading
import subprocess
import urllib.request
import uvicorn
import server

APP_TITLE = "АСУ «ПУТЁВКА» v6.6.0 — Форма № 3 (ОКУД 0345001)"


def find_free_port(preferred=18080):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        if s.connect_ex(("127.0.0.1", preferred)) != 0:
            return preferred
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def wait_for_server(url, timeout=12.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(url + "/api/version", timeout=1.5) as resp:
                if resp.status == 200:
                    return True
        except Exception:
            time.sleep(0.2)
    return False


def run_backend(port):
    config = uvicorn.Config(server.app, host="127.0.0.1", port=port, log_level="warning")
    srv = uvicorn.Server(config)
    srv.run()


def launch_native_edge_or_chrome_app_window(app_url):
    """Fallback native desktop app window using Edge/Chrome --app mode (no browser tabs or address bar) if pywebview is unavailable."""
    candidates = [
        os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"),
        os.path.expandvars(r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"),
        os.path.expandvars(r"%LocalAppData%\Microsoft\Edge\Application\msedge.exe"),
        os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
    ]
    profile_dir = os.path.join(server.BASE_DIR, "data", ".desktop_window_profile")
    os.makedirs(profile_dir, exist_ok=True)
    for exe in candidates:
        if os.path.exists(exe):
            proc = subprocess.Popen([
                exe,
                f"--app={app_url}",
                f"--user-data-dir={profile_dir}",
                "--window-size=1460,920",
                "--disable-extensions",
                "--no-first-run",
                "--no-default-browser-check",
            ])
            proc.wait()
            return True
    return False


def main():
    port = find_free_port(18080)
    app_url = f"http://127.0.0.1:{port}"

    t = threading.Thread(target=run_backend, args=(port,), daemon=True)
    t.start()
    wait_for_server(app_url, timeout=12.0)

    # 1. Основной режим: нативное десктопное окно Windows через pywebview (WinForms / Edge WebView2)
    try:
        import webview
        window = webview.create_window(
            title=APP_TITLE,
            url=app_url,
            width=1460,
            height=920,
            min_size=(1100, 700),
            resizable=True,
            text_select=True,
        )
        webview.start()
        os._exit(0)
    except Exception:
        pass

    # 2. Резервный нативный режим окна приложения (Edge/Chrome App Window без адресной строки и вкладок)
    if launch_native_edge_or_chrome_app_window(app_url):
        os._exit(0)

    # 3. Если ни pywebview, ни Edge не найдены — удерживаем фоновый процесс
    while True:
        time.sleep(1)


if __name__ == "__main__":
    main()
