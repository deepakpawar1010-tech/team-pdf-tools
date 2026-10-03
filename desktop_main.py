import os
import sys
import time
import socket
import threading
import urllib.request
import webbrowser

# Enable Windows High-DPI awareness for crisp rendering
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-monitor DPI aware
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

from waitress import serve
from app import app


def find_free_port(preferred: int = 5050) -> int:
    """Check if preferred port is free; otherwise pick an available ephemeral port."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            s.bind(("127.0.0.1", preferred))
            return preferred
    except OSError:
        pass

    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def run_server(port: int):
    """Run production multi-threaded Waitress WSGI server."""
    serve(
        app,
        host="127.0.0.1",
        port=port,
        threads=8,
        channel_timeout=300,
        cleanup_interval=30,
        _quiet=True,
    )


def wait_for_server(port: int, timeout: float = 12.0) -> bool:
    """Poll health endpoint until server is listening."""
    start = time.time()
    url = f"http://127.0.0.1:{port}/healthz"
    while time.time() - start < timeout:
        try:
            with urllib.request.urlopen(url, timeout=0.8) as resp:
                if resp.status == 200:
                    return True
        except Exception:
            time.sleep(0.1)
    return False


def get_icon_path() -> str:
    """Resolve icon path in both development and PyInstaller bundled environment."""
    if getattr(sys, "frozen", False):
        base_dir = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(base_dir, "app_icon.ico")
    return icon_path if os.path.exists(icon_path) else None


def main():
    port = find_free_port(5050)
    server_thread = threading.Thread(target=run_server, args=(port,), daemon=True)
    server_thread.start()

    if not wait_for_server(port, timeout=15.0):
        print(f"Error: Could not start local application server on port {port}")
        sys.exit(1)

    url = f"http://127.0.0.1:{port}/"

    # Try launching native WebView2 window first
    launched_native = False
    try:
        import webview
        webview.settings['ALLOW_DOWNLOADS'] = True
        icon_path = get_icon_path()
        window = webview.create_window(
            title="Team PDF Tools",
            url=url,
            width=1320,
            height=880,
            min_size=(960, 640),
            text_select=True,
            confirm_close=False,
        )
        launched_native = True
        webview.start(private_mode=False)
    except Exception as err:
        print(f"Native WebView2 could not start ({err}). Falling back to default browser...")
        launched_native = False

    # Fallback if WebView2 is missing: open default browser and keep server alive
    if not launched_native:
        webbrowser.open(url)
        try:
            while True:
                time.sleep(1)
        except (KeyboardInterrupt, SystemExit):
            pass


if __name__ == "__main__":
    main()
