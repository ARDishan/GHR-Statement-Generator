"""
Double-clickable launcher for the Streamlit app.
Runs Streamlit in-process (no subprocess) so a PyInstaller .exe
never relaunches itself recursively.
"""

import os
import sys
import socket
import threading
import time
import webbrowser
from multiprocessing import freeze_support


def get_app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def find_free_port(start=8501, tries=20):
    for port in range(start, start + tries):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return start


def open_browser_once(url, delay=2.0):
    def _open():
        time.sleep(delay)
        webbrowser.open(url)
    threading.Thread(target=_open, daemon=True).start()


def main():
    app_dir = get_app_dir()
    app_file = os.path.join(app_dir, "app_streamlit.py")
    port = find_free_port()
    url = f"http://localhost:{port}"

    print("Starting Payment Schedule Statement Generator...")
    print(f"If your browser doesn't open automatically, go to: {url}")

    open_browser_once(url)

    sys.argv = [
        "streamlit", "run", app_file,
        "--server.port", str(port),
        "--server.headless", "true",
        "--browser.gatherUsageStats", "false",
    ]

    try:
        from streamlit.web import cli as stcli   # Streamlit >= 1.12
    except ImportError:
        from streamlit import cli as stcli       # older Streamlit

    sys.exit(stcli.main())


if __name__ == "__main__":
    freeze_support()   # required for frozen .exe on Windows
    main()