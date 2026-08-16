import os
import sys
import socket
import threading
import time
import traceback
import webbrowser
from multiprocessing import freeze_support

# Must be set before streamlit is imported anywhere, or its config
# system will auto-detect "development mode" inside a PyInstaller
# bundle and refuse to accept --server.port.
os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"


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

    if not os.path.exists(app_file):
        raise FileNotFoundError(f"Could not find app_streamlit.py at: {app_file}")

    port = find_free_port()
    url = f"http://localhost:{port}"

    print("Starting Payment Schedule Statement Generator...")
    print(f"App dir: {app_dir}")
    print(f"App file: {app_file}")
    print(f"If your browser doesn't open automatically, go to: {url}")

    open_browser_once(url)

    sys.argv = [
        "streamlit", "run", app_file,
        "--server.port", str(port),
        "--server.headless", "true",
        "--browser.gatherUsageStats", "false",
        "--global.developmentMode", "false",
    ]

    try:
        from streamlit.web import cli as stcli
    except ImportError:
        from streamlit import cli as stcli

    sys.exit(stcli.main())


if __name__ == "__main__":
    freeze_support()
    try:
        main()
    except Exception:
        log_path = os.path.join(get_app_dir(), "launcher_error.log")
        with open(log_path, "w") as f:
            f.write(traceback.format_exc())
        traceback.print_exc()
        input("An error occurred. Press Enter to close...")