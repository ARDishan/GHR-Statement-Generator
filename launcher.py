"""
Double-clickable launcher for the Streamlit app.
Runs Streamlit in-process (no subprocess) so a PyInstaller .exe
never relaunches itself recursively, and forces developmentMode
off so packaged builds don't hit Streamlit's dev-mode/port conflict.
"""

import os
import sys
import socket
import threading
import time
import traceback
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


def ensure_config_toml(app_dir, port):
    """Write a .streamlit/config.toml next to the exe that explicitly
    disables developmentMode. A config file value overrides Streamlit's
    auto-detected default, which is what actually causes the
    'server.port does not work when global.developmentMode is true'
    crash inside PyInstaller bundles."""
    config_dir = os.path.join(app_dir, ".streamlit")
    os.makedirs(config_dir, exist_ok=True)
    config_path = os.path.join(config_dir, "config.toml")

    contents = f"""[global]
developmentMode = false

[server]
headless = true
port = {port}

[browser]
gatherUsageStats = false
"""
    with open(config_path, "w") as f:
        f.write(contents)


def main():
    app_dir = get_app_dir()
    app_file = os.path.join(app_dir, "app_streamlit.py")

    if not os.path.exists(app_file):
        raise FileNotFoundError(f"Could not find app_streamlit.py at: {app_file}")

    port = find_free_port()
    url = f"http://localhost:{port}"

    # Belt-and-suspenders: env var too, in case bootstrap reads it first.
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"

    ensure_config_toml(app_dir, port)

    print("Starting Payment Schedule Statement Generator...")
    print(f"App dir: {app_dir}")
    print(f"App file: {app_file}")
    print(f"If your browser doesn't open automatically, go to: {url}")

    open_browser_once(url)

    # Run Streamlit directly via its bootstrap module instead of the
    # click-based CLI. This sidesteps the CLI parsing path that was
    # triggering the developmentMode conflict in the traceback.
    from streamlit.web import bootstrap

    flag_options = {
        "server.port": port,
        "server.headless": True,
        "browser.gatherUsageStats": False,
        "global.developmentMode": False,
    }

    try:
        # Newer Streamlit signature
        bootstrap.load_config_options(flag_options=flag_options)
        bootstrap.run(app_file, "streamlit run", [], flag_options)
    except TypeError:
        # Older Streamlit signature: bootstrap.run(file, is_hello, args, flag_options)
        bootstrap.load_config_options(flag_options=flag_options)
        bootstrap.run(app_file, False, [], flag_options)


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