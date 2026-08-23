import os
import sys
import time
import socket
import webbrowser
import threading
import multiprocessing

import streamlit.web.cli as stcli


def get_base_dir():
    """
    Get the directory containing bundled files.

    When running normally:
        project directory

    When running as PyInstaller executable:
        temporary PyInstaller extraction directory
    """
    if getattr(sys, "frozen", False):
        return sys._MEIPASS

    return os.path.dirname(
        os.path.abspath(__file__)
    )


def get_executable_dir():
    """
    Get the directory where StatementGenerator.exe is located.
    Useful for external assets such as assets/.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(
            os.path.abspath(sys.executable)
        )

    return os.path.dirname(
        os.path.abspath(__file__)
    )


def find_free_port(start=8501, tries=20):
    for port in range(start, start + tries):
        with socket.socket(
            socket.AF_INET,
            socket.SOCK_STREAM
        ) as s:
            if s.connect_ex(
                ("127.0.0.1", port)
            ) != 0:
                return port

    return start


def open_browser(url):
    time.sleep(2.5)
    webbrowser.open(url)


def main():

    multiprocessing.freeze_support()

    base_dir = get_base_dir()

    app_file = os.path.join(
        base_dir,
        "app_streamlit.py"
    )

    if not os.path.exists(app_file):
        raise FileNotFoundError(
            f"Streamlit application not found:\n{app_file}"
        )

    port = find_free_port()

    url = f"http://127.0.0.1:{port}"

    threading.Thread(
        target=open_browser,
        args=(url,),
        daemon=True
    ).start()

    sys.argv = [
        "streamlit",
        "run",
        app_file,

        "--global.developmentMode=false",

        "--server.port",
        str(port),

        "--server.address",
        "127.0.0.1",

        "--server.headless=true",

        "--browser.gatherUsageStats=false",

        "--server.fileWatcherType",
        "none",
    ]

    sys.exit(
        stcli.main()
    )


if __name__ == "__main__":
    main()