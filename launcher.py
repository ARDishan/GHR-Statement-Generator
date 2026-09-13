import os
import sys
import time
import socket
import webbrowser
import multiprocessing
import threading


def get_app_dir():
    """
    Returns the directory where bundled files are located.
    """

    if getattr(sys, "frozen", False):
        # PyInstaller one-file executable
        return sys._MEIPASS

    # Running normally from Python
    return os.path.dirname(os.path.abspath(__file__))


def find_free_port(start=8501, tries=20):

    for port in range(start, start + tries):

        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:

            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port

    return start


def open_browser(url):

    # Give Streamlit a moment to start
    time.sleep(3)

    webbrowser.open(url)


def main():

    app_dir = get_app_dir()

    app_file = os.path.join(
        app_dir,
        "app_streamlit.py"
    )

    # --------------------------------------------------------
    # Check Streamlit application
    # --------------------------------------------------------

    if not os.path.isfile(app_file):

        raise FileNotFoundError(
            "Streamlit application not found:\n"
            + app_file
        )

    # --------------------------------------------------------
    # Find free port
    # --------------------------------------------------------

    port = find_free_port()

    url = f"http://localhost:{port}"

    # --------------------------------------------------------
    # Start browser
    # --------------------------------------------------------

    threading.Thread(
        target=open_browser,
        args=(url,),
        daemon=True
    ).start()

    # --------------------------------------------------------
    # Start Streamlit
    # --------------------------------------------------------

    import streamlit.web.cli as stcli

    sys.argv = [
        "streamlit",
        "run",
        app_file,

        "--global.developmentMode=false",

        "--server.port",
        str(port),

        "--server.headless=true",

        "--browser.gatherUsageStats=false",

        "--server.fileWatcherType",
        "none",
    ]

    sys.exit(
        stcli.main()
    )

if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()