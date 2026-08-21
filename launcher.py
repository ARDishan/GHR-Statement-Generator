import os
import sys
import time
import socket
import webbrowser
import multiprocessing
import streamlit.web.cli as stcli


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


def open_browser(url):
    time.sleep(2.5)
    webbrowser.open(url)


if __name__ == "__main__":
    multiprocessing.freeze_support()

    app_dir = get_app_dir()
    app_file = os.path.join(app_dir, "app_streamlit.py")
    port = find_free_port()
    url = f"http://localhost:{port}"

    print("Starting Payment Schedule Statement Generator...")
    print(f"If your browser doesn't open automatically, go to: {url}")

    # Launch browser in a background thread
    import threading
    threading.Thread(target=open_browser, args=(url,), daemon=True).start()

    # Pass command line flags directly to Streamlit CLI inside the same process
    sys.argv = [
        "streamlit",
        "run",
        app_file,
        "--global.developmentMode=false",
        "--server.port", str(port),
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
    ]

    sys.exit(stcli.main())