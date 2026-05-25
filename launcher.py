import time
import threading
import socket
import webview

from dotenv import load_dotenv
from run import app

webview.settings['OPEN_EXTERNAL_LINKS_IN_BROWSER'] = True


def is_port_open(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(0.5)
        return sock.connect_ex((host, port)) == 0


#def start_flask():
#    app.run(
#        host="127.0.0.1",
#        port=5000,
#        debug=False,
#        use_reloader=False
#    )

def start_flask():
    app.run(
        host="0.0.0.0",
        port=5000,
        debug=False,
        use_reloader=False
    )

if __name__ == "__main__":
    load_dotenv()

    if not is_port_open("127.0.0.1", 5000):
        flask_thread = threading.Thread(target=start_flask, daemon=True)
        flask_thread.start()

        for _ in range(20):
            if is_port_open("127.0.0.1", 5000):
                break
            time.sleep(0.5)

    webview.create_window(
        "Repair Tracker",
        "http://127.0.0.1:5000/login",
        width=1280,
        height=850,
        min_size=(1100, 700)
    )
    webview.start()