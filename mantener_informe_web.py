import os
import sys
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
WEB_ROOT = BASE_DIR / "publicacion_web"
HOST = "127.0.0.1"
PORT = 8082


def main():
    os.chdir(WEB_ROOT)
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")
    server = ThreadingHTTPServer((HOST, PORT), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == "__main__":
    main()