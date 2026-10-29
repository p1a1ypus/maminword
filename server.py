"""МаминWord — лёгкий веб-редактор документов Word.

Запуск:  python server.py [порт]
Документы обрабатываются в памяти и нигде на диске не сохраняются.
"""
import gzip
import json
import mimetypes
import os
import secrets
import socket
import sys
import threading
import time
import traceback
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import quote, unquote, urlparse

import docx2html
import html2docx

HERE = os.path.dirname(os.path.abspath(__file__))
STATIC = os.path.join(HERE, "static")
MAX_BODY = 150 * 1024 * 1024
MIME = {".js": "application/javascript; charset=utf-8", ".css": "text/css; charset=utf-8",
        ".html": "text/html; charset=utf-8", ".svg": "image/svg+xml", ".json": "application/json",
        ".woff": "font/woff", ".woff2": "font/woff2", ".png": "image/png", ".ico": "image/x-icon"}
_gz_cache = {}
_gz_lock = threading.Lock()

# Исходники открытых документов держим в памяти, чтобы при сохранении браузер
# не пересылал файл повторно (через медленный туннель это ощутимо).
DOC_TTL = 24 * 3600
DOC_CACHE_BYTES = 400 * 1024 * 1024
_docs = {}
_docs_lock = threading.Lock()


def doc_put(data):
    token = secrets.token_urlsafe(18)
    now = time.time()
    with _docs_lock:
        for k in [k for k, (t, _) in _docs.items() if now - t > DOC_TTL]:
            del _docs[k]
        while _docs and sum(len(v) for _, v in _docs.values()) + len(data) > DOC_CACHE_BYTES:
            del _docs[min(_docs, key=lambda k: _docs[k][0])]
        _docs[token] = (now, data)
    return token


def doc_get(token):
    with _docs_lock:
        item = _docs.get(token or "")
        if item is None:
            return None
        _docs[token] = (time.time(), item[1])
        return item[1]


class Handler(BaseHTTPRequestHandler):
    server_version = "MaminWord/1.0"
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):
        if "/static/vendor/" in (self.path or ""):
            return
        sys.stderr.write("[%s] %s\n" % (self.log_date_time_string(), fmt % args))

    # ------------------------------------------------------------- ответы
    def send_bytes(self, data, ctype, status=200, extra=None, compress=False):
        if compress and len(data) > 1024 and "gzip" in (self.headers.get("Accept-Encoding") or ""):
            data = gzip.compress(data, 6)
            extra = dict(extra or {}, **{"Content-Encoding": "gzip"})
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("X-Content-Type-Options", "nosniff")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(data)

    def send_json(self, obj, status=200):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_bytes(data, "application/json; charset=utf-8", status, {"Cache-Control": "no-store"},
                        compress=True)

    def send_error_json(self, msg, status=400):
        self.send_json({"error": msg}, status)

    def read_body(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n <= 0:
            raise ValueError("Пустой запрос")
        if n > MAX_BODY:
            raise ValueError("Файл слишком большой (максимум %d МБ)" % (MAX_BODY // 1024 // 1024))
        chunks, left = [], n
        while left > 0:
            chunk = self.rfile.read(min(left, 1 << 20))
            if not chunk:
                break
            chunks.append(chunk)
            left -= len(chunk)
        return b"".join(chunks)

    # ---------------------------------------------------------------- GET
    def do_HEAD(self):
        self.do_GET()

    def do_GET(self):
        path = unquote(urlparse(self.path).path)
        if path in ("/", "/index.html"):
            with open(os.path.join(HERE, "index.html"), "rb") as f:
                return self.send_bytes(f.read(), MIME[".html"], 200, {"Cache-Control": "no-cache"}, compress=True)
        if path == "/favicon.ico":
            return self.static_file("favicon.svg")
        if path == "/api/ping":
            # по этому ответу страница понимает, что рядом сервер (иначе конвертирует сама, в браузере)
            return self.send_json({"ok": True})
        if path.startswith("/static/"):
            return self.static_file(path[len("/static/"):])
        if path == "/api/new":
            try:
                with open(html2docx.TEMPLATE, "rb") as f:
                    return self.send_json(docx2html.convert(f.read()))
            except Exception as e:
                traceback.print_exc()
                return self.send_error_json(str(e), 500)
        self.send_bytes("Не найдено".encode("utf-8"), "text/plain; charset=utf-8", 404)

    def static_file(self, rel):
        full = os.path.realpath(os.path.join(STATIC, rel))
        if not full.startswith(os.path.realpath(STATIC) + os.sep) or not os.path.isfile(full):
            return self.send_bytes(b"Not found", "text/plain", 404)
        ext = os.path.splitext(full)[1].lower()
        ctype = MIME.get(ext) or mimetypes.guess_type(full)[0] or "application/octet-stream"
        vendor = "/vendor/" in full.replace("\\", "/")
        cache = {"Cache-Control": "public, max-age=86400" if vendor else "no-cache"}
        compressible = ext in (".js", ".css", ".html", ".svg", ".json")
        if compressible and "gzip" in (self.headers.get("Accept-Encoding") or ""):
            mtime = os.path.getmtime(full)
            with _gz_lock:
                hit = _gz_cache.get(full)
                if not hit or hit[0] != mtime:
                    with open(full, "rb") as f:
                        hit = (mtime, gzip.compress(f.read(), 6))
                    _gz_cache[full] = hit
            return self.send_bytes(hit[1], ctype, 200, dict(cache, **{"Content-Encoding": "gzip"}))
        with open(full, "rb") as f:
            self.send_bytes(f.read(), ctype, 200, cache)

    # --------------------------------------------------------------- POST
    def do_POST(self):
        path = urlparse(self.path).path
        try:
            body = self.read_body()
        except ValueError as e:
            return self.send_error_json(str(e), 413)
        if path == "/api/open":
            try:
                result = docx2html.convert(body)
                result["token"] = doc_put(body)
                return self.send_json(result)
            except ValueError as e:
                return self.send_error_json(str(e))
            except Exception as e:
                traceback.print_exc()
                return self.send_error_json("Не удалось прочитать документ: %s" % e, 500)
        if path == "/api/save":
            try:
                payload = json.loads(body.decode("utf-8"))
                original = None
                if payload.get("token") and not payload.get("original"):
                    original = doc_get(payload["token"])
                    if original is None:
                        # сервер перезапускали — пусть браузер пришлёт исходник сам
                        return self.send_json({"error": "need_original"}, 409)
                data, warnings = html2docx.convert(payload, original)
                name = (payload.get("filename") or "Документ.docx").replace("\r", "").replace("\n", "")
                if not name.lower().endswith(".docx"):
                    name += ".docx"
                headers = {"Content-Disposition": "attachment; filename=\"document.docx\"; filename*=UTF-8''%s"
                                                  % quote(name),
                           "Cache-Control": "no-store",
                           "X-MaminWord-Warnings": quote(json.dumps(warnings, ensure_ascii=False))}
                return self.send_bytes(data, "application/vnd.openxmlformats-officedocument."
                                             "wordprocessingml.document", 200, headers)
            except Exception as e:
                traceback.print_exc()
                return self.send_error_json("Не удалось сохранить документ: %s" % e, 500)
        self.send_error_json("Неизвестный адрес", 404)


def local_addresses():
    addrs = set()
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            addrs.add(info[4][0])
    except OSError:
        pass
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("10.255.255.255", 1))
        addrs.add(s.getsockname()[0])
        s.close()
    except OSError:
        pass
    return sorted(a for a in addrs if not a.startswith("127."))


def main():
    port = int(sys.argv[1]) if len(sys.argv) > 1 else int(os.environ.get("LITEWORD_PORT", "8000"))
    no_browser = "--no-browser" in sys.argv or os.environ.get("LITEWORD_NO_BROWSER")
    srv = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    srv.daemon_threads = True
    print("=" * 62)
    print(" МаминWord запущен. Не закрывайте это окно, пока работаете.")
    print("=" * 62)
    print(" На этом компьютере:      http://localhost:%d" % port)
    for a in local_addresses():
        kind = "Radmin VPN" if a.startswith("26.") else "локальная сеть"
        print(" Ссылка (%s): http://%s:%d" % (kind, a, port))
    print("-" * 62)
    print(" Остановить: Ctrl+C или закрыть окно.")
    sys.stdout.flush()
    if not no_browser:
        threading.Timer(1.0, lambda: webbrowser.open("http://localhost:%d" % port)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        srv.server_close()


if __name__ == "__main__":
    main()
