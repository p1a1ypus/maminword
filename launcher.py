"""Запуск МаминWord одним окном.

    python launcher.py             — редактор + ссылки для этого компьютера и локальной сети
    python launcher.py --internet  — то же + публичная ссылка через туннель localhost.run

Вся логика здесь, а не в .bat: cmd.exe неправильно читает .bat-файлы с русским текстом.
"""
import os
import re
import shutil
import socket
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser

APP = "МаминWord"
APP_MARK = APP.encode("utf-8")  # по этой строке в странице узнаём «свой» сервер
HERE = os.path.dirname(os.path.abspath(__file__))
LINK_FILE = os.path.join(HERE, "link.txt")
ANSI = re.compile(r"\x1b\[[0-9;?]*[A-Za-z]")
TUNNEL_URL = re.compile(r"tunneled with tls termination,\s*(https://[a-z0-9-]+\.lhr\.life)")


def say(text=""):
    print(text, flush=True)


def ensure_deps():
    try:
        import docx  # noqa: F401
        import lxml  # noqa: F401
    except ImportError:
        say("Устанавливаю недостающие компоненты (один раз)...")
        subprocess.call([sys.executable, "-m", "pip", "install", "-r", os.path.join(HERE, "requirements.txt")])


def is_liteword(port):
    try:
        with urllib.request.urlopen("http://127.0.0.1:%d/" % port, timeout=3) as r:
            return APP_MARK in r.read(4096)
    except Exception:
        return False


def port_free(port):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("0.0.0.0", port))
        return True
    except OSError:
        return False
    finally:
        s.close()


def start_server(port):
    """Запускает сервер в фоне. Если МаминWord уже работает на порту — использует его."""
    sys.path.insert(0, HERE)
    import server
    from http.server import ThreadingHTTPServer

    while not port_free(port):
        if is_liteword(port):
            say("Редактор уже запущен в другом окне (порт %d) — использую его." % port)
            return port, server
        port += 1
    srv = ThreadingHTTPServer(("0.0.0.0", port), server.Handler)
    srv.daemon_threads = True
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return port, server


def check_url(url, tries=6):
    for _ in range(tries):
        try:
            with urllib.request.urlopen(url, timeout=20) as r:
                if r.status == 200 and APP_MARK in r.read(4096):
                    return True
        except Exception:
            pass
        time.sleep(3)
    return False


def announce(url):
    with open(LINK_FILE, "w", encoding="utf-8") as f:
        f.write(url + "\n")
    try:
        subprocess.run(["clip"], input=url.encode("ascii"), check=False)
        copied = " (уже скопирована в буфер обмена)"
    except Exception:
        copied = ""
    say()
    say("=" * 64)
    say("  ССЫЛКА ДЛЯ ДРУГОГО ЧЕЛОВЕКА%s:" % copied)
    say()
    say("      " + url)
    say()
    say("  Она же записана в файл link.txt рядом с программой.")
    say("=" * 64)
    ok = check_url(url)
    say("  Проверка: ссылка открывается — всё работает." if ok else
        "  Проверка не прошла — подождите минуту или перезапустите окно.")
    say("  Не закрывайте это окно, пока человек работает с документом.")
    say()


def run_tunnel(port):
    ssh = shutil.which("ssh")
    if not ssh:
        say("Не найден ssh (OpenSSH). В Windows 10/11: Параметры > Приложения > "
            "Дополнительные компоненты > Клиент OpenSSH.")
        return
    attempt = 0
    while True:
        attempt += 1
        if attempt > 1:
            say("Связь с туннелем прервалась — переподключаюсь (ссылка будет новой)...")
            time.sleep(5)
        else:
            say("Получаю публичную ссылку через localhost.run...")
        cmd = [ssh, "-T", "-o", "StrictHostKeyChecking=accept-new", "-o", "ServerAliveInterval=30",
               "-o", "ServerAliveCountMax=3", "-o", "ExitOnForwardFailure=yes",
               "-R", "80:127.0.0.1:%d" % port, "nokey@localhost.run"]
        try:
            # ssh в той же консоли: закрыли окно — туннель закрылся вместе с ним
            proc = subprocess.Popen(cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        except OSError as e:
            say("Не удалось запустить ssh: %s" % e)
            return
        got = False
        for raw in iter(proc.stdout.readline, b""):
            line = ANSI.sub("", raw.decode("utf-8", "replace")).strip()
            m = TUNNEL_URL.search(line)
            if m and not got:
                got = True
                threading.Thread(target=announce, args=(m.group(1),), daemon=True).start()
            elif "Permission denied" in line or "Could not resolve" in line or "Connection refused" in line:
                say("ssh: " + line)
        proc.wait()
        if not got and attempt >= 3:
            say("Не получается подключиться к localhost.run. Проверьте интернет или используйте")
            say("ссылку для локальной сети / Radmin VPN выше.")
            time.sleep(30)


def set_title(text):
    try:
        import ctypes
        ctypes.windll.kernel32.SetConsoleTitleW(text)
    except Exception:
        pass


def main():
    internet = "--internet" in sys.argv
    set_title(APP + (" — ссылка через интернет" if internet else ""))
    ensure_deps()
    port, server = start_server(8000)
    say("=" * 64)
    say(" %s запущен. Это окно не закрывайте, пока работаете." % APP)
    say("=" * 64)
    say(" На этом компьютере:        http://localhost:%d" % port)
    for a in server.local_addresses():
        kind = "Radmin VPN" if a.startswith("26.") else "локальная сеть"
        say(" Ссылка (%s): http://%s:%d" % (kind, a, port))
    say("-" * 64)
    say(" Остановить: закрыть окно или Ctrl+C.")
    if not internet:
        webbrowser.open("http://localhost:%d" % port)
    try:
        if internet:
            run_tunnel(port)
        else:
            while True:
                time.sleep(3600)
    except KeyboardInterrupt:
        say("Остановлено.")


if __name__ == "__main__":
    main()
