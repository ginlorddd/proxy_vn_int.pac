from __future__ import annotations

import argparse
import json
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Any

from config import load_config, save_config
from logger import setup_logger
from scheduler import MailScheduler
from sender import send_mail
from state import load_state

logger = setup_logger()


def run_api(scheduler: MailScheduler) -> ThreadingHTTPServer:
    cfg = load_config()
    host = cfg["api"].get("host", "127.0.0.1")
    port = int(cfg["api"].get("port", 8765))
    token = cfg["api"].get("token", "change-me")

    class Handler(BaseHTTPRequestHandler):
        def _json(self, code: int, data: dict[str, Any]) -> None:
            body = json.dumps(data, ensure_ascii=False).encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def _authorized(self) -> bool:
            return self.headers.get("X-AutoMail-Token") == token

        def do_GET(self) -> None:  # noqa: N802
            if not self._authorized():
                self._json(401, {"ok": False, "message": "Unauthorized"})
                return
            if self.path == "/health":
                self._json(200, {"ok": True})
            elif self.path == "/config":
                self._json(200, {"ok": True, "config": load_config()})
            elif self.path == "/state":
                self._json(200, {"ok": True, "state": load_state()})
            else:
                self._json(404, {"ok": False, "message": "Not found"})

        def do_POST(self) -> None:  # noqa: N802
            if not self._authorized():
                self._json(401, {"ok": False, "message": "Unauthorized"})
                return
            length = int(self.headers.get("Content-Length", "0") or 0)
            payload = json.loads(self.rfile.read(length) or b"{}")
            config = load_config()
            if self.path == "/send":
                mail = {**config.get("mail", {}), **payload.get("mail", payload)}
                self._json(200, send_mail(mail))
            elif self.path == "/config":
                config.update(payload)
                save_config(config)
                self._json(200, {"ok": True, "message": "Saved"})
            elif self.path == "/scheduler/start":
                scheduler.start()
                self._json(200, {"ok": True, "message": "Scheduler started"})
            elif self.path == "/scheduler/stop":
                scheduler.stop()
                self._json(200, {"ok": True, "message": "Scheduler stopped"})
            else:
                self._json(404, {"ok": False, "message": "Not found"})

        def log_message(self, fmt: str, *args: Any) -> None:
            logger.info("API " + fmt, *args)

    server = ThreadingHTTPServer((host, port), Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    logger.info("API started at http://%s:%s", host, port)
    return server


def main() -> None:
    parser = argparse.ArgumentParser(description="AutoMail - Outlook scheduler + HTTP API")
    parser.add_argument("--no-ui", action="store_true", help="Chạy API/scheduler không mở giao diện")
    parser.add_argument("--send-now", action="store_true", help="Gửi mail ngay theo config")
    args = parser.parse_args()

    scheduler = MailScheduler(load_config)
    api_server = run_api(scheduler)
    if load_config().get("schedule", {}).get("enabled"):
        scheduler.start()
    if args.send_now or load_config().get("send_on_startup"):
        send_mail(load_config().get("mail", {}))
    if args.no_ui:
        threading.Event().wait()
    else:
        import tkinter as tk
        from ui import AutoMailUI
        root = tk.Tk()
        AutoMailUI(root, scheduler)
        try:
            root.mainloop()
        finally:
            scheduler.stop()
            api_server.shutdown()


if __name__ == "__main__":
    main()
