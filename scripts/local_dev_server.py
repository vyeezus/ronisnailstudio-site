#!/usr/bin/env python3
"""
Local preview: static files from public/ + proxy /api/booking → production (same as Netlify).

Usage (from repo root):
  python3 scripts/local_dev_server.py

Then open:
  http://127.0.0.1:8765/admin-owner-booking.html

Uses live https://ronisnailstudio.com/api/booking so your admin secret hits the deployed
Apps Script unless you change UPSTREAM_BOOKING below to another deployment.
"""
from __future__ import annotations

import json
import os
import sys
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

REPO_ROOT = Path(__file__).resolve().parent.parent
PUBLIC_DIR = REPO_ROOT / "public"
PORT = int(os.environ.get("PORT", "8765"))
UPSTREAM_BOOKING = os.environ.get(
    "LOCAL_BOOKING_UPSTREAM", "https://ronisnailstudio.com/api/booking"
)


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(PUBLIC_DIR), **kwargs)

    def log_message(self, fmt, *args):
        sys.stderr.write("%s - %s\n" % (self.address_string(), fmt % args))

    def _proxy_booking(self, method: str) -> None:
        path = self.path
        if "?" in path:
            qs = path.split("?", 1)[1]
            url = UPSTREAM_BOOKING + "?" + qs
        else:
            url = UPSTREAM_BOOKING

        body = None
        if method == "POST":
            length = int(self.headers.get("Content-Length", "0") or "0")
            body = self.rfile.read(length) if length else b""

        headers = {}
        ct = self.headers.get("Content-Type")
        if ct:
            headers["Content-Type"] = ct
        ua = self.headers.get("User-Agent")
        if ua:
            headers["User-Agent"] = ua

        try:
            req = Request(url, data=body, headers=headers, method=method)
            with urlopen(req, timeout=120) as resp:
                data = resp.read()
                st = resp.status
                out_ct = resp.headers.get("Content-Type", "application/json; charset=utf-8")
        except HTTPError as e:
            data = e.read() if e.fp else b""
            st = e.code
            out_ct = e.headers.get("Content-Type", "application/json; charset=utf-8")
        except URLError as e:
            msg = json.dumps({"status": "error", "message": "proxy_upstream_failed", "detail": str(e.reason)})
            data = msg.encode("utf-8")
            st = HTTPStatus.BAD_GATEWAY
            out_ct = "application/json; charset=utf-8"

        self.send_response(st)
        self.send_header("Content-Type", out_ct)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Local-Preview", "1")
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        if self.path.startswith("/api/booking"):
            self._proxy_booking("GET")
        else:
            super().do_GET()

    def do_POST(self):
        if self.path.startswith("/api/booking"):
            self._proxy_booking("POST")
        else:
            self.send_error(HTTPStatus.NOT_FOUND, "POST only allowed for /api/booking")


def main() -> None:
    if not PUBLIC_DIR.is_dir():
        print("Missing public/ directory:", PUBLIC_DIR, file=sys.stderr)
        sys.exit(1)
    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    print("Serving", PUBLIC_DIR)
    print("Open http://127.0.0.1:%s/admin-owner-booking.html" % PORT)
    print("Proxy /api/booking →", UPSTREAM_BOOKING)
    print("Ctrl+C to stop")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()
