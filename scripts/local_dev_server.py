#!/usr/bin/env python3
"""
Local preview: static files from public/ + proxy /api/booking → production (same as Netlify).

Usage (from repo root):
  python3 scripts/local_dev_server.py

Then open:
  http://127.0.0.1:8765/admin-owner-booking.html

By default, POSTs for the admin native calendar (ownerCalendarWeek, week or month) are answered with
**mock data** on this machine only, so you can preview the UI without deploying Apps Script.
Other POSTs still go to production.

  LOCAL_MOCK_ADMIN_CALENDAR=0   # send ownerCalendarWeek to live API instead
  LOCAL_BOOKING_UPSTREAM=...    # alternate booking URL
"""
import calendar
import json
import os
import re
import sys
from typing import Optional
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.error import HTTPError, URLError
from datetime import datetime, timedelta
from urllib.request import Request, urlopen

REPO_ROOT = Path(__file__).resolve().parent.parent
PUBLIC_DIR = REPO_ROOT / "public"
PORT = int(os.environ.get("PORT", "8765"))
UPSTREAM_BOOKING = os.environ.get(
    "LOCAL_BOOKING_UPSTREAM", "https://ronisnailstudio.com/api/booking"
)


def _json_bool_true(v) -> bool:
    return v is True or v == "true" or v == 1 or v == "1"


def _local_dt_to_ms(dt: datetime) -> int:
    return int(dt.timestamp() * 1000)


def mock_owner_calendar_week_if_applicable(body: bytes) -> Optional[bytes]:
    """If body is admin calendar POST, return fake JSON so UI works without deployed GAS."""
    try:
        d = json.loads(body.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError, TypeError, ValueError):
        return None
    if not _json_bool_true(d.get("ownerCalendarWeek")):
        return None
    view = str(d.get("calView") or "week").strip().lower()
    now = datetime.now()

    if view == "month":
        mk = str(d.get("month") or "").strip()
        if re.match(r"^\d{4}-\d{2}$", mk):
            y, mo = int(mk[:4]), int(mk[5:7])
        else:
            y, mo = now.year, now.month
        if mo < 1 or mo > 12:
            y, mo = now.year, now.month
        first = datetime(y, mo, 1, 0, 0, 0, 0)
        # JS Date.getDay(): Sunday=0. Python weekday(): Monday=0 … Sunday=6
        leading_blank_days = (first.weekday() + 1) % 7
        last_dom = calendar.monthrange(y, mo)[1]
        days = []
        for dom in range(1, last_dom + 1):
            cur = datetime(y, mo, dom, 0, 0, 0, 0)
            nxt = cur + timedelta(days=1)
            ymd = cur.strftime("%Y-%m-%d")
            label = cur.strftime("%a, %b ") + str(dom)
            days.append(
                {
                    "ymd": ymd,
                    "startMs": _local_dt_to_ms(cur),
                    "endMs": _local_dt_to_ms(nxt),
                    "label": label,
                    "dayOfMonth": dom,
                }
            )
        month_key = f"{y:04d}-{mo:02d}"
        d_studio = min(15, last_dom)
        d_personal = min(22, last_dom)
        t0 = datetime(y, mo, d_studio, 10, 0, 0, 0)
        t1 = datetime(y, mo, d_studio, 11, 30, 0, 0)
        t2 = datetime(y, mo, d_personal, 14, 0, 0, 0)
        t3 = datetime(y, mo, d_personal, 15, 30, 0, 0)
        events = [
            {
                "start": _local_dt_to_ms(t0),
                "end": _local_dt_to_ms(t1),
                "title": "Preview · studio (mock data)",
                "allDay": False,
                "calendar": "studio",
            },
            {
                "start": _local_dt_to_ms(t2),
                "end": _local_dt_to_ms(t3),
                "title": "Preview · personal (mock data)",
                "allDay": False,
                "calendar": "personal",
            },
        ]
        payload = {
            "status": "success",
            "timeZone": "America/Chicago",
            "calView": "month",
            "month": month_key,
            "leadingBlankDays": leading_blank_days,
            "days": days,
            "events": events,
        }
        return json.dumps(payload).encode("utf-8")

    days_back = (now.weekday() + 1) % 7
    sun = (now - timedelta(days=days_back)).replace(hour=0, minute=0, second=0, microsecond=0)
    days = []
    cur = sun
    for _ in range(7):
        ymd = cur.strftime("%Y-%m-%d")
        nxt = cur + timedelta(days=1)
        label = cur.strftime("%a, %b ") + str(cur.day)
        days.append(
            {
                "ymd": ymd,
                "startMs": _local_dt_to_ms(cur),
                "endMs": _local_dt_to_ms(nxt),
                "label": label,
            }
        )
        cur = nxt
    mon_10 = sun + timedelta(days=1, hours=10)
    mon_1130 = sun + timedelta(days=1, hours=11, minutes=30)
    thu_14 = sun + timedelta(days=3, hours=14)
    thu_1530 = sun + timedelta(days=3, hours=15, minutes=30)
    events = [
        {
            "start": _local_dt_to_ms(mon_10),
            "end": _local_dt_to_ms(mon_1130),
            "title": "Preview · studio (mock data)",
            "allDay": False,
            "calendar": "studio",
        },
        {
            "start": _local_dt_to_ms(thu_14),
            "end": _local_dt_to_ms(thu_1530),
            "title": "Preview · personal (mock data)",
            "allDay": False,
            "calendar": "personal",
        },
    ]
    payload = {
        "status": "success",
        "timeZone": "America/Chicago",
        "calView": "week",
        "days": days,
        "events": events,
    }
    return json.dumps(payload).encode("utf-8")


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

        if (
            method == "POST"
            and body
            and os.environ.get("LOCAL_MOCK_ADMIN_CALENDAR", "1") != "0"
        ):
            mocked = mock_owner_calendar_week_if_applicable(body)
            if mocked is not None:
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(mocked)))
                self.send_header("Cache-Control", "no-store")
                self.send_header("X-Local-Preview", "mock-owner-calendar-week")
                self.end_headers()
                self.wfile.write(mocked)
                return

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
    if os.environ.get("LOCAL_MOCK_ADMIN_CALENDAR", "1") != "0":
        print("Admin week calendar: MOCK data (no deploy needed). Set LOCAL_MOCK_ADMIN_CALENDAR=0 to use live API.")
    print("Ctrl+C to stop")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()
