# -*- coding: utf-8 -*-
"""
nadlan_ui.py
============
Local web UI for the Nadlan scraper.

Running this file starts a tiny local web server (Python standard library
only - no extra dependencies) and opens the HTML interface in your default
browser. The flow is:

    1. Press "Start" in the page  -> a browser window/tab opens at the Tax
       Authority start page (Microsoft Edge, dedicated profile).
    2. Log in and run your search in that tab until the results grid shows.
    3. Press "Run" in the page     -> the scraper verifies the grid, walks
       every page, extracts the data and saves output.xlsx.
    4. Press "Open Excel"          -> opens the saved file.

Why a server and not just an .html file:
    A web page opened from disk cannot launch browser automation by itself,
    so a small local backend owns the Playwright session and the HTML page
    talks to it over http://127.0.0.1.

The Playwright sync API must be used from one thread, so a single
BrowserController thread owns the session and processes commands from a
queue. The HTTP handlers only enqueue commands and read a shared snapshot.

Run:
    python nadlan_ui.py
"""

from __future__ import annotations

import json
import os
import queue
import threading
import traceback
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import nadlan_scraper as scraper
from playwright.sync_api import sync_playwright

HERE = os.path.dirname(os.path.abspath(__file__))
INDEX_HTML = os.path.join(HERE, "index.html")
HOST = "127.0.0.1"
PORT = 8765


# ======================================================================
#  Browser controller - owns the Playwright session in its own thread
# ======================================================================
class BrowserController(threading.Thread):
    def __init__(self):
        super().__init__(daemon=True)
        self._cmd_q: queue.Queue = queue.Queue()
        self._lock = threading.Lock()
        self._logs: list[str] = []
        self._state = "idle"          # idle|opening|ready|scraping|done|no_deals|error
        self._progress = {
            "current": 0, "total": 0, "deals": 0, "page": 0, "pages": 0,
            "page_deal": 0, "page_deals": 0,
        }
        self._stop = threading.Event()
        # NOTE: do not name this `self._context` - threading.Thread uses that
        # attribute internally for contextvars propagation (Python 3.13+).
        self._browser_ctx = None
        self._page = None
        # Path of the Excel produced by the most recent run (unique per run).
        self._last_output = None
        # Live table rows of the current/most recent run (list of 11-col lists).
        self._rows: list[list[str]] = []

    # -- shared-state writers (called from the controller thread) ----------
    def _log(self, msg):
        with self._lock:
            for line in str(msg).splitlines():
                self._logs.append(line)

    def _set_state(self, state):
        with self._lock:
            self._state = state

    def _set_progress(self, current, total, deals, page=0, pages=0,
                      page_deal=0, page_deals=0):
        with self._lock:
            self._progress = {
                "current": current, "total": total, "deals": deals,
                "page": page, "pages": pages,
                "page_deal": page_deal, "page_deals": page_deals,
            }

    def _reset_rows(self):
        with self._lock:
            self._rows = []

    def _add_row(self, row):
        with self._lock:
            self._rows.append(row)

    # -- shared-state reader (called from HTTP threads) --------------------
    def snapshot(self):
        with self._lock:
            last = self._last_output
            return {
                "state": self._state,
                "logs": self._logs[-600:],
                "progress": dict(self._progress),
                "output": bool(last and os.path.exists(last)),
                "headers": list(scraper.MAIN_HEADERS),
                "rows": [list(r) for r in self._rows],
            }

    def open_output(self) -> bool:
        """Open the most recent run's Excel file. Returns True on success."""
        with self._lock:
            last = self._last_output
        if last and os.path.exists(last):
            try:
                os.startfile(last)  # type: ignore[attr-defined]
                return True
            except Exception:  # noqa: BLE001
                return False
        return False

    # -- public commands (called from HTTP threads) ------------------------
    def cmd_open(self):
        self._cmd_q.put(("open",))

    def cmd_run(self):
        self._stop.clear()
        self._cmd_q.put(("run",))

    def cmd_stop(self):
        self._stop.set()

    # -- the controller loop ----------------------------------------------
    def run(self):
        scraper.set_logger(self._log)
        with sync_playwright() as p:
            while True:
                action = self._cmd_q.get()[0]
                if action == "shutdown":
                    break
                try:
                    if action == "open":
                        self._do_open(p)
                    elif action == "run":
                        self._do_run()
                except Exception as e:  # noqa: BLE001
                    self._log(f"[!] שגיאה: {e}")
                    self._log(traceback.format_exc())
                    self._set_state("error")

    def _do_open(self, p):
        self._set_state("opening")
        if self._browser_ctx is None:
            # First time: launch the dedicated-profile browser at the start page
            self._browser_ctx, self._page = scraper.open_browser_own_profile(p)
        else:
            # Already open: bring up a fresh tab on the same start URL
            self._page = self._browser_ctx.new_page()
            try:
                self._page.goto(scraper.START_URL, wait_until="domcontentloaded")
            except Exception:  # noqa: BLE001
                pass
            self._page.bring_to_front()
        self._set_state("ready")

    def _do_run(self):
        if self._page is None:
            self._log("[!] אין דפדפן פתוח. לחץ/י קודם על 'פתח דפדפן'.")
            self._set_state("idle")
            return
        self._set_state("validating")
        if not scraper.grid_is_ready(self._page):
            self._set_state("ready")
            return
        self._set_state("scraping")
        self._set_progress(0, 0, 0)
        self._reset_rows()
        # Unique output file per run, based on the current search + timestamp,
        # so running several searches produces several Excel files.
        out_path = scraper.make_output_path(HERE, scraper.get_search_info(self._page))
        with self._lock:
            self._last_output = out_path
        # Pass output_path so the Excel file is saved after every deal:
        # if anything breaks mid-run, the data collected so far is preserved.
        deals, info = scraper.run(
            self._page,
            on_progress=self._set_progress,
            should_stop=self._stop.is_set,
            output_path=out_path,
            on_deal=lambda main_row, raw: self._add_row(main_row),
        )
        if deals:
            scraper.write_excel(deals, out_path, info)
            self._log(f"[\u2713] נשמרו {len(deals)} עסקאות לקובץ {os.path.basename(out_path)}")
            self._set_state("done")
        else:
            self._set_state("no_deals")


# ======================================================================
#  HTTP request handler
# ======================================================================
class Handler(BaseHTTPRequestHandler):
    controller: BrowserController  # injected on the class before serving

    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        data = body if isinstance(body, bytes) else body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _send_json(self, obj, code=200):
        self._send(code, json.dumps(obj, ensure_ascii=False))

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            try:
                with open(INDEX_HTML, "rb") as f:
                    self._send(200, f.read(), "text/html; charset=utf-8")
            except FileNotFoundError:
                self._send(404, "index.html not found", "text/plain; charset=utf-8")
        elif self.path == "/api/status":
            self._send_json(self.controller.snapshot())
        else:
            self._send(404, "not found", "text/plain; charset=utf-8")

    def do_POST(self):
        if self.path == "/api/open":
            self.controller.cmd_open()
            self._send_json({"ok": True})
        elif self.path == "/api/run":
            self.controller.cmd_run()
            self._send_json({"ok": True})
        elif self.path == "/api/stop":
            self.controller.cmd_stop()
            self._send_json({"ok": True})
        elif self.path == "/api/open-excel":
            ok = self.controller.open_output()
            self._send_json({"ok": ok})
        else:
            self._send(404, "not found", "text/plain; charset=utf-8")

    # Silence the default per-request console logging
    def log_message(self, *_args):
        pass


def main():
    controller = BrowserController()
    controller.start()
    Handler.controller = controller

    httpd = ThreadingHTTPServer((HOST, PORT), Handler)
    url = f"http://{HOST}:{PORT}/"
    print(f"[i] Nadlan UI running at {url}")
    print("[i] Close this window (Ctrl+C) to stop the server.")
    webbrowser.open(url)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n[i] Shutting down.")
    finally:
        httpd.shutdown()


if __name__ == "__main__":
    main()
