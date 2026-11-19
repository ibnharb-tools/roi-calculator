#!/usr/bin/env python3
"""
Celestial Automation ROI calculator: local web app.

    python roi_calculator.py            # opens http://127.0.0.1:8765
    python roi_calculator.py --port 9000 --no-browser

Runs only on your machine (bound to 127.0.0.1). Nothing is sent anywhere.
Needs Python 3.9+; Excel export needs openpyxl (pip install openpyxl).
"""
from __future__ import annotations

import argparse
import datetime as dt
import io
import json
import os
import re
import sys
import threading
import traceback
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import roi_engine as E  # noqa: E402

STATIC = os.path.join(HERE, "static")
SCEN_DIR = os.path.join(HERE, "scenarios")
EXPORT_DIR = os.path.join(HERE, "exports")
AUTOSAVE = "_autosave.json"
os.makedirs(SCEN_DIR, exist_ok=True)
os.makedirs(EXPORT_DIR, exist_ok=True)

CONTENT_TYPES = {".html": "text/html; charset=utf-8", ".js": "application/javascript; charset=utf-8",
                 ".css": "text/css; charset=utf-8", ".png": "image/png", ".svg": "image/svg+xml",
                 ".ico": "image/x-icon"}


def _finite(o):
    """JSON has no Infinity/NaN; turn them into null."""
    if isinstance(o, float):
        return o if o == o and o not in (float("inf"), float("-inf")) else None
    if isinstance(o, dict):
        return {k: _finite(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)):
        return [_finite(v) for v in o]
    return o


def safe_name(name: str) -> str:
    base = os.path.basename(str(name or "")).strip()
    if base.lower().endswith(".json"):
        base = base[:-5]
    base = re.sub(r"[^A-Za-z0-9 _\-]", "", base).strip().replace(" ", "_")[:80]
    return (base or "scenario") + ".json"


def list_scenarios() -> list[dict]:
    out = []
    for f in sorted(os.listdir(SCEN_DIR)):
        if not f.endswith(".json"):
            continue
        p = os.path.join(SCEN_DIR, f)
        try:
            with open(p, encoding="utf-8") as fh:
                title = (json.load(fh).get("meta") or {}).get("title") or f
        except Exception:  # noqa: BLE001
            title = f + " (unreadable)"
        out.append({"file": f, "title": title,
                    "modified": dt.datetime.fromtimestamp(os.path.getmtime(p)).isoformat(timespec="minutes")})
    out.sort(key=lambda r: (r["file"] != AUTOSAVE, r["modified"]), reverse=False)
    return out


class Handler(BaseHTTPRequestHandler):
    server_version = "CelestialROI/1.0"

    def log_message(self, fmt, *args):  # quieter console
        if "/api/autosave" in (args[0] if args else ""):
            return
        sys.stderr.write("  %s\n" % (fmt % args))

    # -- helpers ---------------------------------------------------------
    def _send(self, code: int, body: bytes, ctype: str, extra: dict | None = None) -> None:
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _json(self, obj, code: int = 200) -> None:
        self._send(code, json.dumps(_finite(obj), allow_nan=False).encode("utf-8"), "application/json")

    def _body(self) -> dict:
        n = int(self.headers.get("Content-Length") or 0)
        if n > 5_000_000:
            raise ValueError("Request too large")
        raw = self.rfile.read(n) if n else b"{}"
        return json.loads(raw.decode("utf-8") or "{}")

    def _origin_ok(self) -> bool:
        # Block other websites in your browser from calling this local server.
        origin = self.headers.get("Origin")
        if not origin:
            return True
        host = self.headers.get("Host", "")
        return origin in (f"http://{host}",)

    # -- routes ------------------------------------------------------------
    def do_GET(self):  # noqa: N802
        url = urlparse(self.path)
        try:
            if url.path in ("/", "/index.html"):
                return self._static("index.html")
            if url.path.startswith("/static/"):
                return self._static(url.path[len("/static/"):])
            if url.path == "/api/schema":
                return self._json(E.schema())
            if url.path == "/api/scenarios":
                return self._json(list_scenarios())
            if url.path == "/api/scenario":
                f = safe_name(parse_qs(url.query).get("file", [""])[0])
                p = os.path.join(SCEN_DIR, f)
                if not os.path.exists(p):
                    return self._json({"error": "not found"}, 404)
                with open(p, encoding="utf-8") as fh:
                    return self._json({"file": f, "scenario": json.load(fh)})
            self._json({"error": "not found"}, 404)
        except Exception as exc:  # noqa: BLE001
            traceback.print_exc()
            self._json({"error": str(exc)}, 500)

    def do_POST(self):  # noqa: N802
        if not self._origin_ok():
            return self._json({"error": "forbidden origin"}, 403)
        url = urlparse(self.path)
        try:
            body = self._body()
            if url.path == "/api/calc":
                sc = body.get("scenario") or {}
                cands = sc.get("candidates") or [E.blank_candidate()]
                i = max(0, min(int(body.get("active", 0)), len(cands) - 1))
                mc_runs = max(0, min(int(body.get("mc_runs", 3000)), 20000))
                res = E.evaluate(cands[i], mc_runs=mc_runs)
                return self._json({"result": res, "compare": E.compare(cands)})
            if url.path in ("/api/scenario", "/api/autosave"):
                f = AUTOSAVE if url.path == "/api/autosave" else safe_name(body.get("file"))
                sc = body.get("scenario")
                if not isinstance(sc, dict) or not isinstance(sc.get("candidates"), list):
                    return self._json({"error": "bad scenario"}, 400)
                p = os.path.join(SCEN_DIR, f)
                tmp = p + ".tmp"
                with open(tmp, "w", encoding="utf-8") as fh:
                    json.dump(sc, fh, indent=2)
                os.replace(tmp, p)
                return self._json({"ok": True, "file": f})
            if url.path == "/api/export":
                return self._export(body)
            self._json({"error": "not found"}, 404)
        except Exception as exc:  # noqa: BLE001
            traceback.print_exc()
            self._json({"error": str(exc)}, 500)

    def _static(self, rel: str) -> None:
        p = os.path.normpath(os.path.join(STATIC, rel))
        if not p.startswith(STATIC) or not os.path.isfile(p):
            return self._json({"error": "not found"}, 404)
        with open(p, "rb") as fh:
            data = fh.read()
        self._send(200, data, CONTENT_TYPES.get(os.path.splitext(p)[1], "application/octet-stream"))

    def _export(self, body: dict) -> None:
        try:
            import excel_export as X
        except ImportError:
            return self._json({"error": "Excel export needs openpyxl. Run: pip install openpyxl"}, 500)
        sc = body.get("scenario") or {}
        cands = sc.get("candidates") or []
        if not cands:
            return self._json({"error": "no candidates"}, 400)
        i = max(0, min(int(body.get("active", 0)), len(cands) - 1))
        ev = E.evaluate(cands[i], mc_runs=5000)
        wb = X.build_workbook(cands[i], sc.get("meta") or {}, ev, E.compare(cands))
        buf = io.BytesIO()
        wb.save(buf)
        data = buf.getvalue()
        client = re.sub(r"[^A-Za-z0-9]+", "_", (sc.get("meta") or {}).get("client") or "client").strip("_")[:40]
        task = re.sub(r"[^A-Za-z0-9]+", "_", cands[i].get("name") or "candidate").strip("_")[:40]
        fname = f"ROI_{client}_{task}_{dt.date.today().isoformat()}.xlsx"
        with open(os.path.join(EXPORT_DIR, fname), "wb") as fh:
            fh.write(data)
        self._send(200, data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   {"Content-Disposition": f'attachment; filename="{fname}"', "X-Saved-To": fname})


def main() -> None:
    ap = argparse.ArgumentParser(description="Celestial Automation ROI calculator")
    ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--no-browser", action="store_true")
    args = ap.parse_args()
    server = None
    for port in range(args.port, args.port + 20):
        try:
            server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
            break
        except OSError:
            continue
    if server is None:
        sys.exit("No free port found.")
    url = f"http://127.0.0.1:{server.server_address[1]}/"
    print(f"\n  Celestial Automation ROI calculator running at {url}")
    print("  Local only. Press Ctrl+C to stop.\n")
    if not args.no_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n  Stopped.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
