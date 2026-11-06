"""Optional local server: serves the output folder and lets the dashboard start scan / dry-run / build jobs.

Security: binds to 127.0.0.1 only; the only thing it can run is `python -m resco scan|build` with validated arguments
(no shell, no arbitrary commands); paths must exist and have the expected extension.
"""
from __future__ import annotations

import json
import logging
import os
import re
import subprocess
import sys
import threading
import webbrowser
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse

from resco.dashboard import render_home
from resco.runs import DECK_LABEL, list_runs, name_says_uk
from resco.selection import Selection, SelectionError, preview

log = logging.getLogger(__name__)
PROJECT_ROOT = Path(__file__).resolve().parents[2]
LOB_RE = re.compile(r"^[A-Za-z0-9_,]*$")
PARTS_RE = re.compile(r"^[A-Za-z0-9_, \-]*$")
SLIDES_RE = re.compile(r"^[0-9, \-\u2013]*$")


def candidates(pattern: str, where: list[Path], uk: bool | None = None) -> list[Path]:
    """Files matching `pattern` in the folders, newest (by modification time) first; `uk` filters decks by 'UK' in the file name."""
    found = [f for folder in where for f in folder.glob(pattern) if f.is_file() and (uk is None or name_says_uk(f.name) == uk)]
    return sorted(found, key=lambda f: f.stat().st_mtime, reverse=True)


def pack_dir() -> Path:
    """The quarter's reserving pack folder (…/2026/202609/06_ReservingPack): RESCO_PACK_DIR when set (local.cmd), otherwise
    the folder above the project, which is where the tool used to sit inside the pack."""
    env = os.environ.get("RESCO_PACK_DIR", "").strip().strip('"')
    return Path(env) if env else PROJECT_ROOT.parent


def deck_folders() -> list[Path]:
    return [p for p in dict.fromkeys([PROJECT_ROOT / "Inputs", PROJECT_ROOT, pack_dir()]) if p.is_dir()]


def config_for(deck: str) -> Path:
    return PROJECT_ROOT / "config" / ("uk" if deck == "uk" else "")


class JobRunner:
    """One job at a time; keeps the log in memory for the dashboard to poll."""

    def __init__(self, out_dir: Path) -> None:
        self.out_dir = out_dir
        self.lock = threading.Lock()
        self.state = "idle"  # idle | running | done
        self.mode = ""
        self.deck = ""
        self.log: list[str] = []
        self.progress = 0.0
        self.exit_code: int | None = None
        self.run_dir: str | None = None

    def snapshot(self, defaults: dict[str, str]) -> dict[str, Any]:
        with self.lock:
            return {
                "state": self.state, "mode": self.mode, "deck": self.deck, "log": list(self.log[-200:]), "progress": self.progress,
                "exit_code": self.exit_code, "run_dir": self.run_dir, "defaults": defaults,
            }  # fmt: skip

    def start(self, params: dict[str, Any]) -> str | None:
        err = self._validate(params)
        if err:
            return err
        mode, deck = params["mode"], params["deck"]
        config = config_for(deck)
        if deck == "uk" and not (PROJECT_ROOT / "config" / "uk").exists():
            return "the UK configuration (config/uk) does not exist yet"
        cmd = [sys.executable, "-u", "-m", "resco"]
        if mode == "scan":
            cmd += ["scan", "--template", params["template"], "--out", str(self.out_dir), "--config", str(config)]
        else:
            cmd += ["build", "--workbook", params["workbook"], "--template", params["template"], "--out", str(self.out_dir), "--config", str(config)]
            if params["lobs"]:
                cmd += ["--lobs", params["lobs"]]
            if params.get("parts"):
                cmd += ["--parts", params["parts"]]
            if params.get("slides"):
                cmd += ["--slides", params["slides"]]
            if params.get("netting"):
                cmd += ["--netting", params["netting"]]
            if params.get("netting_prior"):
                cmd += ["--netting-prior", params["netting_prior"]]
            if params.get("ibnr"):
                cmd += ["--ibnr", params["ibnr"]]
            if mode == "dry-run":
                cmd += ["--dry-run"]
            elif mode == "check":
                cmd += ["--check"]
        with self.lock:
            if self.state == "running":
                return "a job is already running"
            self.state, self.mode, self.deck, self.log, self.progress, self.exit_code, self.run_dir = "running", mode, deck, [], 0.0, None, None
        threading.Thread(target=self._run, args=(cmd,), daemon=True).start()
        return None

    @staticmethod
    def _validate(p: dict[str, Any]) -> str | None:
        if p.get("mode") not in ("scan", "check", "dry-run", "build"):
            return "unknown mode"
        if p.get("deck") not in ("main", "uk"):
            return "unknown deck"
        if not LOB_RE.match(p.get("lobs", "")):
            return "LoB list may only contain letters, digits, underscores and commas"
        if not PARTS_RE.match(p.get("parts", "") or "") or not SLIDES_RE.match(p.get("slides", "") or ""):
            return "the content types / slide numbers contain characters that are not allowed"
        try:
            Selection.parse(p.get("lobs", ""), p.get("parts", ""), p.get("slides", ""))
        except SelectionError as exc:
            return str(exc)
        if not Path(p.get("template", "")).is_file() or Path(p["template"]).suffix.lower() != ".pptx":
            return "template deck not found (.pptx)"
        if p["deck"] == "main" and name_says_uk(p["template"]):
            return f"The template file name says UK, but '{DECK_LABEL['main']}' is selected. Choose the UK branch deck, or pick the Lux template."
        if p["deck"] == "uk" and not name_says_uk(p["template"]):
            return f"The template file name does not say UK, but '{DECK_LABEL['uk']}' is selected. Choose the Lux deck, or pick the UK template."
        if p["mode"] != "scan" and (not Path(p.get("workbook", "")).is_file() or Path(p["workbook"]).suffix.lower() not in (".xlsm", ".xlsx")):
            return "workbook not found (.xlsm/.xlsx)"
        for key, label in (("netting", "netting workbook"), ("netting_prior", "Q-1 netting workbook"), ("ibnr", "IBNR allocation workbook")):
            value = p.get(key) or ""
            if value and (not Path(value).is_file() or Path(value).suffix.lower() not in (".xlsm", ".xlsx")):
                return f"{label} not found (.xlsm/.xlsx)"
        return None

    def _run(self, cmd: list[str]) -> None:
        env = {"PYTHONPATH": str(PROJECT_ROOT / "src"), "PYTHONIOENCODING": "utf-8"}
        import os

        proc = subprocess.Popen(
            cmd, cwd=PROJECT_ROOT, env={**os.environ, **env}, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace", bufsize=1,
        )  # fmt: skip
        assert proc.stdout is not None
        for line in proc.stdout:
            line = line.rstrip()
            with self.lock:
                self.log.append(line)
                m = re.search(r"PROGRESS (\d+)/(\d+)", line)
                if m:
                    self.progress = (int(m.group(1)) - 1) / int(m.group(2)) * 90
                m = re.search(r"(?:->|written:)\s*(.+?)\s*$", line)
                if m:
                    self.run_dir = Path(m.group(1)).name
        code = proc.wait()
        with self.lock:
            self.exit_code, self.state, self.progress = code, "done", 100.0


class DeckPreview:
    """Scans a template once (cached by file, modification time and deck) and answers 'what would this selection touch?'."""

    def __init__(self) -> None:
        self.lock = threading.Lock()
        self.cache: dict[tuple[str, float, str], Any] = {}

    def answer(self, query: dict[str, list[str]]) -> dict[str, Any]:
        one = lambda k: (query.get(k) or [""])[0]  # noqa: E731
        deck, template = one("deck") or "main", Path(one("template"))
        if deck not in ("main", "uk"):
            return {"error": "unknown deck"}
        if not template.is_file() or template.suffix.lower() != ".pptx":
            return {"error": "template deck not found (.pptx)"}
        lobs, parts, slides = one("lobs"), one("parts"), one("slides")
        if not (LOB_RE.match(lobs) and PARTS_RE.match(parts) and SLIDES_RE.match(slides)):
            return {"error": "the selection contains characters that are not allowed"}
        try:
            sel = Selection.parse(lobs, parts, slides)
        except SelectionError as exc:
            return {"error": str(exc)}
        from resco.config import load_deck_rules, load_lobs, load_ranges, load_table_config
        import tempfile

        from resco.deck_scan import TemplateLocked, build_map, copy_for_reading, scan_deck

        config = config_for(deck)
        key = (str(template.resolve()), template.stat().st_mtime, deck)
        with self.lock:  # one scan at a time; a repeated request for the same file is instant
            dm = self.cache.get(key)
            if dm is None:
                rules = load_deck_rules(config)
                with tempfile.TemporaryDirectory(prefix="resco_preview_") as tmp:
                    try:
                        scans = scan_deck(copy_for_reading(template, Path(tmp), tries=2, wait=1.0), rules)
                    except TemplateLocked as exc:
                        return {"error": str(exc)}
                dm = build_map(scans, load_lobs(config), load_ranges(config).life_resclasses, rules, load_table_config(config).targets)
                self.cache = {key: dm}
        from resco.config import load_slides_to_add
        from resco.slide_adder import missing_specs

        todo = missing_specs([m.title for m in dm.mapped] + [sc.title for sc in dm.narrative], [sp for sp in load_slides_to_add(config) if sel.allows_new(sp.lob, sp.part)])
        result = preview(dm, sel, [{"title": sp.title, "lob": sp.lob} for sp in todo])
        mtime = template.stat().st_mtime
        newer = [c for c in candidates("ReservingPresentation*.pptx", deck_folders(), uk=(deck == "uk")) if c.stat().st_mtime > mtime + 2 and c.resolve() != template.resolve()]
        from datetime import datetime

        result["template"] = {
            "name": template.name, "modified": datetime.fromtimestamp(mtime).isoformat(timespec="minutes"),
            "newer": [{"path": str(c), "name": c.name, "modified": datetime.fromtimestamp(c.stat().st_mtime).isoformat(timespec="minutes")} for c in newer[:3]],
        }
        return result


def make_server(out_dir: Path, port: int, defaults: dict[str, str]) -> tuple[ThreadingHTTPServer, JobRunner]:
    out_dir.mkdir(parents=True, exist_ok=True)
    runner = JobRunner(out_dir.resolve())
    previews = DeckPreview()

    class Handler(SimpleHTTPRequestHandler):
        def log_message(self, fmt: str, *args: Any) -> None:  # keep the console quiet
            pass

        def _json(self, obj: Any, status: int = 200) -> None:
            body = json.dumps(obj).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json")
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self) -> None:  # noqa: N802
            if self.path.startswith("/api/status"):
                return self._json(runner.snapshot(defaults))
            if self.path.startswith("/api/preview"):
                try:
                    return self._json(previews.answer(parse_qs(urlparse(self.path).query)))
                except Exception as exc:  # noqa: BLE001 - a broken template must not take the page down
                    log.exception("preview")
                    return self._json({"error": f"could not read the template: {exc}"})
            if self.path.startswith("/api/runs"):
                return self._json({"runs": list_runs(out_dir)})
            if self.path in ("/", "/index.html"):
                body = render_home().encode("utf-8")  # never a stale result: the home page only starts and lists runs
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Cache-Control", "no-store")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
                return
            return super().do_GET()

        def do_POST(self) -> None:  # noqa: N802
            if self.path != "/api/run":
                return self._json({"error": "not found"}, 404)
            length = int(self.headers.get("Content-Length", "0"))
            if length > 10_000:
                return self._json({"error": "request too large"}, 413)
            try:
                params = json.loads(self.rfile.read(length) or b"{}")
            except json.JSONDecodeError:
                return self._json({"error": "bad request"}, 400)
            err = runner.start(params)
            return self._json({"error": err} if err else {"ok": True}, 400 if err else 200)

    return ThreadingHTTPServer(("127.0.0.1", port), partial(Handler, directory=str(out_dir))), runner


def serve(out_dir: Path, port: int, defaults: dict[str, str], open_browser: bool = True) -> int:
    server, _ = make_server(out_dir, port, defaults)
    url = f"http://127.0.0.1:{port}/"
    log.info("dashboard at %s (Ctrl+C to stop)", url)
    if open_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0
