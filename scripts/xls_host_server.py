#!/usr/bin/env python3
"""
Small host-side HTTP wrapper for the real .xls flow.

N8N runs in Docker, but Excel COM must run on the Windows/WSL host through
powershell.exe. This server is started on the host and calls run_xls_host.sh.
"""

import json
import os
import subprocess
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse


ROOT_DIR = Path(__file__).resolve().parents[1]
SCRIPT_PATH = ROOT_DIR / "scripts" / "run_xls_host.sh"
PAYLOAD_PATH = ROOT_DIR / "data" / "output" / "tekla_payload.json"
DATA_OUTPUT_DIR = ROOT_DIR / "data" / "output"


def resolve_repo_path(value, default_path):
    if not value:
        return default_path
    path = Path(value)
    if not path.is_absolute():
        path = ROOT_DIR / path
    return path


def n8n_output_path(host_path):
    try:
        rel = host_path.resolve().relative_to(DATA_OUTPUT_DIR.resolve())
        return "/data/output/" + rel.as_posix()
    except ValueError:
        return str(host_path)


def read_payload():
    if not PAYLOAD_PATH.exists():
        return {}
    with PAYLOAD_PATH.open(encoding="utf-8") as f:
        return json.load(f)


def read_json_from_stdout(stdout):
    for line in (stdout or "").splitlines():
        line = line.strip()
        if not line.startswith("{"):
            continue
        try:
            return json.loads(line)
        except json.JSONDecodeError:
            continue
    return {}


def run_xls_flow(options):
    source_folder = resolve_repo_path(
        options.get("source_folder") or os.environ.get("TEKLA_SOURCE_FOLDER"),
        ROOT_DIR / "data" / "tekla",
    )
    template_xls = resolve_repo_path(
        options.get("template_xls") or os.environ.get("TEMPLATE_XLS"),
        ROOT_DIR / "data" / "macro" / "Allied_Macro_original.xls",
    )
    output_dir = resolve_repo_path(
        options.get("output_dir") or os.environ.get("OUTPUT_DIR"),
        DATA_OUTPUT_DIR,
    )
    timeout = int(options.get("timeout_seconds") or os.environ.get("XLS_HOST_TIMEOUT", "600"))

    started = datetime.now()
    cmd = [str(SCRIPT_PATH), str(source_folder), str(template_xls), str(output_dir)]
    result = subprocess.run(
        cmd,
        cwd=ROOT_DIR,
        capture_output=True,
        text=True,
        timeout=timeout,
    )

    payload = read_payload()
    stdout_payload = read_json_from_stdout(result.stdout)
    if stdout_payload:
        payload = {**payload, **stdout_payload}
    duration = int((datetime.now() - started).total_seconds())

    if result.returncode != 0:
        if payload.get("status") == "no_files":
            return {
                "status": "no_files",
                "message": payload.get("message", "No hay archivos .xls para procesar"),
                "job_number": payload.get("job_number"),
                "files_processed": payload.get("files_processed", 0),
                "files_with_errors": payload.get("files_with_errors", []),
                "payload_file": n8n_output_path(PAYLOAD_PATH),
                "duration_seconds": duration,
                "log_entries": payload.get("log_entries", []),
                "stdout": result.stdout,
                "stderr": result.stderr,
            }, 200

        return {
            "status": "error",
            "message": "El flujo host .xls fallo",
            "returncode": result.returncode,
            "duration_seconds": duration,
            "stdout": result.stdout,
            "stderr": result.stderr,
            "files_with_errors": payload.get("files_with_errors", []),
            "log_entries": payload.get("log_entries", []),
        }, 500

    job_number = payload.get("job_number") or "SIN_NUMERO"
    output_host_path = output_dir / f"{job_number}_Secondary_Shipper.xls"

    response = {
        "status": payload.get("status", "success"),
        "job_number": job_number,
        "files_processed": payload.get("files_processed", 0),
        "files_with_errors": payload.get("files_with_errors", []),
        "output_file": n8n_output_path(output_host_path),
        "host_output_file": str(output_host_path),
        "payload_file": n8n_output_path(PAYLOAD_PATH),
        "duration_seconds": duration,
        "log_entries": payload.get("log_entries", []),
        "stdout": result.stdout,
        "stderr": result.stderr,
    }
    return response, 200


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        print("%s - %s" % (self.address_string(), fmt % args), flush=True)

    def send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def read_json_body(self):
        length = int(self.headers.get("Content-Length") or "0")
        if length <= 0:
            return {}
        body = self.rfile.read(length).decode("utf-8")
        if not body.strip():
            return {}
        return json.loads(body)

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/":
            self.send_json({
                "status": "ok",
                "mode": "xls_host",
                "message": "Servidor host .xls activo",
                "endpoints": {
                    "health": "/health",
                    "run": "/run"
                }
            })
            return
        if path == "/health":
            self.send_json({"status": "ok", "mode": "xls_host"})
            return
        if path == "/run":
            self.handle_run({})
            return
        self.send_json({"status": "not_found", "message": "Use /health or /run"}, 404)

    def do_POST(self):
        path = urlparse(self.path).path
        if path != "/run":
            self.send_json({"status": "not_found", "message": "Use /run"}, 404)
            return
        try:
            self.handle_run(self.read_json_body())
        except json.JSONDecodeError as exc:
            self.send_json({"status": "error", "message": f"JSON invalido: {exc}"}, 400)

    def handle_run(self, options):
        try:
            data, status = run_xls_flow(options)
            self.send_json(data, status)
        except subprocess.TimeoutExpired as exc:
            self.send_json({
                "status": "timeout",
                "message": "El flujo host .xls tardo mas de lo permitido",
                "stdout": exc.stdout or "",
                "stderr": exc.stderr or "",
            }, 500)
        except Exception as exc:
            self.send_json({"status": "fatal_error", "message": str(exc)}, 500)


def main():
    host = os.environ.get("XLS_HOST_BIND", "0.0.0.0")
    port = int(os.environ.get("XLS_HOST_PORT", "5055"))
    server = ThreadingHTTPServer((host, port), Handler)
    print(f"Servidor host .xls escuchando en http://{host}:{port}", flush=True)
    server.serve_forever()


if __name__ == "__main__":
    main()
