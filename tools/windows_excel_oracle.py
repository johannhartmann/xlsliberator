#!/usr/bin/env python3
"""One-request-per-child JSON-lines supervisor for a trusted Windows host."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path


def _terminate_tree(process: subprocess.Popen[str], excel_pid_file: Path) -> None:
    if process.poll() is None:
        subprocess.run(
            ["taskkill", "/PID", str(process.pid), "/T", "/F"],
            capture_output=True,
            check=False,
            timeout=15,
        )
    if excel_pid_file.is_file():
        pid = excel_pid_file.read_text(encoding="ascii").strip()
        if pid.isdigit():
            subprocess.run(
                ["taskkill", "/PID", pid, "/T", "/F"],
                capture_output=True,
                check=False,
                timeout=15,
            )


def handle(line: str) -> str:
    request = json.loads(line)
    timeout = float(request.get("timeout_seconds", 180))
    with tempfile.TemporaryDirectory(prefix="xlsliberator-excel-supervisor-") as directory:
        pid_file = Path(directory) / "excel.pid"
        request["excel_pid_file"] = str(pid_file)
        process = subprocess.Popen(
            [sys.executable, "-m", "xlsliberator.windows_excel_worker"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        try:
            stdout, stderr = process.communicate(json.dumps(request), timeout=timeout)
        except subprocess.TimeoutExpired:
            _terminate_tree(process, pid_file)
            return json.dumps(
                {
                    "schema_version": "1.0.0",
                    "status": "failed",
                    "trace": None,
                    "attachments": {},
                    "error": {"type": "timeout", "message": f"Excel exceeded {timeout:g}s"},
                }
            )
        finally:
            if process.poll() is None:
                _terminate_tree(process, pid_file)
        if process.returncode not in {0, 1} and not stdout.strip():
            return json.dumps(
                {
                    "schema_version": "1.0.0",
                    "status": "failed",
                    "trace": None,
                    "attachments": {},
                    "error": {"type": "worker_exit", "message": stderr[-2000:]},
                }
            )
        lines = [item for item in stdout.splitlines() if item.strip()]
        if len(lines) != 1:
            raise RuntimeError("worker returned an invalid response count")
        return lines[0]


def main() -> int:
    for line in sys.stdin:
        if not line.strip():
            continue
        try:
            response = handle(line)
        except Exception as exc:
            response = json.dumps(
                {
                    "schema_version": "1.0.0",
                    "status": "failed",
                    "trace": None,
                    "attachments": {},
                    "error": {"type": "supervisor_failure", "message": str(exc)},
                }
            )
        sys.stdout.write(response + "\n")
        sys.stdout.flush()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
