"""
Append stdout/stderr to Admin/Logs for scheduled and manual script runs.

Task Scheduler often invokes the .py directly (no .bat redirect), so scripts
should install this at the start of main().
"""

from __future__ import annotations

import os
import sys
from datetime import datetime
from typing import TextIO


_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(_SCRIPT_DIR)
LOG_DIR = os.path.join(_BASE_DIR, "Admin", "Logs")


class _Tee:
    """Write to the original stream and a log file."""

    def __init__(self, stream: TextIO, log_file: TextIO) -> None:
        self._stream = stream
        self._log_file = log_file

    def write(self, data: str) -> int:
        self._stream.write(data)
        self._log_file.write(data)
        self._log_file.flush()
        return len(data)

    def flush(self) -> None:
        self._stream.flush()
        self._log_file.flush()

    def isatty(self) -> bool:
        return False


def install_run_log(job_name: str) -> str:
    """
    Tee stdout/stderr to Admin/Logs/{job_name}.log (append).
    Returns the log file path.
    """
    os.makedirs(LOG_DIR, exist_ok=True)
    log_path = os.path.join(LOG_DIR, f"{job_name}.log")
    log_file = open(log_path, "a", encoding="utf-8", newline="\n", buffering=1)
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_file.write(f"\n===== {stamp} | {job_name} =====\n")
    log_file.flush()

    sys.stdout = _Tee(sys.__stdout__, log_file)  # type: ignore[assignment]
    sys.stderr = _Tee(sys.__stderr__, log_file)  # type: ignore[assignment]
    print(f"Logging to {log_path}")
    return log_path