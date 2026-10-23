"""Minimal local scheduler.

Watches `signo_folder` for new/modified Excel files and imports them
automatically. Detects documents that need reminders based on `days_advance`.

Design constraints:
- Single background thread (process-safe for single-worker uvicorn).
- Import lock prevents simultaneous imports.
- Tracks imported file hashes to avoid reimporting the same file.
- Reminder detection is advisory: draft creation remains an explicit action.
- No automatic email sending.
"""
from __future__ import annotations

import os
import threading
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Optional

import msvcrt

from database import APP_DATA_DIR

_lock = threading.Lock()
_import_lock = threading.Lock()
_thread: Optional[threading.Thread] = None
_running = False
_lock_handle = None
_lock_file = APP_DATA_DIR / "scheduler.lock"

_status: Dict[str, Any] = {
    "running": False,
    "last_check": None,
    "last_import": None,
    "errors": [],
    "pending_reminders": 0,
    "owner_pid": None,
}

# Set of (path, mtime) already imported this session – persistent protection
# uses the imports table (file_hash).
_seen_files: set = set()


def get_scheduler_status() -> Dict[str, Any]:
    with _lock:
        return dict(_status)


def _append_error(message: str) -> None:
    with _lock:
        errors = list(_status.get("errors", []))
        errors.append(message)
        _status["errors"] = errors[-10:]


def _acquire_process_lock() -> bool:
    global _lock_handle
    APP_DATA_DIR.mkdir(parents=True, exist_ok=True)
    handle = open(_lock_file, "a+b")
    try:
        handle.seek(0)
        msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
    except OSError:
        handle.close()
        return False
    _lock_handle = handle
    return True


def _release_process_lock() -> None:
    global _lock_handle
    handle = _lock_handle
    if handle is None:
        return
    try:
        handle.seek(0)
        msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
    except OSError:
        pass
    try:
        handle.close()
    except OSError:
        pass
    _lock_handle = None


def _check_folder(cfg: Dict[str, Any]) -> None:
    """Run one iteration: import new files, update reminder count."""
    from services import import_signo_file
    from database import connect

    folder_path = cfg.get("signo_folder", "")
    days_advance = int(cfg.get("days_advance", 7))

    if not folder_path:
        return

    folder = Path(folder_path)
    if not folder.is_dir():
        with _lock:
            _status["errors"] = [f"Carpeta no encontrada: {folder_path}"]
        return

    # --- Import new/modified Excel files ---
    for xlsx in sorted(folder.glob("*.xlsx")):
        if xlsx.name.startswith("~$"):
            continue
        try:
            mtime = xlsx.stat().st_mtime
        except OSError:
            continue
        fingerprint = (str(xlsx), mtime)
        if fingerprint in _seen_files:
            continue

        with _import_lock:
            if fingerprint in _seen_files:
                continue
            try:
                import_signo_file(xlsx)
                _seen_files.add(fingerprint)
                with _lock:
                    _status["last_import"] = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
            except Exception as exc:
                _append_error(f"Error importando {xlsx.name}: {exc}")

    # --- Count pending reminders ---
    cutoff = (datetime.now() + timedelta(days=days_advance)).date().isoformat()
    try:
        with connect() as conn:
            count = conn.execute(
                "SELECT COUNT(*) FROM documents"
                " WHERE estado IN ('PENDIENTE', 'SALDO_PARCIAL')"
                " AND CAST(saldo_observado AS REAL) > 0"
                " AND revision_required = 0"
                " AND (fecha_vencimiento IS NULL OR fecha_vencimiento <= ?)",
                (cutoff,),
            ).fetchone()[0]
        with _lock:
            _status["pending_reminders"] = count
    except Exception as exc:
        _append_error(f"Error contando recordatorios: {exc}")


def _scheduler_loop(cfg: Dict[str, Any]) -> None:
    global _running
    interval = max(1, int(cfg.get("check_frequency_minutes", 15))) * 60
    while _running:
        try:
            _check_folder(cfg)
        except Exception as exc:
            _append_error(f"Error en ciclo de revisión: {exc}")
        with _lock:
            _status["last_check"] = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
        # Sleep in 5-second segments so we can exit promptly
        for _ in range(int(interval / 5)):
            if not _running:
                break
            time.sleep(5)
    with _lock:
        _status["running"] = False


def start_scheduler(cfg: Dict[str, Any]) -> None:
    """Start (or restart) the background scheduler with the given config."""
    global _thread, _running

    # Stop any existing scheduler first
    stop_scheduler()

    if not _acquire_process_lock():
        with _lock:
            _status["running"] = False
            _status["owner_pid"] = None
            _status["errors"] = [
                "No se inició el scheduler: otro proceso ya está ejecutando la revisión periódica.",
            ]
        return

    _running = True
    with _lock:
        _status["running"] = True
        _status["errors"] = []
        _status["owner_pid"] = os.getpid()

    t = threading.Thread(target=_scheduler_loop, args=(cfg,), daemon=True, name="signo-scheduler")
    t.start()
    _thread = t


def stop_scheduler() -> None:
    global _running, _thread
    _running = False
    if _thread is not None:
        _thread.join(timeout=15)
        _thread = None
    _release_process_lock()
    with _lock:
        _status["running"] = False
        _status["owner_pid"] = None
