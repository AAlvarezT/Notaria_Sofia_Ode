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

import threading
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional

_lock = threading.Lock()
_import_lock = threading.Lock()
_thread: Optional[threading.Thread] = None
_running = False

_status: Dict[str, Any] = {
    "running": False,
    "last_check": None,
    "last_import": None,
    "errors": [],
    "pending_reminders": 0,
}

# Set of (path, mtime) already imported this session – persistent protection
# uses the imports table (file_hash).
_seen_files: set = set()


def get_scheduler_status() -> Dict[str, Any]:
    with _lock:
        return dict(_status)


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

    with _lock:
        _status["errors"] = []

    # --- Import new/modified Excel files ---
    for xlsx in sorted(folder.glob("*.xlsx")):
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
                with _lock:
                    _status["errors"].append(f"Error importando {xlsx.name}: {exc}")

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
        with _lock:
            _status["errors"].append(f"Error contando recordatorios: {exc}")


def _scheduler_loop(cfg: Dict[str, Any]) -> None:
    global _running
    interval = max(1, int(cfg.get("check_frequency_minutes", 15))) * 60
    while _running:
        try:
            _check_folder(cfg)
        except Exception as exc:
            with _lock:
                _status["errors"].append(f"Error en ciclo de revisión: {exc}")
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

    _running = True
    with _lock:
        _status["running"] = True
        _status["errors"] = []

    t = threading.Thread(target=_scheduler_loop, args=(cfg,), daemon=True, name="signo-scheduler")
    t.start()
    _thread = t


def stop_scheduler() -> None:
    global _running, _thread
    _running = False
    if _thread is not None:
        _thread.join(timeout=15)
        _thread = None
    with _lock:
        _status["running"] = False
