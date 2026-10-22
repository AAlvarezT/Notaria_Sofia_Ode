from __future__ import annotations

import json
import os
import sqlite3
from pathlib import Path
from typing import Any, Dict, List, Optional


APP_DATA_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")) / "notaria_sofia_ode" / "creditos_web_local"
DB_PATH = APP_DATA_DIR / "creditos.sqlite3"
CONFIG_PATH = APP_DATA_DIR / "config.json"
BACKUP_DIR = APP_DATA_DIR / "backups"


def ensure_base_dirs() -> None:
    APP_DATA_DIR.mkdir(parents=True, exist_ok=True)
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)


def get_db_path() -> Path:
    ensure_base_dirs()
    return DB_PATH


def connect(db_path: Optional[str | Path] = None) -> sqlite3.Connection:
    if db_path is None:
        db_path = get_db_path()
    else:
        db_path = Path(db_path)
        db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db(db_path: Optional[str | Path] = None) -> None:
    if db_path is None:
        ensure_base_dirs()
        db_path = get_db_path()
    else:
        db_path = Path(db_path)
        db_path.parent.mkdir(parents=True, exist_ok=True)
    with connect(db_path) as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS imports (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_name TEXT NOT NULL,
                file_hash TEXT NOT NULL,
                file_path TEXT,
                imported_at TEXT NOT NULL,
                cut_off TEXT,
                status TEXT NOT NULL,
                errors TEXT,
                rows_seen INTEGER DEFAULT 0,
                rows_saved INTEGER DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS documents (
                document_key TEXT PRIMARY KEY,
                tipo TEXT,
                serie TEXT,
                numero TEXT,
                doi TEXT,
                cliente TEXT,
                fecha_emision TEXT,
                importe_facturado TEXT,
                saldo_observado TEXT,
                forma_pago TEXT,
                fecha_vencimiento TEXT,
                empleado TEXT,
                fecha_actualizacion TEXT,
                estado TEXT,
                fuente TEXT,
                ultima_importacion_id INTEGER,
                revision_required INTEGER DEFAULT 0,
                observaciones TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                FOREIGN KEY(ultima_importacion_id) REFERENCES imports(id)
            );

            CREATE TABLE IF NOT EXISTS contacts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_key TEXT,
                nombre TEXT,
                correo TEXT,
                telefono TEXT,
                rol TEXT,
                es_principal INTEGER DEFAULT 0,
                UNIQUE(document_key, correo)
            );

            CREATE TABLE IF NOT EXISTS comments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_key TEXT NOT NULL,
                texto TEXT NOT NULL,
                autor TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(document_key) REFERENCES documents(document_key)
            );

            CREATE TABLE IF NOT EXISTS payment_commitments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_key TEXT NOT NULL,
                cliente TEXT,
                monto TEXT,
                fecha_compromiso TEXT,
                proxima_gestion TEXT,
                responsable TEXT,
                descripcion TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(document_key) REFERENCES documents(document_key)
            );

            CREATE TABLE IF NOT EXISTS draft_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_key TEXT,
                cliente TEXT,
                destinatario TEXT,
                subject TEXT,
                outlook_message_id TEXT,
                draft_status TEXT,
                created_at TEXT NOT NULL,
                html_preview TEXT,
                UNIQUE(document_key, subject)
            );

            CREATE TABLE IF NOT EXISTS review_queue (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                document_key TEXT,
                kind TEXT,
                detail TEXT,
                created_at TEXT NOT NULL,
                resolved INTEGER DEFAULT 0,
                UNIQUE(document_key, kind, detail)
            );
            """
        )


def set_setting(key: str, value: Any) -> None:
    with connect() as conn:
        conn.execute(
            "INSERT INTO settings(key, value) VALUES(?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (key, json.dumps(value, ensure_ascii=False)),
        )


def get_setting(key: str, default: Any = None) -> Any:
    with connect() as conn:
        row = conn.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
    if row is None:
        return default
    try:
        return json.loads(row["value"])
    except (TypeError, json.JSONDecodeError):
        return row["value"]


def save_config(data: Dict[str, Any]) -> None:
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    for key, value in data.items():
        set_setting(key, value)


def load_config() -> Dict[str, Any]:
    if CONFIG_PATH.exists():
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
        except json.JSONDecodeError:
            pass
    with connect() as conn:
        rows = conn.execute("SELECT key, value FROM settings").fetchall()
    data: Dict[str, Any] = {}
    for row in rows:
        try:
            data[row["key"]] = json.loads(row["value"])
        except json.JSONDecodeError:
            data[row["key"]] = row["value"]
    return data


def list_documents(filters: Optional[Dict[str, Any]] = None) -> List[sqlite3.Row]:
    query = "SELECT * FROM documents WHERE 1=1"
    params: List[Any] = []
    if filters:
        if filters.get("cliente"):
            query += " AND cliente LIKE ?"
            params.append(f"%{filters['cliente']}%")
        if filters.get("documento"):
            query += " AND (numero LIKE ? OR document_key LIKE ?)"
            params.append(f"%{filters['documento']}%")
            params.append(f"%{filters['documento']}%")
        if filters.get("responsable"):
            query += " AND empleado LIKE ?"
            params.append(f"%{filters['responsable']}%")
        if filters.get("estado"):
            query += " AND estado = ?"
            params.append(filters["estado"])
        if filters.get("from_due"):
            query += " AND fecha_vencimiento >= ?"
            params.append(filters["from_due"])
        if filters.get("to_due"):
            query += " AND fecha_vencimiento <= ?"
            params.append(filters["to_due"])
    query += " ORDER BY fecha_vencimiento IS NULL, fecha_vencimiento ASC, updated_at DESC"
    with connect() as conn:
        return conn.execute(query, params).fetchall()


def get_document(document_key: str) -> Optional[sqlite3.Row]:
    with connect() as conn:
        return conn.execute("SELECT * FROM documents WHERE document_key = ?", (document_key,)).fetchone()


def get_dashboard_summary() -> Dict[str, Any]:
    with connect() as conn:
        stats = {
            "deuda_pendiente": conn.execute("SELECT COALESCE(SUM(CAST(saldo_observado AS REAL)), 0) FROM documents WHERE estado IN ('PENDIENTE', 'SALDO_PARCIAL')").fetchone()[0],
            "documentos_revision": conn.execute("SELECT COUNT(*) FROM documents WHERE revision_required = 1 OR estado = 'EN_REVISION'").fetchone()[0],
            "documentos_urgentes": conn.execute("SELECT COUNT(*) FROM documents WHERE estado IN ('PENDIENTE', 'SALDO_PARCIAL') AND fecha_vencimiento IS NOT NULL AND fecha_vencimiento <= date('now','+7 days')").fetchone()[0],
            "ultima_importacion": conn.execute("SELECT imported_at FROM imports ORDER BY imported_at DESC LIMIT 1").fetchone(),
        }
        return stats


def backup_database() -> Path:
    ensure_base_dirs()
    db_path = get_db_path()
    if not db_path.exists():
        raise FileNotFoundError("No existe la base de datos para respaldar.")
    ts = __import__("datetime").datetime.now().strftime("%Y%m%d_%H%M%S")
    target = BACKUP_DIR / f"creditos_{ts}.sqlite3"
    import shutil

    shutil.copy2(db_path, target)
    return target
