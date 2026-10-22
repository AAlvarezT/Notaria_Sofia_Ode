from __future__ import annotations

import hashlib
import json
import re
import sqlite3
import unicodedata
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from openpyxl import load_workbook

from database import connect, get_db_path


def normalize_name(value: Any) -> str:
    text = str(value or "").strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.upper()
    text = re.sub(r"[^A-Z0-9]+", "", text)
    return text


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return re.sub(r"\s+", " ", text)


def parse_decimal(value: Any) -> Decimal:
    if value in (None, ""):
        return Decimal("0")
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    text = str(value).strip().replace("S/", "").replace("US$", "").replace("$", "")
    text = text.replace(" ", "")
    if text in ("", "-"):
        return Decimal("0")
    if text.count(",") and text.count("."):
        text = text.replace(".", "").replace(",", ".")
    elif text.count(","):
        text = text.replace(".", "").replace(",", ".")
    try:
        return Decimal(text)
    except InvalidOperation as exc:  # pragma: no cover - exercised by validation path
        raise ValueError(f"Monto no válido: {value!r}") from exc


def parse_date(value: Any) -> Optional[str]:
    if value in (None, "", "0000-00-00", "0000/00/00"):
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, str):
        text = clean_text(value)
        if not text:
            return None
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"):
            try:
                return datetime.strptime(text, fmt).date().isoformat()
            except ValueError:
                continue
        return text
    return str(value)


def document_key(tipo: Any, serie: Any, numero: Any) -> str:
    t = clean_text(tipo).upper()
    s = clean_text(serie).upper()
    n = clean_text(numero)
    return f"{t}|{s}|{n}"


def file_hash(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


HEADER_ALIASES = {
    "DOI": "DOI",
    "RUC": "DOI",
    "RUCDNI": "DOI",
    "CLIENTE": "CLIENTE",
    "GRUPO": "CLIENTE",
    "FECHAEMISION": "FECHAEMISION",
    "EMISION": "FECHAEMISION",
    "TIPO": "TIPO",
    "SERIE": "SERIE",
    "SERIEA": "SERIE",
    "SERIES": "SERIE",
    "NUMERO": "NUMERO",
    "DOC": "NUMERO",
    "DOCUMENTO": "NUMERO",
    "KARDEX": "KARDEX",
    "TRAMITE": "KARDEX",
    "IMPORTEFACTURADO": "IMPORTEFACTURADO",
    "FACTURADO": "IMPORTEFACTURADO",
    "SALDOPENDIENTE": "SALDOPENDIENTE",
    "SALDO": "SALDOPENDIENTE",
    "FORMAPAGO": "FORMAPAGO",
    "PAGO": "FORMAPAGO",
    "FECHAVENCIMIENTO": "FECHAVENCIMIENTO",
    "VENCIMIENTO": "FECHAVENCIMIENTO",
    "EMPLEADO": "EMPLEADO",
    "FECHAACTUALIZACION": "FECHAACTUALIZACION",
    "ACTUALIZACION": "FECHAACTUALIZACION",
    "FECHACANCELACION": "FECHACANCELACION",
    "CANCELACION": "FECHACANCELACION",
}

SHEET_ALIASES = {
    "VENTAS": "VENTAS",
    "PENDIENTES": "PENDIENTES",
    "CANCELADOS": "CANCELADOS",
    "RECUPERADOHOJA1": "PENDIENTES",
    "HOJARECUPERADA": "PENDIENTES",
}


def _normalize_header(value: Any) -> str:
    return normalize_name(value).replace("\n", "")


def _canonical_header(value: Any) -> str:
    header = _normalize_header(value)
    return HEADER_ALIASES.get(header, header)


def _find_header_row(ws, required: Sequence[str]) -> Tuple[int, Dict[str, int]]:
    required_set = set(required)
    for row in range(1, min(ws.max_row, 40) + 1):
        found: Dict[str, int] = {}
        for col in range(1, min(ws.max_column, 40) + 1):
            standard = _canonical_header(ws.cell(row, col).value)
            if standard in required_set:
                found[standard] = col
        if required_set.issubset(found):
            return row, found
    raise ValueError(f"No se encontró la cabecera esperada en la hoja {ws.title!r}.")


def _header_columns(ws, header_row: int) -> Dict[str, int]:
    columns: Dict[str, int] = {}
    for col in range(1, min(ws.max_column, 80) + 1):
        standard = _canonical_header(ws.cell(header_row, col).value)
        if standard:
            columns[standard] = col
    return columns


def _first_present(ws, row: int, cols: Dict[str, int], *keys: str) -> Any:
    for key in keys:
        if key in cols:
            return ws.cell(row, cols[key]).value
    return None


SHEET_ALIASES = {
    "VENTAS": {"#": "numero", "DOI": "doi", "CLIENTE": "cliente", "FECHAEMISION": "fecha_emision", "TIPO": "tipo", "SERIE": "serie", "NUMERO": "numero", "KARDEX": "kardex", "IMPORTEFACTURADO": "importe_facturado", "SALDOPENDIENTE": "saldo_pendiente", "FORMAPAGO": "forma_pago", "FECHAVENCIMIENTO": "fecha_vencimiento", "EMPLEADO": "empleado", "FECHAACTUALIZACION": "fecha_actualizacion"},
    "PENDIENTES": {"#": "numero", "DOI": "doi", "CLIENTE": "cliente", "FECHAEMISION": "fecha_emision", "TIPO": "tipo", "SERIE": "serie", "NUMERO": "numero", "KARDEX": "kardex", "IMPORTEFACTURADO": "importe_facturado", "SALDOPENDIENTE": "saldo_pendiente", "FECHAVENCIMIENTO": "fecha_vencimiento", "EMPLEADO": "empleado", "FECHAACTUALIZACION": "fecha_actualizacion"},
    "CANCELADOS": {"#": "numero", "DOI": "doi", "CLIENTE": "cliente", "FECHAEMISION": "fecha_emision", "TIPO": "tipo", "SERIE": "serie", "NUMERO": "numero", "KARDEX": "kardex", "IMPORTEFACTURADO": "importe_facturado", "SALDOPENDIENTE": "saldo_pendiente", "FORMAPAGO": "forma_pago", "FECHACANCELACION": "fecha_cancelacion", "FECHAVENCIMIENTO": "fecha_vencimiento", "EMPLEADO": "empleado", "FECHAACTUALIZACION": "fecha_actualizacion"},
}


def _read_sheet(ws, sheet_name: str, workbook_path: Path) -> Tuple[List[Dict[str, Any]], List[str]]:
    normalized_sheet = normalize_name(sheet_name)
    aliases = SHEET_ALIASES.get(normalized_sheet) or {}
    if not aliases:
        return [], []
    required = ["TIPO", "SERIE", "NUMERO"]
    try:
        header_row, _ = _find_header_row(ws, required)
    except ValueError:
        return [], [f"Hoja {sheet_name}: sin cabecera reconocida"]

    cols = _header_columns(ws, header_row)
    parsed_rows: List[Dict[str, Any]] = []
    errors: List[str] = []
    for row in range(header_row + 1, ws.max_row + 1):
        tipo = _first_present(ws, row, cols, "TIPO")
        serie = _first_present(ws, row, cols, "SERIE")
        numero = _first_present(ws, row, cols, "NUMERO")
        if tipo in (None, "") and serie in (None, "") and numero in (None, ""):
            continue
        record = {"tipo": clean_text(tipo), "serie": clean_text(serie), "numero": clean_text(numero), "document_key": document_key(tipo, serie, numero), "source_sheet": sheet_name}
        if not record["document_key"] or record["document_key"].endswith("||"):
            errors.append(f"{sheet_name}: fila {row} sin clave de comprobante válida")
            continue
        for col_key, field_name in aliases.items():
            value = _first_present(ws, row, cols, col_key) if col_key in cols else None
            if field_name == "fecha_emision":
                record[field_name] = parse_date(value)
            elif field_name in {"importe_facturado", "saldo_pendiente"}:
                try:
                    record[field_name] = str(parse_decimal(value))
                except ValueError:
                    errors.append(f"{sheet_name}: fila {row} tiene monto inválido en {field_name} ({value!r})")
                    record[field_name] = "0"
            else:
                record[field_name] = clean_text(value)
        parsed_rows.append(record)
    return parsed_rows, errors


def _document_sort_key(record: Dict[str, Any]) -> tuple:
    sheet_priority = {"PENDIENTES": 3, "VENTAS": 2, "CANCELADOS": 1}
    sheet_name = (record.get("source_sheet") or "").upper()
    fecha = record.get("fecha_actualizacion") or "0000-00-00"
    saldo = Decimal(str(record.get("saldo_pendiente") or "0"))
    return (sheet_priority.get(sheet_name, 0), fecha, saldo)


def _merge_document_rows(rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    if not rows:
        return {}
    merged = dict(rows[0])
    source_sheets: set[str] = set()
    for row in rows:
        source_sheets.add((row.get("source_sheet") or "").upper())
        for key in ("tipo", "serie", "numero", "doi", "cliente", "fecha_emision", "importe_facturado", "saldo_pendiente", "forma_pago", "fecha_vencimiento", "empleado", "fecha_actualizacion"):
            value = row.get(key)
            if value in (None, ""):
                continue
            existing = merged.get(key)
            if existing in (None, ""):
                merged[key] = value
                continue
            if key == "fecha_actualizacion":
                if str(value) > str(existing):
                    merged[key] = value
                continue
            if key in {"importe_facturado", "saldo_pendiente"}:
                try:
                    if Decimal(str(value)) > Decimal(str(existing)):
                        merged[key] = value
                except InvalidOperation:
                    pass
                continue
            if key == "cliente" and len(str(value)) > len(str(existing)):
                merged[key] = value
    merged["source_sheet"] = ", ".join(sorted(source_sheets))
    return merged


def _is_newer_report(existing: Optional[sqlite3.Row], candidate: Dict[str, Any]) -> bool:
    if existing is None:
        return True
    existing_date = existing["fecha_actualizacion"] if existing["fecha_actualizacion"] else "0000-00-00"
    candidate_date = candidate.get("fecha_actualizacion") or "0000-00-00"
    return str(candidate_date) >= str(existing_date)


def import_signo_file(path: str | Path, db_path: Optional[str | Path] = None) -> Dict[str, Any]:
    excel_path = Path(path)
    file_digest = file_hash(excel_path)
    conn = sqlite3.connect(str(db_path or get_db_path()))
    conn.row_factory = sqlite3.Row
    try:
        workbook = load_workbook(excel_path, read_only=True, data_only=True)
    except Exception as exc:  # pragma: no cover - file validation path
        raise ValueError(f"No se pudo abrir el archivo Excel: {exc}") from exc

    all_rows: List[Dict[str, Any]] = []
    errors: List[str] = []
    for names in workbook.sheetnames:
        ws = workbook[names]
        normalized = normalize_name(names)
        if normalized not in {"VENTAS", "PENDIENTES", "CANCELADOS"}:
            continue
        rows, sheet_errors = _read_sheet(ws, names, excel_path)
        all_rows.extend(rows)
        errors.extend(sheet_errors)
    workbook.close()

    doc_rows: Dict[str, List[Dict[str, Any]]] = {}
    for row in all_rows:
        key = row["document_key"]
        doc_rows.setdefault(key, []).append(row)

    for key in list(doc_rows):
        if len(doc_rows[key]) > 1:
            if any(item.get("source_sheet", "").upper() == "PENDIENTES" for item in doc_rows[key]):
                preferred = max(doc_rows[key], key=_document_sort_key)
                doc_rows[key] = [preferred]
            else:
                doc_rows[key] = [max(doc_rows[key], key=_document_sort_key)]

    import_at = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    import_id = conn.execute(
        "INSERT INTO imports(file_name, file_hash, file_path, imported_at, status, errors, rows_seen, rows_saved) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (excel_path.name, file_digest, str(excel_path), import_at, "OK" if not errors else "WARN", json.dumps(errors, ensure_ascii=False), len(all_rows), len(doc_rows)),
    ).lastrowid

    saved = 0
    for doc_key, rows in sorted(doc_rows.items()):
        row = _merge_document_rows(rows)
        existing = conn.execute("SELECT * FROM documents WHERE document_key = ?", (doc_key,)).fetchone()
        if existing is not None and not _is_newer_report(existing, row):
            continue
        state = determine_document_state(doc_key, row)
        doc = {
            "document_key": doc_key,
            "tipo": row.get("tipo", ""),
            "serie": row.get("serie", ""),
            "numero": row.get("numero", ""),
            "doi": row.get("doi", ""),
            "cliente": row.get("cliente", ""),
            "fecha_emision": row.get("fecha_emision"),
            "importe_facturado": row.get("importe_facturado", "0"),
            "saldo_observado": row.get("saldo_pendiente", "0"),
            "forma_pago": row.get("forma_pago", ""),
            "fecha_vencimiento": row.get("fecha_vencimiento"),
            "empleado": row.get("empleado", ""),
            "fecha_actualizacion": row.get("fecha_actualizacion"),
            "estado": state,
            "fuente": row.get("source_sheet", ""),
            "ultima_importacion_id": import_id,
            "revision_required": 1 if state in {"EN_REVISION", "SALDO_PARCIAL"} else 0,
            "observaciones": existing["observaciones"] if existing is not None else "",
            "created_at": existing["created_at"] if existing is not None else import_at,
            "updated_at": import_at,
        }
        conn.execute(
            "INSERT INTO documents(document_key, tipo, serie, numero, doi, cliente, fecha_emision, importe_facturado, saldo_observado, forma_pago, fecha_vencimiento, empleado, fecha_actualizacion, estado, fuente, ultima_importacion_id, revision_required, observaciones, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) ON CONFLICT(document_key) DO UPDATE SET tipo = excluded.tipo, serie = excluded.serie, numero = excluded.numero, doi = excluded.doi, cliente = excluded.cliente, fecha_emision = excluded.fecha_emision, importe_facturado = excluded.importe_facturado, saldo_observado = excluded.saldo_observado, forma_pago = excluded.forma_pago, fecha_vencimiento = CASE WHEN documents.fecha_vencimiento IS NOT NULL AND documents.fecha_vencimiento <> '' THEN documents.fecha_vencimiento ELSE excluded.fecha_vencimiento END, empleado = excluded.empleado, fecha_actualizacion = excluded.fecha_actualizacion, estado = excluded.estado, fuente = excluded.fuente, ultima_importacion_id = excluded.ultima_importacion_id, revision_required = excluded.revision_required, observaciones = CASE WHEN documents.observaciones IS NOT NULL AND documents.observaciones <> '' THEN documents.observaciones ELSE excluded.observaciones END, updated_at = excluded.updated_at",
            (
                doc["document_key"], doc["tipo"], doc["serie"], doc["numero"], doc["doi"], doc["cliente"], doc["fecha_emision"], doc["importe_facturado"], doc["saldo_observado"], doc["forma_pago"], doc["fecha_vencimiento"], doc["empleado"], doc["fecha_actualizacion"], doc["estado"], doc["fuente"], doc["ultima_importacion_id"], doc["revision_required"], doc["observaciones"], doc["created_at"], doc["updated_at"],
            ),
        )
        saved += 1
    conn.commit()
    conn.close()
    return {
        "import_id": import_id,
        "rows_seen": len(all_rows),
        "documents": saved,
        "errors": errors,
        "status": "OK" if not errors else "WARN",
    }


def determine_document_state(doc_key: str, row: Dict[str, Any]) -> str:
    saldo = Decimal(str(row.get("saldo_pendiente") or "0"))
    importe = Decimal(str(row.get("importe_facturado") or "0"))
    sheet = (row.get("source_sheet") or "").upper()
    if saldo <= 0:
        return "EXCLUIDO"
    if importe > 0 and saldo < importe:
        return "SALDO_PARCIAL"
    if sheet == "PENDIENTES":
        return "PENDIENTE"
    return "EN_REVISION"


def list_reviews() -> List[Dict[str, Any]]:
    with connect() as conn:
        rows = conn.execute("SELECT * FROM documents WHERE revision_required = 1 OR estado IN ('EN_REVISION', 'EXCLUIDO') ORDER BY updated_at DESC").fetchall()
    return [dict(row) for row in rows]


def list_clients() -> List[Dict[str, Any]]:
    with connect() as conn:
        rows = conn.execute("SELECT DISTINCT cliente, doi FROM documents ORDER BY cliente").fetchall()
    return [dict(row) for row in rows]


def create_review_item(document_key: str, kind: str, detail: str) -> None:
    with connect() as conn:
        conn.execute(
            "INSERT OR IGNORE INTO review_queue(document_key, kind, detail, created_at) VALUES (?, ?, ?, datetime('now'))",
            (document_key, kind, detail),
        )


def get_recent_imports(limit: int = 5) -> List[sqlite3.Row]:
    with connect() as conn:
        return conn.execute("SELECT * FROM imports ORDER BY imported_at DESC LIMIT ?", (limit,)).fetchall()
