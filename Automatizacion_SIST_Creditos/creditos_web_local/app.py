from __future__ import annotations

import json
import re
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, Form, HTTPException, Query, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from database import (
    backup_database, draft_already_exists, get_dashboard_summary, get_db_path,
    get_document, get_recent_imports, get_setting, init_db, list_documents,
    load_config, mark_draft_created, mark_draft_failed, reserve_draft_slot, save_config,
)
from services import import_signo_file, list_reviews
from outlook import OutlookDraftService


app = FastAPI(title="Créditos Web Local", version="0.1.0")
app.mount("/static", StaticFiles(directory=str(Path(__file__).with_name("static"))), name="static")
templates = Jinja2Templates(directory=str(Path(__file__).with_name("templates")))
EMAIL_RE = re.compile(
    r"^(?=.{1,254}$)(?=.{1,64}@)[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@"
    r"[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+$"
)


def is_valid_email(value: str) -> bool:
    return bool(value and EMAIL_RE.fullmatch(value.strip()))


def config_defaults() -> Dict[str, Any]:
    return {
        "signo_folder": "",
        "local_db": str(get_db_path()),
        "currency_default": "PEN",
        "check_frequency_minutes": 15,
        "days_advance": 7,
        "outlook_account": "",
        "outlook_drafts_store": "",
        "outlook_enabled": False,
        "show_recent_docs": 20,
    }


@app.on_event("startup")
def startup_event() -> None:
    init_db()
    current = load_config()
    defaults = config_defaults()
    merged = {**defaults, **current}
    save_config(merged)
    from scheduler import start_scheduler
    start_scheduler(merged)


@app.on_event("shutdown")
def shutdown_event() -> None:
    from scheduler import stop_scheduler
    stop_scheduler()


@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request) -> HTMLResponse:
    cfg = load_config()
    summary = get_dashboard_summary(days_advance=int(cfg.get("days_advance", 7)))
    imports = get_recent_imports(limit=5)
    docs = list_documents({"estado": ["PENDIENTE", "SALDO_PARCIAL"]})[:10]
    return templates.TemplateResponse(
        "dashboard.html",
        {"request": request, "summary": summary, "imports": imports, "documents": docs, "config": cfg},
    )


@app.get("/cartera", response_class=HTMLResponse)
async def cartera(request: Request, cliente: str = Query(default=""), documento: str = Query(default=""), responsable: str = Query(default=""), estado: str = Query(default="")) -> HTMLResponse:
    filters = {k: v for k, v in {"cliente": cliente, "documento": documento, "responsable": responsable, "estado": estado}.items() if v}
    docs = list_documents(filters)
    return templates.TemplateResponse("cartera.html", {"request": request, "documents": docs, "filters": filters})


@app.get("/cliente/{document_key}", response_class=HTMLResponse)
async def client_detail(request: Request, document_key: str) -> HTMLResponse:
    from database import connect as _connect
    document = get_document(document_key)
    if document is None:
        raise HTTPException(status_code=404, detail="Documento no encontrado")
    with _connect() as conn:
        contacts = conn.execute(
            "SELECT * FROM contacts WHERE document_key = ? ORDER BY es_principal DESC, id",
            (document_key,),
        ).fetchall()
    return templates.TemplateResponse(
        "cliente.html",
        {
            "request": request,
            "document": document,
            "contacts": contacts,
            "message": request.query_params.get("message", ""),
            "error": request.query_params.get("error", ""),
        },
    )


@app.get("/cobranza", response_class=HTMLResponse)
async def cobranza(request: Request) -> HTMLResponse:
    documents = list_documents({"estado": ["PENDIENTE", "SALDO_PARCIAL"]})
    # Exclude zero-saldo and documents with unresolved inconsistencies from reminder selection
    cobrables = [d for d in documents if float(d["saldo_observado"] or 0) > 0 and not d["revision_required"]]
    return templates.TemplateResponse(
        "cobranza.html",
        {
            "request": request,
            "documents": cobrables,
            "message": request.query_params.get("message", ""),
            "error": request.query_params.get("error", ""),
        },
    )


@app.get("/revision", response_class=HTMLResponse)
async def revision(request: Request) -> HTMLResponse:
    items = list_reviews()
    return templates.TemplateResponse("revision.html", {"request": request, "items": items})


@app.get("/configuracion", response_class=HTMLResponse)
async def configuracion(request: Request) -> HTMLResponse:
    from scheduler import get_scheduler_status
    return templates.TemplateResponse(
        "configuracion.html",
        {
            "request": request,
            "config": load_config(),
            "scheduler": get_scheduler_status(),
            "message": request.query_params.get("message", ""),
            "error": request.query_params.get("error", ""),
        },
    )


@app.post("/configuracion")
async def save_configuration(
    request: Request,
    signo_folder: str = Form(default=""),
    currency_default: str = Form(default="PEN"),
    check_frequency_minutes: int = Form(default=15),
    days_advance: int = Form(default=7),
    outlook_account: str = Form(default=""),
    outlook_drafts_store: str = Form(default=""),
    outlook_enabled: bool = Form(default=False),
) -> RedirectResponse:
    config = load_config()
    config.update(
        {
            "signo_folder": signo_folder,
            "currency_default": currency_default,
            "check_frequency_minutes": check_frequency_minutes,
            "days_advance": days_advance,
            "outlook_account": outlook_account,
            "outlook_drafts_store": outlook_drafts_store,
            "outlook_enabled": outlook_enabled,
        }
    )
    save_config(config)
    from scheduler import start_scheduler
    start_scheduler(config)
    return RedirectResponse(url="/configuracion?message=Configuración actualizada", status_code=303)


@app.post("/importar")
async def importar(request: Request, file: UploadFile = None) -> JSONResponse:
    if not file:
        raise HTTPException(status_code=400, detail="No se recibió ningún archivo.")
    # Use a platform-safe temp file instead of /tmp (which fails on Windows)
    suffix = Path(file.filename or "upload").suffix or ".xlsx"
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(await file.read())
        import_path = Path(tmp.name)
    try:
        result = import_signo_file(import_path)
    finally:
        try:
            import_path.unlink(missing_ok=True)
        except OSError:
            pass
    return JSONResponse(result)


@app.post("/preview-email")
async def preview_email(request: Request, cliente: str = Form(default=""), document_key: str = Form(default="")) -> HTMLResponse:
    document = get_document(document_key)
    if document is None:
        raise HTTPException(status_code=404, detail="Documento no encontrado")
    html = f"<html><body><p>Estimado/a {cliente},</p><p>Se adjunta el comprobante {document['document_key']} con saldo pendiente de S/ {document['saldo_observado']}.</p></body></html>"
    return HTMLResponse(content=html)


@app.post("/draft")
async def create_draft(
    document_key: str = Form(default=""),
    email: str = Form(default=""),
    subject: str = Form(default="Recordatorio de pago"),
) -> JSONResponse:
    if not document_key:
        raise HTTPException(status_code=400, detail="Documento requerido.")
    document = get_document(document_key)
    if document is None:
        raise HTTPException(status_code=404, detail="Documento no encontrado")
    if document["estado"] not in {"PENDIENTE", "SALDO_PARCIAL"}:
        raise HTTPException(status_code=400, detail="Solo se permiten comprobantes cobrables (PENDIENTE o SALDO_PARCIAL).")
    if document["revision_required"]:
        raise HTTPException(status_code=400, detail="El comprobante tiene inconsistencias pendientes de revisión.")
    try:
        saldo = float(document["saldo_observado"] or 0)
    except (TypeError, ValueError):
        raise HTTPException(status_code=400, detail="Saldo inválido para generar borrador.")
    if saldo <= 0:
        raise HTTPException(status_code=400, detail="El comprobante no tiene saldo pendiente.")

    cfg = load_config()
    preview = (
        f"<html><body>"
        f"<p>Asunto: {subject}</p>"
        f"<p>Cliente: {document['cliente']}</p>"
        f"<p>Saldo: S/ {document['saldo_observado']}</p>"
        f"</body></html>"
    )

    if not is_valid_email(email):
        raise HTTPException(
            status_code=400,
            detail="Se requiere un correo electrónico válido como destinatario.",
        )

    if not cfg.get("outlook_enabled"):
        return JSONResponse({
            "success": False,
            "preview": preview,
            "message": "Integración Outlook desactivada. El borrador no fue creado.",
        })

    if draft_already_exists(document_key, subject):
        return JSONResponse({
            "success": True,
            "duplicate": True,
            "message": "Ya existe un borrador para este documento y asunto.",
            "preview": preview,
        })

    reserved, reservation_status = reserve_draft_slot(
        document_key=document_key,
        cliente=document["cliente"],
        destinatario=email.strip(),
        subject=subject,
        html_preview=preview,
    )
    if not reserved:
        if reservation_status == "PROCESSING":
            return JSONResponse(
                {"success": False, "error": "Ya existe una solicitud en proceso para este documento y asunto."},
                status_code=409,
            )
        if reservation_status == "CREATED":
            return JSONResponse(
                {"success": True, "duplicate": True, "message": "Ya existe un borrador para este documento y asunto.", "preview": preview},
            )
        return JSONResponse(
            {"success": False, "error": "No se pudo reservar el borrador. Intente nuevamente."},
            status_code=409,
        )

    dedupe_key = f"{document_key}|{email.strip()}|{subject}"
    service = OutlookDraftService(enabled=True)
    try:
        result = service.create_draft(
            subject,
            preview,
            email.strip(),
            account=str(cfg.get("outlook_account") or ""),
            dedupe_key=dedupe_key,
        )
        if not result.success:
            mark_draft_failed(document_key, subject, result.error or "Error desconocido de Outlook")
            return JSONResponse({"success": False, "error": result.error}, status_code=400)

        mark_draft_created(
            document_key=document_key,
            subject=subject,
            outlook_message_id=result.message_id or "",
            html_preview=preview,
        )
    finally:
        service.release()
    return JSONResponse({"success": True, "message_id": result.message_id, "preview": preview})


@app.post("/backup")
async def backup_db() -> JSONResponse:
    target = backup_database()
    return JSONResponse({"path": str(target)})


@app.get("/contacts/{document_key_path}", response_class=JSONResponse)
async def list_contacts(document_key_path: str) -> JSONResponse:
    from database import connect as _connect
    with _connect() as conn:
        rows = conn.execute(
            "SELECT * FROM contacts WHERE document_key = ? ORDER BY es_principal DESC, id",
            (document_key_path,),
        ).fetchall()
    return JSONResponse([dict(r) for r in rows])


@app.post("/contacts")
async def upsert_contact(
    document_key: str = Form(...),
    nombre: str = Form(default=""),
    correo: str = Form(default=""),
    telefono: str = Form(default=""),
    rol: str = Form(default=""),
    es_principal: bool = Form(default=False),
) -> JSONResponse:
    if not is_valid_email(correo):
        raise HTTPException(status_code=400, detail="Correo no válido.")
    from database import connect as _connect
    with _connect() as conn:
        conn.execute(
            """INSERT INTO contacts(document_key, nombre, correo, telefono, rol, es_principal)
               VALUES (?, ?, ?, ?, ?, ?)
               ON CONFLICT(document_key, correo) DO UPDATE SET
                 nombre = excluded.nombre,
                 telefono = excluded.telefono,
                 rol = excluded.rol,
                 es_principal = excluded.es_principal""",
            (document_key, nombre, correo, telefono, rol, int(es_principal)),
        )
    return JSONResponse({"success": True})


@app.post("/cliente/{document_key}/contactos")
async def upsert_contact_form(
    document_key: str,
    nombre: str = Form(default=""),
    correo: str = Form(default=""),
    telefono: str = Form(default=""),
    rol: str = Form(default=""),
    es_principal: bool = Form(default=False),
    contact_id: int = Form(default=0),
) -> RedirectResponse:
    from database import connect as _connect
    if not is_valid_email(correo):
        return RedirectResponse(
            url=f"/cliente/{document_key}?error=Correo%20no%20válido",
            status_code=303,
        )
    with _connect() as conn:
        if contact_id:
            conn.execute(
                """UPDATE contacts
                   SET nombre = ?, correo = ?, telefono = ?, rol = ?, es_principal = ?
                   WHERE id = ? AND document_key = ?""",
                (nombre, correo, telefono, rol, int(es_principal), contact_id, document_key),
            )
        else:
            conn.execute(
                """INSERT INTO contacts(document_key, nombre, correo, telefono, rol, es_principal)
                   VALUES (?, ?, ?, ?, ?, ?)
                   ON CONFLICT(document_key, correo) DO UPDATE SET
                     nombre = excluded.nombre,
                     telefono = excluded.telefono,
                     rol = excluded.rol,
                     es_principal = excluded.es_principal""",
                (document_key, nombre, correo, telefono, rol, int(es_principal)),
            )
        if es_principal:
            conn.execute(
                "UPDATE contacts SET es_principal = 0 WHERE document_key = ? AND correo <> ?",
                (document_key, correo),
            )
    return RedirectResponse(url=f"/cliente/{document_key}?message=Contacto%20guardado", status_code=303)


@app.get("/health")
async def health() -> JSONResponse:
    from scheduler import get_scheduler_status
    sched = get_scheduler_status()
    return JSONResponse({"status": "ok", "db": str(get_db_path()), "scheduler": sched})


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app:app", host="127.0.0.1", port=8000)
