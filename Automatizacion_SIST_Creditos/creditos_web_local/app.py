from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, Form, HTTPException, Query, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from database import backup_database, get_db_path, get_dashboard_summary, get_document, get_recent_imports, get_setting, init_db, list_documents, load_config, save_config
from services import import_signo_file, list_reviews
from outlook import MockOutlookDraftService, OutlookDraftService


app = FastAPI(title="Créditos Web Local", version="0.1.0")
app.mount("/static", StaticFiles(directory=str(Path(__file__).with_name("static"))), name="static")
templates = Jinja2Templates(directory=str(Path(__file__).with_name("templates")))


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


@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request) -> HTMLResponse:
    cfg = load_config()
    summary = get_dashboard_summary()
    imports = get_recent_imports(limit=5)
    docs = list_documents({"estado": "PENDIENTE"})[:10]
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
    document = get_document(document_key)
    if document is None:
        raise HTTPException(status_code=404, detail="Documento no encontrado")
    return templates.TemplateResponse("cliente.html", {"request": request, "document": document})


@app.get("/cobranza", response_class=HTMLResponse)
async def cobranza(request: Request) -> HTMLResponse:
    documents = list_documents({"estado": "PENDIENTE"})
    return templates.TemplateResponse("cobranza.html", {"request": request, "documents": documents})


@app.get("/revision", response_class=HTMLResponse)
async def revision(request: Request) -> HTMLResponse:
    items = list_reviews()
    return templates.TemplateResponse("revision.html", {"request": request, "items": items})


@app.get("/configuracion", response_class=HTMLResponse)
async def configuracion(request: Request) -> HTMLResponse:
    return templates.TemplateResponse("configuracion.html", {"request": request, "config": load_config()})


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
    return RedirectResponse(url="/configuracion", status_code=303)


@app.post("/importar")
async def importar(request: Request, file: UploadFile = None) -> JSONResponse:
    if not file:
        raise HTTPException(status_code=400, detail="No se recibió ningún archivo.")
    import_path = Path("/tmp") / file.filename
    with import_path.open("wb") as handle:
        handle.write(await file.read())
    result = import_signo_file(import_path)
    return JSONResponse(result)


@app.post("/preview-email")
async def preview_email(request: Request, cliente: str = Form(default=""), document_key: str = Form(default="")) -> HTMLResponse:
    document = get_document(document_key)
    if document is None:
        raise HTTPException(status_code=404, detail="Documento no encontrado")
    html = f"<html><body><p>Estimado/a {cliente},</p><p>Se adjunta el comprobante {document['document_key']} con saldo pendiente de S/ {document['saldo_observado']}.</p></body></html>"
    return HTMLResponse(content=html)


@app.post("/draft")
async def create_draft(document_key: str = Form(default=""), email: str = Form(default=""), subject: str = Form(default="Recordatorio de pago")) -> JSONResponse:
    if not document_key:
        raise HTTPException(status_code=400, detail="Documento requerido.")
    document = get_document(document_key)
    if document is None:
        raise HTTPException(status_code=404, detail="Documento no encontrado")
    cfg = load_config()
    service = OutlookDraftService(enabled=bool(cfg.get("outlook_enabled"))) if cfg.get("outlook_enabled") else MockOutlookDraftService(available=True)
    preview = f"<html><body><p>Asunto: {subject}</p><p>Cliente: {document['cliente']}</p><p>Saldo: S/ {document['saldo_observado']}</p></body></html>"
    result = service.create_draft(subject, preview, email or document['cliente'])
    if not result.success:
        return JSONResponse({"success": False, "error": result.error}, status_code=400)
    return JSONResponse({"success": True, "message_id": result.message_id, "preview": preview})


@app.post("/backup")
async def backup_db() -> JSONResponse:
    target = backup_database()
    return JSONResponse({"path": str(target)})


@app.get("/health")
async def health() -> JSONResponse:
    return JSONResponse({"status": "ok", "db": str(get_db_path())})


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app:app", host="127.0.0.1", port=8000)
