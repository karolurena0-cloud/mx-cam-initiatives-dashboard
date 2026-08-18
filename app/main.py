"""Dashboard de Proyectos IA -- FastAPI app.

Sirve un dashboard HTML (HTMX + Tailwind + Chart.js) que se alimenta de una
cache SQLite. Los datos se actualizan cuando alguien sube el .xlsx
actualizado desde la propia UI (boton "Subir Excel actualizado") -- no hay
integracion externa, cero configuracion de credenciales. Todos los usuarios
ven los cambios en su siguiente refresco automatico (cada 30s via HTMX).
"""
import contextlib
import json
import logging

from fastapi import FastAPI, Request, UploadFile, File
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from markupsafe import Markup

from app import db
from app.config import settings
from app.excel_import import import_workbook, ExcelImportError

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("dashboard.main")

templates = Jinja2Templates(directory="app/templates")


def _tojson_safe(value) -> Markup:
    """json.dumps + escape '</' para poder incrustar el JSON dentro de <script>."""
    return Markup(json.dumps(value).replace("</", "<\\/"))


templates.env.filters["tojson_safe"] = _tojson_safe


@contextlib.asynccontextmanager
async def lifespan(app: FastAPI):
    db.init_db()
    yield


app = FastAPI(title="Dashboard de Proyectos IA", lifespan=lifespan)
app.mount("/static", StaticFiles(directory="app/static"), name="static")


def _status_bucket(overall: str) -> str:
    overall = (overall or "").strip().upper()
    if overall.startswith("G"):
        return "green"
    if overall.startswith("A") or overall.startswith("Y"):
        return "amber"
    if overall.startswith("R"):
        return "red"
    return "unknown"


def _dashboard_context(upload_message: str | None = None, upload_ok: bool = True) -> dict:
    projects = [dict(r) for r in db.fetch_all("projects")]
    roadmap = [dict(r) for r in db.fetch_all("roadmap")]
    risks = [dict(r) for r in db.fetch_all("risk")]

    for p in projects:
        p["status_color"] = _status_bucket(p.get("overall", ""))

    summary = {
        "total_projects": len(projects),
        "green": sum(1 for p in projects if p["status_color"] == "green"),
        "amber": sum(1 for p in projects if p["status_color"] == "amber"),
        "red": sum(1 for p in projects if p["status_color"] == "red"),
        "open_risks": len(risks),
    }
    return {
        "projects": projects,
        "roadmap": roadmap,
        "risks": risks,
        "summary": summary,
        "last_synced": db.get_sync_meta("last_synced_at"),
        "last_upload_filename": db.get_sync_meta("last_upload_filename"),
        "source_file_modified": db.get_sync_meta("source_file_modified"),
        "upload_message": upload_message,
        "upload_ok": upload_ok,
    }


@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request):
    return templates.TemplateResponse(request, "dashboard.html", _dashboard_context())


@app.get("/partials/content", response_class=HTMLResponse)
async def dashboard_content(request: Request):
    """Partial que HTMX vuelve a pedir cada 30s para refrescar la vista con datos frescos."""
    return templates.TemplateResponse(request, "partials/content.html", _dashboard_context())


@app.post("/upload", response_class=HTMLResponse)
async def upload_excel(request: Request, file: UploadFile = File(...)):
    """Recibe el .xlsx subido desde la UI, lo parsea y actualiza la cache SQLite."""
    if not file.filename.lower().endswith(".xlsx"):
        ctx = _dashboard_context(upload_message="El archivo debe ser .xlsx", upload_ok=False)
        return templates.TemplateResponse(request, "partials/content.html", ctx)

    raw = await file.read()
    if len(raw) > settings.max_upload_bytes:
        ctx = _dashboard_context(
            upload_message=f"El archivo pesa mas de {settings.max_upload_bytes // (1024*1024)}MB",
            upload_ok=False,
        )
        return templates.TemplateResponse(request, "partials/content.html", ctx)

    try:
        result = import_workbook(raw, file.filename)
        if result["is_duplicate"]:
            msg = (
                f"Aviso: este archivo es identico (byte por byte) al ultimo que subiste. "
                f"Si esperabas ver cambios, probablemente descargaste una copia vieja de "
                f"SharePoint en vez de la mas reciente. Conteos: {result['counts']}"
            )
            ctx = _dashboard_context(upload_message=msg, upload_ok=False)
        else:
            msg = f"Archivo '{file.filename}' importado: {result['counts']}"
            ctx = _dashboard_context(upload_message=msg, upload_ok=True)
    except ExcelImportError as exc:
        logger.warning("Import fallido: %s", exc)
        ctx = _dashboard_context(upload_message=str(exc), upload_ok=False)

    return templates.TemplateResponse(request, "partials/content.html", ctx)
