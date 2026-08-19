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

from fastapi import FastAPI, Request, UploadFile, File, Form
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


app = FastAPI(title="MX/CAM Initiatives App Portfolio Dashboard", lifespan=lifespan)
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


def _dashboard_context(
    upload_message: str | None = None,
    upload_ok: bool = True,
    status_filter: str = "all",
    engineer_filter: str = "all",
    app_filter: str = "all",
) -> dict:
    all_projects = [dict(r) for r in db.fetch_all("projects")]
    for p in all_projects:
        p["status_color"] = _status_bucket(p.get("overall", ""))

    # Listas para los dropdowns -- siempre a partir de TODOS los proyectos
    # (no de los ya filtrados), para no hacer desaparecer opciones del
    # filtro cuando otro filtro ya esta activo.
    engineers = sorted({p["engineering"] for p in all_projects if p.get("engineering")})
    apps = sorted({p["app"] for p in all_projects if p.get("app")})

    # 'app_filter' es una SELECCION (para resaltar), no un filtro que oculte
    # filas -- asi el usuario siempre puede hacer clic en otra app sin tener
    # que limpiar nada primero. status/engineer si ocultan filas de verdad.
    projects = all_projects
    if status_filter != "all":
        projects = [p for p in projects if p["status_color"] == status_filter]
    if engineer_filter != "all":
        projects = [p for p in projects if p.get("engineering") == engineer_filter]

    # Roadmap y riesgos tambien respetan status/engineer, pero muestran
    # TODAS las apps que queden dentro de ese alcance (la app seleccionada
    # se resalta en el template/JS, no se oculta el resto).
    visible_apps = {p["app"] for p in projects}
    roadmap = [dict(r) for r in db.fetch_all("roadmap") if dict(r).get("app") in visible_apps]
    risks = [dict(r) for r in db.fetch_all("risk") if dict(r).get("app") in visible_apps]

    # Detalle de la app seleccionada -- se muestra en la seccion Estado/
    # Dependencias (texto largo que no cabe bien en la tabla). Se busca en
    # TODOS los proyectos, no solo en los ya filtrados por status/engineer.
    selected_project = None
    if app_filter != "all":
        selected_project = next((p for p in all_projects if p["app"] == app_filter), None)

    summary = {
        "total_projects": len(projects),
        "green": sum(1 for p in projects if p["status_color"] == "green"),
        "amber": sum(1 for p in projects if p["status_color"] == "amber"),
        "red": sum(1 for p in projects if p["status_color"] == "red"),
        "open_risks": len(risks),
    }

    # Datos para las graficas de analisis (siempre sobre el set ya filtrado).
    status_chart = {
        "labels": ["Verde", "Amarillo", "Rojo", "Desconocido"],
        "data": [
            summary["green"],
            summary["amber"],
            summary["red"],
            sum(1 for p in projects if p["status_color"] == "unknown"),
        ],
    }
    engineer_counts: dict[str, int] = {}
    for p in projects:
        name = p.get("engineering") or "Sin asignar"
        engineer_counts[name] = engineer_counts.get(name, 0) + 1
    engineer_chart = {
        "labels": list(engineer_counts.keys()),
        "data": list(engineer_counts.values()),
    }

    return {
        "projects": projects,
        "roadmap": roadmap,
        "risks": risks,
        "summary": summary,
        "engineers": engineers,
        "apps": apps,
        "status_filter": status_filter,
        "engineer_filter": engineer_filter,
        "app_filter": app_filter,
        "selected_project": selected_project,
        "status_chart": status_chart,
        "engineer_chart": engineer_chart,
        "last_synced": db.get_sync_meta("last_synced_at"),
        "last_upload_filename": db.get_sync_meta("last_upload_filename"),
        "source_file_modified": db.get_sync_meta("source_file_modified"),
        "upload_message": upload_message,
        "upload_ok": upload_ok,
    }


@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request, status: str = "all", engineer: str = "all", app: str = "all"):
    ctx = _dashboard_context(status_filter=status, engineer_filter=engineer, app_filter=app)
    return templates.TemplateResponse(request, "dashboard.html", ctx)


@app.get("/partials/content", response_class=HTMLResponse)
async def dashboard_content(request: Request, status: str = "all", engineer: str = "all", app: str = "all"):
    """Partial que HTMX vuelve a pedir cada 30s (o al cambiar/hacer clic en un filtro) para refrescar la vista."""
    ctx = _dashboard_context(status_filter=status, engineer_filter=engineer, app_filter=app)
    return templates.TemplateResponse(request, "partials/content.html", ctx)


@app.post("/upload", response_class=HTMLResponse)
async def upload_excel(
    request: Request,
    file: UploadFile = File(...),
    status: str = Form("all"),
    engineer: str = Form("all"),
    app: str = Form("all"),
):
    """Recibe el .xlsx subido desde la UI, lo parsea y actualiza la cache SQLite.

    Conserva los filtros activos (vienen incluidos via hx-include en el form)
    para que subir un archivo no le resetee la vista al usuario.
    """
    if not file.filename.lower().endswith(".xlsx"):
        ctx = _dashboard_context(
            upload_message="El archivo debe ser .xlsx", upload_ok=False,
            status_filter=status, engineer_filter=engineer, app_filter=app,
        )
        return templates.TemplateResponse(request, "partials/content.html", ctx)

    raw = await file.read()
    if len(raw) > settings.max_upload_bytes:
        ctx = _dashboard_context(
            upload_message=f"El archivo pesa mas de {settings.max_upload_bytes // (1024*1024)}MB",
            upload_ok=False, status_filter=status, engineer_filter=engineer, app_filter=app,
        )
        return templates.TemplateResponse(request, "partials/content.html", ctx)

    try:
        result = import_workbook(raw, file.filename)
        if result["is_duplicate"]:
            msg = (
                "Aviso: este archivo es identico (byte por byte) al ultimo que subiste. "
                "Si esperabas ver cambios, probablemente descargaste una copia vieja de "
                "SharePoint en vez de la mas reciente."
            )
            ctx = _dashboard_context(
                upload_message=msg, upload_ok=False,
                status_filter=status, engineer_filter=engineer, app_filter=app,
            )
        else:
            ctx = _dashboard_context(
                upload_message="El dashboard se actualizo con exito.", upload_ok=True,
                status_filter=status, engineer_filter=engineer, app_filter=app,
            )
    except ExcelImportError as exc:
        logger.warning("Import fallido: %s", exc)
        ctx = _dashboard_context(
            upload_message=str(exc), upload_ok=False,
            status_filter=status, engineer_filter=engineer, app_filter=app,
        )

    return templates.TemplateResponse(request, "partials/content.html", ctx)
