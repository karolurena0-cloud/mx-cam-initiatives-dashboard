"""Dashboard de Proyectos IA -- FastAPI app.

Sirve un dashboard HTML (HTMX + Tailwind + Chart.js) que se alimenta de una
cache SQLite, la cual un loop en background mantiene sincronizada con el
Excel de SharePoint via Microsoft Graph. Los usuarios siguen editando el
Excel normalmente en SharePoint; el dashboard solo refleja esos cambios.
"""
import asyncio
import contextlib
import json
import logging

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from markupsafe import Markup

from app import db
from app.config import settings
from app.sync import sync_once

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("dashboard.main")

templates = Jinja2Templates(directory="app/templates")


def _tojson_safe(value) -> Markup:
    """json.dumps + escape '</' para poder incrustar el JSON dentro de <script>."""
    return Markup(json.dumps(value).replace("</", "<\\/"))


templates.env.filters["tojson_safe"] = _tojson_safe


async def _background_sync_loop() -> None:
    while True:
        await asyncio.sleep(settings.poll_interval_seconds)
        result = sync_once()
        logger.info("Background sync: %s", result)


@contextlib.asynccontextmanager
async def lifespan(app: FastAPI):
    db.init_db()
    logger.info("Sync inicial: %s", sync_once(force=True))
    task = asyncio.create_task(_background_sync_loop())
    yield
    task.cancel()
    with contextlib.suppress(asyncio.CancelledError):
        await task


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


def _dashboard_context() -> dict:
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
        "graph_auth_configured": settings.graph_auth_configured,
    }


@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request):
    return templates.TemplateResponse(request, "dashboard.html", _dashboard_context())


@app.get("/partials/content", response_class=HTMLResponse)
async def dashboard_content(request: Request, refresh: bool = False):
    """Partial que HTMX vuelve a pedir cada N segundos para refrescar la vista.

    Con ?refresh=1 fuerza un sync inmediato (botón 'Actualizar ahora').
    """
    if refresh:
        sync_once(force=True)
    return templates.TemplateResponse(request, "partials/content.html", _dashboard_context())