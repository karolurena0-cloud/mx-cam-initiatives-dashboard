"""Job de sincronizacion: Excel (via Graph) -> SQLite.

Se corre una vez al arrancar y luego en un loop cada POLL_INTERVAL_SECONDS
(ver app/main.py, lifespan). Solo relee el workbook completo si
lastModifiedDateTime cambio desde el ultimo sync exitoso -- evita pegarle
a Graph sin necesidad cuando nadie ha tocado el Excel.
"""
import logging

from app.config import settings, SHEETS
from app.graph_client import graph_client, GraphAuthError
from app import db

logger = logging.getLogger("dashboard.sync")

# Los encabezados normalizados del Excel no siempre coinciden con los nombres
# de columna que usamos en SQLite (ver app/db.py TABLE_COLUMNS). Este mapa
# traduce solo donde hace falta; cualquier encabezado no listado se usa tal cual.
HEADER_ALIASES = {
    "roadmap": {
        "roadmap_task": "task",
        "roadmap_start": "start_date",
        "roadmap_end": "end_date",
        "roadmap_%": "pct",
    },
}


def _rows_from_values(values: list[list], table: str) -> list[dict]:
    """values[0] son los encabezados; el resto son filas de datos."""
    if not values:
        return []
    aliases = HEADER_ALIASES.get(table, {})
    headers = [str(h).strip().lower().replace(" ", "_") for h in values[0]]
    headers = [aliases.get(h, h) for h in headers]
    rows = []
    for raw_row in values[1:]:
        # Si Excel deja columnas vacias al final, values puede venir corto -- normalizamos.
        padded = raw_row + [""] * (len(headers) - len(raw_row))
        row = dict(zip(headers, padded))
        # Saltar filas totalmente vacias (blank rows dentro del usedRange).
        if any(str(v).strip() for v in row.values()):
            rows.append(row)
    return rows


def sync_once(force: bool = False) -> dict:
    """Ejecuta un ciclo de sync. Retorna un resumen para logging/manual trigger."""
    if not settings.graph_auth_configured:
        return {"status": "skipped", "reason": "graph_auth_not_configured"}

    try:
        last_modified = graph_client.get_item_last_modified()
    except GraphAuthError as exc:
        logger.error("Auth de Graph fallo: %s", exc)
        return {"status": "error", "reason": str(exc)}
    except Exception as exc:  # httpx errors, etc.
        logger.error("No se pudo consultar metadata del workbook: %s", exc)
        return {"status": "error", "reason": str(exc)}

    previous_modified = db.get_sync_meta("source_last_modified")
    if not force and last_modified == previous_modified:
        return {"status": "unchanged", "last_modified": last_modified}

    for table, sheet_name in SHEETS.items():
        try:
            values = graph_client.get_worksheet_values(sheet_name)
        except Exception as exc:
            logger.error("Error leyendo hoja %s: %s", sheet_name, exc)
            return {"status": "error", "reason": f"sheet '{sheet_name}': {exc}"}
        rows = _rows_from_values(values, table)
        db.replace_table(table, rows)

    db.set_sync_meta("source_last_modified", last_modified or "")
    db.mark_synced_now()
    logger.info("Sync OK -- workbook modificado %s", last_modified)
    return {"status": "synced", "last_modified": last_modified}
