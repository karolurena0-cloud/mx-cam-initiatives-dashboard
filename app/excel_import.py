"""Importador de Excel: archivo .xlsx subido a mano -> SQLite.

Reemplaza el enfoque anterior (Microsoft Graph + Azure AD app registration)
por algo mucho mas simple: cualquier persona con el archivo actualizado lo
sube desde el dashboard, se parsea aqui y todos los demas usuarios ven los
cambios en su siguiente refresco automatico (ver hx-trigger en el template).
"""
import io
import logging

import openpyxl

from app import db
from app.config import SHEETS

logger = logging.getLogger("dashboard.excel_import")


class ExcelImportError(ValueError):
    """Se lanza cuando el archivo subido no tiene la forma esperada."""


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
    headers = [str(h or "").strip().lower().replace(" ", "_") for h in values[0]]
    headers = [aliases.get(h, h) for h in headers]
    rows = []
    for raw_row in values[1:]:
        raw_row = list(raw_row)
        padded = raw_row + [""] * (len(headers) - len(raw_row))
        row = dict(zip(headers, padded))
        # Saltar filas totalmente vacias (comunes al final de un rango en Excel).
        if any(str(v).strip() for v in row.values() if v is not None):
            rows.append(row)
    return rows


def import_workbook(file_bytes: bytes, filename: str) -> dict:
    """Parsea el .xlsx subido y reemplaza el contenido de las 3 tablas en SQLite."""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    except Exception as exc:
        raise ExcelImportError(f"No se pudo abrir el archivo como Excel valido: {exc}") from exc

    missing_sheets = [name for name in SHEETS.values() if name not in wb.sheetnames]
    if missing_sheets:
        raise ExcelImportError(
            f"Al workbook le faltan las hojas: {', '.join(missing_sheets)}. "
            f"Hojas encontradas: {', '.join(wb.sheetnames)}"
        )

    imported_counts = {}
    for table, sheet_name in SHEETS.items():
        ws = wb[sheet_name]
        values = [list(row) for row in ws.iter_rows(values_only=True)]
        rows = _rows_from_values(values, table)
        db.replace_table(table, rows)
        imported_counts[table] = len(rows)

    db.set_sync_meta("last_upload_filename", filename)
    db.mark_synced_now()
    logger.info("Import OK desde '%s': %s", filename, imported_counts)
    return {"status": "imported", "filename": filename, "counts": imported_counts}
