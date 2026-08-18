"""Capa de datos SQLite: cache local de las 3 hojas del Excel.

Cada sync borra e inserta de nuevo (los datos son pequenos, ~65 filas
totales) en vez de hacer diffing fila por fila -- mas simple, YAGNI.
"""
import sqlite3
from contextlib import contextmanager
from datetime import datetime, timezone

from app.config import settings

SCHEMA = """
CREATE TABLE IF NOT EXISTS projects (
    app TEXT,
    overall TEXT,
    s4_dependency TEXT,
    product TEXT,
    engineering TEXT,
    project_overview TEXT,
    status TEXT,
    dependencies TEXT,
    platform TEXT
);

CREATE TABLE IF NOT EXISTS roadmap (
    app TEXT,
    task TEXT,
    start_date TEXT,
    end_date TEXT,
    pct REAL
);

CREATE TABLE IF NOT EXISTS risk (
    app TEXT,
    risk TEXT,
    mitigation TEXT,
    resolution_date TEXT
);

CREATE TABLE IF NOT EXISTS sync_meta (
    key TEXT PRIMARY KEY,
    value TEXT
);
"""

# Columnas del Excel -> columnas de la tabla, en orden. Ver README para el
# esquema original de cada hoja (viene del sub-agente msgraph).
TABLE_COLUMNS = {
    "projects": ["app", "overall", "s4_dependency", "product", "engineering",
                 "project_overview", "status", "dependencies", "platform"],
    "roadmap": ["app", "task", "start_date", "end_date", "pct"],
    "risk": ["app", "risk", "mitigation", "resolution_date"],
}


@contextmanager
def get_conn():
    conn = sqlite3.connect(settings.db_path)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db() -> None:
    with get_conn() as conn:
        conn.executescript(SCHEMA)


def replace_table(table: str, rows: list[dict]) -> None:
    """Vacia y repuebla una tabla con las filas dadas (lista de dicts)."""
    cols = TABLE_COLUMNS[table]
    placeholders = ", ".join(["?"] * len(cols))
    col_list = ", ".join(cols)
    with get_conn() as conn:
        conn.execute(f"DELETE FROM {table}")
        conn.executemany(
            f"INSERT INTO {table} ({col_list}) VALUES ({placeholders})",
            [[row.get(c, "") for c in cols] for row in rows],
        )


def set_sync_meta(key: str, value: str) -> None:
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO sync_meta (key, value) VALUES (?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (key, value),
        )


def get_sync_meta(key: str) -> str | None:
    with get_conn() as conn:
        row = conn.execute("SELECT value FROM sync_meta WHERE key = ?", (key,)).fetchone()
        return row["value"] if row else None


def mark_synced_now() -> None:
    set_sync_meta("last_synced_at", datetime.now(timezone.utc).isoformat())


def fetch_all(table: str) -> list[sqlite3.Row]:
    with get_conn() as conn:
        return conn.execute(f"SELECT * FROM {table}").fetchall()
