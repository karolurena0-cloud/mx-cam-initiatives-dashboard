"""Configuracion centralizada del dashboard. Un solo lugar para leer el .env."""
import os
from dotenv import load_dotenv

load_dotenv()


class Settings:
    # Storage
    db_path: str = os.getenv("DASHBOARD_DB_PATH", "dashboard.db")
    # Limite de tamano para el archivo subido (bytes). 10 MB de sobra para este workbook.
    max_upload_bytes: int = int(os.getenv("MAX_UPLOAD_BYTES", str(10 * 1024 * 1024)))


settings = Settings()

# Nombres de las 3 hojas que alimentan el dashboard.
SHEETS = {
    "projects": "Projects",
    "roadmap": "Roadmap",
    "risk": "Risk",
}
