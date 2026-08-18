"""Configuracion centralizada del dashboard. Un solo lugar para leer el .env."""
import os
from dotenv import load_dotenv

load_dotenv()


class Settings:
    # Azure AD / Graph
    tenant_id: str = os.getenv("AZURE_TENANT_ID", "")
    client_id: str = os.getenv("AZURE_CLIENT_ID", "")
    client_secret: str = os.getenv("AZURE_CLIENT_SECRET", "")
    graph_scope: str = os.getenv("GRAPH_SCOPE", "https://graph.microsoft.com/.default")

    # SharePoint / workbook target
    site_id: str = os.getenv("SHAREPOINT_SITE_ID", "")
    drive_id: str = os.getenv("SHAREPOINT_DRIVE_ID", "")
    item_id: str = os.getenv("WORKBOOK_ITEM_ID", "")

    # Sync behaviour
    poll_interval_seconds: int = int(os.getenv("POLL_INTERVAL_SECONDS", "60"))

    # Storage
    db_path: str = os.getenv("DASHBOARD_DB_PATH", "dashboard.db")

    @property
    def graph_auth_configured(self) -> bool:
        return bool(self.tenant_id and self.client_id and self.client_secret)


settings = Settings()

# Nombres de las 3 hojas que alimentan el dashboard (ver README para el esquema).
SHEETS = {
    "projects": "Projects",
    "roadmap": "Roadmap",
    "risk": "Risk",
}
