"""Cliente delgado para Microsoft Graph: solo lo que necesitamos (auth + leer usedRange).

Sigue YAGNI a proposito: no envolvemos todo el SDK de Graph, solo las dos
llamadas que este dashboard realmente usa.
"""
import time
import httpx
import msal

from app.config import settings

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class GraphAuthError(RuntimeError):
    """Se lanza cuando falta configuracion de Azure AD o la autenticacion falla."""


class GraphClient:
    def __init__(self) -> None:
        self._token: str | None = None
        self._token_expires_at: float = 0.0
        self._msal_app: msal.ConfidentialClientApplication | None = None

    def _get_msal_app(self) -> msal.ConfidentialClientApplication:
        if not settings.graph_auth_configured:
            raise GraphAuthError(
                "Faltan AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET en el .env. "
                "Revisa el README para el proceso de app registration."
            )
        if self._msal_app is None:
            authority = f"https://login.microsoftonline.com/{settings.tenant_id}"
            self._msal_app = msal.ConfidentialClientApplication(
                client_id=settings.client_id,
                client_credential=settings.client_secret,
                authority=authority,
            )
        return self._msal_app

    def _get_token(self) -> str:
        # Cache simple en memoria: no pedimos token nuevo si el actual sigue vivo.
        if self._token and time.time() < self._token_expires_at - 60:
            return self._token

        app = self._get_msal_app()
        result = app.acquire_token_for_client(scopes=[settings.graph_scope])
        if "access_token" not in result:
            raise GraphAuthError(
                f"No se pudo obtener token de Graph: {result.get('error_description', result)}"
            )
        self._token = result["access_token"]
        self._token_expires_at = time.time() + result.get("expires_in", 3600)
        return self._token

    def get_worksheet_values(self, worksheet_name: str) -> list[list]:
        """Lee usedRange (solo valores) de una hoja. Fila 0 = encabezados."""
        token = self._get_token()
        url = (
            f"{GRAPH_BASE}/sites/{settings.site_id}/drives/{settings.drive_id}"
            f"/items/{settings.item_id}/workbook/worksheets/{worksheet_name}"
            f"/usedRange(valuesOnly=true)"
        )
        headers = {"Authorization": f"Bearer {token}"}
        resp = httpx.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        return resp.json().get("values", [])

    def get_item_last_modified(self) -> str | None:
        """Revisa lastModifiedDateTime del archivo para evitar releer si no cambio."""
        token = self._get_token()
        url = f"{GRAPH_BASE}/sites/{settings.site_id}/drives/{settings.drive_id}/items/{settings.item_id}"
        headers = {"Authorization": f"Bearer {token}"}
        resp = httpx.get(url, headers=headers, timeout=30, params={"$select": "lastModifiedDateTime"})
        resp.raise_for_status()
        return resp.json().get("lastModifiedDateTime")


graph_client = GraphClient()
