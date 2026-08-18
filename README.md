# Dashboard de Proyectos IA

Dashboard que se alimenta en vivo del Excel `BU-promxcam.xlsx` alojado en
SharePoint (sitio `EBSLAMX-Portfolio`, biblioteca `Technology`, carpeta
`Dashboard Proyectos IA`). Los usuarios siguen editando el Excel normalmente
desde SharePoint/Teams -- este servicio solo lee esos cambios y los muestra
a todos en un dashboard web que se refresca solo.

## Como funciona

```
Excel en SharePoint  --(Microsoft Graph, polling cada 60s)-->  SQLite (cache)  --(FastAPI + HTMX)-->  Dashboard web
```

- Un job en background (`app/sync.py`) consulta Graph cada `POLL_INTERVAL_SECONDS`.
  Antes de releer todo el workbook, revisa `lastModifiedDateTime` del archivo
  para no pegarle a Graph si nadie ha editado nada.
- Los datos se guardan en una base SQLite local (`dashboard.db`, se crea sola).
- El dashboard (`/`) usa HTMX para pedir el contenido actualizado cada 30s,
  sin recargar la pagina completa.
- Los usuarios **no editan nada dentro de la app** -- siguen usando el Excel
  de SharePoint como siempre. La app es de solo lectura/visualizacion.

Hojas leidas del workbook:

| Hoja | Columnas | Uso en el dashboard |
|---|---|---|
| `Projects` | App, Overall, S4 Dependency, Product, Engineering, Project Overview, Status, Dependencies, platform | Tabla principal + tarjetas resumen (verde/amarillo/rojo) |
| `Roadmap` | App, Roadmap task, Roadmap Start, Roadmap End, Roadmap % | Grafica de barras de % de avance |
| `Risk` | App, Risk, Mitigation, Resolution date | Tabla de riesgos |

## Setup

### 1. Entorno Python (uv)

```bash
uv venv --python 3.13.5
source .venv/bin/activate
uv pip install --index-url https://pypi.ci.artifacts.walmart.com/artifactory/api/pypi/external-pypi/simple \
    --allow-insecure-host pypi.ci.artifacts.walmart.com -r requirements.txt
```

### 2. Autenticacion con Microsoft Graph (la parte importante)

Esta app corre desatendida (nadie hace login interactivo), asi que usa
**OAuth2 client credentials** (app-only) contra Azure AD:

1. Pide a tu equipo de **IAM / Cloud Identity** un **app registration** en
   Azure AD con permiso de **aplicacion** (no delegado):
   - Preferido: **`Sites.Selected`** (least-privilege -- solo le da acceso al
     sitio que tu autorices, no a todo el tenant).
   - Alternativa si `Sites.Selected` no es viable: `Sites.Read.All` + `Files.Read.All`.
   - Requiere **admin consent**.
2. Si usas `Sites.Selected`, pide a un admin del sitio (o hazlo tu con permisos
   adecuados) que corra un `POST /sites/{site-id}/permissions` para darle
   acceso a la app **especificamente** al sitio `EBSLAMX-Portfolio`.
3. Copia `.env.example` a `.env` y llena:
   ```
   AZURE_TENANT_ID=...
   AZURE_CLIENT_ID=...
   AZURE_CLIENT_SECRET=...
   ```
   Los valores de `SHAREPOINT_SITE_ID`, `SHAREPOINT_DRIVE_ID` y `WORKBOOK_ITEM_ID`
   ya vienen prellenados en `.env.example` (resueltos para este workbook especifico).

   **Nunca subas `.env` a git** (a en `.gitignore`). En un ambiente
   desplegado usa un secrets manager en vez de un `.env` plano.

### 3. (Recomendado) Convertir las hojas en Tablas de Excel

El workbook actual no tiene Tablas (`Ctrl+T`) definidas -- el sync lee por
`usedRange`, que funciona pero es menos robusto ante insercion/borrado manual
de filas. Si el dueno del Excel puede convertir cada hoja (`Projects`,
`Roadmap`, `Risk`) en una Tabla con los mismos encabezados, el sync se vuelve
mas resiliente (aunque no es obligatorio para que esto funcione hoy).

### 4. Correr la app

```bash
source .venv/bin/activate
uvicorn app.main:app --reload --port 8000
```

Abre `http://localhost:8000` -- veras el dashboard. Si `.env` no tiene las
credenciales de Graph configuradas, la app arranca igual pero muestra un
aviso de "configuracion pendiente" y tablas vacias (no truena).

## Estructura del proyecto

```
app/
  config.py       # Lee .env, un solo lugar de configuracion
  graph_client.py # Auth MSAL + llamadas a Graph (token cacheado en memoria)
  db.py           # Esquema SQLite + helpers de lectura/escritura
  sync.py         # el -> SQLite, con chequeo de lastModifiedDateTime
  main.py         # FastAPI: rutas + loop de sync en background (lifespan)
  templates/      # Jinja2 + HTMX + Tailwind (CDN) + Chart.js (CDN)
  static/         # Assets estaticos (vacio por ahora)
requirements.txt
.env.example
```

## Notas de seguridad

- El `.env` con secretos de Azure **nunca** se commitea (ver `.gitignore`).
- `dashboard.db` (la cache SQLite) tampoco se commitea -- se regenera sola
  en cada arranque via sync.
- Antes de compartir este repo, revisa que no haya datos sensibles
  hardcodeados en commits previos.
