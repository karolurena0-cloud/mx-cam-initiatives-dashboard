# Dashboard de Proyectos IA

Dashboard que muestra el estado de los proyectos IA a partir del Excel
`BU-promxcam.xlsx` (originalmente en SharePoint, sitio `EBSLAMX-Portfolio`).

## Como funciona

```
Alguien descarga/edita el Excel  -->  lo sube desde el dashboard (boton)  -->  SQLite (cache)  -->  todos ven los cambios (HTMX refresca solo cada 30s)
```

- **Sin credenciales, sin Azure AD, sin Microsoft Graph.** Se probo esa ruta
  (auth app-only contra Graph) pero requeria aprobaciones de IAM que no
  compensaban la complejidad para este caso de uso. En su lugar: cualquiera
  con el archivo lo sube a mano desde el dashboard.
- Cuando alguien sube un `.xlsx` nuevo, se parsea con `openpyxl` y reemplaza
  el contenido de las 3 tablas en SQLite (`dashboard.db`, se crea sola).
- Todos los demas usuarios ven los datos actualizados en su siguiente
  refresco automatico (cada 30s, via HTMX) -- sin recargar la pagina.

Hojas que debe tener el `.xlsx` subido (mismos nombres, sin distinguir mayusculas):

| Hoja | Columnas esperadas | Uso en el dashboard |
|---|---|---|
| `Projects` | App, Overall, S4 Dependency, Product, Engineering, Project Overview, Status, Dependencies, platform | Tabla principal + tarjetas resumen (verde/amarillo/rojo) |
| `Roadmap` | App, Roadmap task, Roadmap Start, Roadmap End, Roadmap % | Grafica de barras de % de avance |
| `Risk` | App, Risk, Mitigation, Resolution date | Tabla de riesgos |

Si al archivo subido le falta alguna de estas 3 hojas, el dashboard rechaza
la subida y muestra un mensaje de error claro (no se pierde lo que ya habia).

## Setup

```bash
uv venv --python 3.13.5
source .venv/bin/activate
uv pip install --index-url https://pypi.ci.artifacts.walmart.com/artifactory/api/pypi/external-pypi/simple \
    --allow-insecure-host pypi.ci.artifacts.walmart.com -r requirements.txt
```

## Correr la app

```bash
source .venv/bin/activate
uvicorn app.main:app --reload --port 8000
```

Abre `http://localhost:8000`. La primera vez no habra datos -- usa el boton
**"Subir Excel actualizado"** para cargar el `.xlsx` con las 3 hojas.

## Estructura del proyecto

```
app/
  config.py        # Configuracion minima (ruta de la DB, limite de tamano de upload)
  db.py             # Esquema SQLite + helpers de lectura/escritura
  excel_import.py   # Parseo del .xlsx subido -> filas normalizadas -> SQLite
  main.py           # FastAPI: rutas (/, /partials/content, /upload)
  templates/        # Jinja2 + HTMX + Tailwind (CDN) + Chart.js (CDN)
  static/           # Assets estaticos (vacio por ahora)
requirements.txt
```

## Notas

- `dashboard.db` (la cache SQLite) no se commitea -- se regenera con cada
  subida de Excel. Si quieres "resetear" el dashboard, simplemente borra el
  archivo y sube el Excel de nuevo.
- No hay control de quien sube el archivo ni historial de versiones -- si
  eso se vuelve necesario mas adelante, es un buen candidato para una
  segunda iteracion (ej. guardar el .xlsx subido con timestamp, mostrar
  quien lo subio).
