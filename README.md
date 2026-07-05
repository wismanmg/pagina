# Validación de Enlaces — v5

## Novedades de esta versión
- **Login con roles** (editor / lector) — Flask-Login. Demo: admin/admin123, lector/lector123.
- **CSRF** en todos los formularios (Flask-WTF) y **escape automático** (plantillas Jinja2 reales, sin f-strings HTML → sin XSS).
- **Arquitectura por módulos**: config.py, models.py, auth.py, routes.py, reports.py, templates/.
- **Validación en servidor**: nombre obligatorio/único/≤120, anillo del catálogo, capacidad en {12,24,48,96,144}.
- **Transacciones** con rollback ante error en todas las escrituras.
- **PostgreSQL** vía variable DATABASE_URL (Render) con SQLite de respaldo local; **Flask-Migrate** para cambios de esquema.
- **Auditoría**: modificado_por/en en cada enlace + tabla AuditLog (historial visible en "Ver").
- **Papelera**: borrado lógico + restaurar.
- **Catálogo de sitios normalizado** (corrige AERIPUETO→AEROPUERTO, CALLO→CALLAO, P0L0→POLO, TELPUERTO→TELEPUERTO, MOTERRICO→MONTERRICO, SAVADOR→SALVADOR) con autocompletado en formularios.
- **Dashboard**: buscador, filtro por anillo, **semáforo de antigüedad** (<30 verde, 30-89 ámbar, ≥90 rojo) y **paginación**.
- **Comparación inteligente**: 'ODF 5' == 'ODF-5' == 'odf  5'.
- **Logging** en operaciones de escritura; **/health** para monitoreo; SECRET_KEY/DEBUG por variables de entorno.
- **Suite pytest** (tests/): 19 pruebas.

## Ejecutar en local
    pip install -r requirements.txt
    python app.py            # http://127.0.0.1:5000  (login: admin/admin123)

## Ejecutar pruebas
    pytest tests/ -v

## Deploy en Render
1. Sube esta carpeta al repo de GitHub.
2. Web Service tipo Docker (usa el Dockerfile) o Python
   (build: pip install -r requirements.txt · start: gunicorn app:app).
3. Variables de entorno: SECRET_KEY (una cadena larga aleatoria)
   y DATABASE_URL (crea un PostgreSQL gratuito en Render y copia su Internal URL).
