# -*- coding: utf-8 -*-
"""
VALIDACIÓN DE ENLACES — Reporte de Hilos (v2: Extremo A y Extremo B)
=====================================================================
Aplicación completa en UN SOLO archivo Python (Flask + SQLite).

FLUJO v2:
  1. "+ Nuevo Enlace"  → se crea solo la FICHA (anillo, nombre, capacidad).
  2. "Completar A"     → se registra Origen A + descripción de cada hilo (lado A).
  3. "Completar B"     → se registra Origen B + descripción de cada hilo (lado B).
  4. Cuando A y B están completos → el enlace pasa a VALIDADO automáticamente
     y aparece el reporte comparativo hilo a hilo (Coincide Sí/No).

CÓMO EJECUTAR:
    pip install flask flask-sqlalchemy openpyxl
    python reporte_hilos.py
    → abrir http://127.0.0.1:5000
"""

import io
import uuid
from datetime import datetime

from flask import (Flask, request, redirect, url_for, send_file,
                   render_template_string)
from flask_sqlalchemy import SQLAlchemy
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///enlaces.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db = SQLAlchemy(app)

# ============================================================
# 1) MODELOS
# ============================================================
class Enlace(db.Model):
    id = db.Column(db.String(36), primary_key=True,
                   default=lambda: str(uuid.uuid4()))
    anillo = db.Column(db.String(30), nullable=False)
    nombre = db.Column(db.String(120), nullable=False)
    capacidad = db.Column(db.Integer, nullable=False, default=24)
    origen_a = db.Column(db.String(80), default="")
    origen_b = db.Column(db.String(80), default="")
    a_completo = db.Column(db.Boolean, default=False)   # ← lado A registrado
    b_completo = db.Column(db.Boolean, default=False)   # ← lado B registrado
    creado = db.Column(db.DateTime, default=datetime.utcnow)
    hilos = db.relationship("Hilo", backref="enlace",
                            cascade="all, delete-orphan", order_by="Hilo.numero")

    @property
    def estado(self):
        """VALIDADO solo cuando AMBOS extremos están completos."""
        return "VALIDADO" if (self.a_completo and self.b_completo) else "INCOMPLETO"

    @property
    def pendiente_texto(self):
        if not self.a_completo and not self.b_completo:
            return "Falta Extremo A y Extremo B"
        if not self.a_completo:
            return "Falta Extremo A"
        if not self.b_completo:
            return "Falta Extremo B"
        return ""


class Hilo(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    enlace_id = db.Column(db.String(36), db.ForeignKey("enlace.id"))
    numero = db.Column(db.Integer)
    descripcion_a = db.Column(db.Text, default="")
    descripcion_b = db.Column(db.Text, default="")

    @property
    def coincide(self):
        na = " ".join((self.descripcion_a or "").upper().split())
        nb = " ".join((self.descripcion_b or "").upper().split())
        return na == nb


# ============================================================
# 2) DATOS INICIALES (tu dashboard: 27 pendientes con A hecho + 3 validados)
#    (anillo, nombre, origen_a, origen_b, capacidad, fecha, a_ok, b_ok)
# ============================================================
SEED = [
    ("ANILLO ESTE",  "SANTA ANITA - LA MOLINA",   "SANTA ANITA", "LA MOLINA",   48, "2026-04-21 17:18:46", True, False),
    ("ANILLO SUR",   "LURIN - TELEPUERTO - 96SM", "LURIN",       "96SM",        96, "2026-04-18 20:10:17", True, False),
    ("ANILLO ESTE",  "AVIACION - HIGUERETA",      "AVIACION",    "HIGUERETA",   96, "2026-03-12 02:40:42", True, False),
    ("ANILLO ESTE",  "CAMACHO - MONTERICO",       "MONTERICO",   "CAMACHO",     48, "2026-03-08 00:52:18", True, False),
    ("ANILLO ESTE",  "SAN LUIS - MOTERRICO",      "MONTERICO",   "SAN LUIS",    48, "2026-03-07 02:30:07", True, False),
    ("ANILLO ESTE",  "HIGUERETA - SAN LUIS",      "HIGUERETA",   "SAN LUIS",    48, "2026-02-05 19:00:47", True, False),
    ("ANILLO ESTE",  "HIGUERETA - TELEPUERTO",    "HIGUERETA",   "TELEPUERTO",  48, "2026-02-05 14:13:29", True, False),
    ("ANILLO SUR",   "OCHARAN - HIGUERETA",       "OCHARAN",     "HIGUERETA",   96, "2026-02-04 21:05:33", True, False),
    ("ANILLO OESTE", "SAN MIGUEL - SAN FELIPE 1", "POP SAN MIGUEL", "SAN FELIPE 1", 48, "2026-01-21 20:40:26", True, False),
    ("ANILLO NORTE", "COTABAMBAS - ZARATE",       "COTABAMBAS",  "ZARATE",      96, "2026-01-20 22:26:54", True, False),
    ("ANILLO NORTE", "COLONIAL - COTABAMBAS",     "COLONIAL",    "COTABAMBAS",  48, "2026-01-20 17:16:31", True, False),
    ("ANILLO OESTE", "AEROPUERTO - CALLAO",       "AEROPUERTO",  "CALLAO",      96, "2026-01-19 20:28:32", True, False),
    ("ANILLO SUR",   "AERIPUETO - LA MILLA",      "AERIPUETO",   "LA MILLA",    24, "2026-01-19 17:26:43", True, False),
    ("ANILLO SUR",   "AEROPUERTO - PAMPA LIBRE",  "AEROPUERTO",  "PAMPA LIBRE", 24, "2026-01-19 14:28:30", True, False),
    ("ANILLO NORTE", "AEROPUERTO - COLONIAL",     "AEROPUERTO",  "COLONIAL",    48, "2026-01-15 16:10:38", True, False),
    ("ANILLO SUR",   "AEROPUERTO - POLO 1",       "AEROPUERTO",  "POLO 1",      24, "2026-01-15 02:07:20", True, False),
    ("ANILLO OESTE", "LA PUNTA - SAN MIGUEL",     "LA PUNTA",    "SAN MIGUEL",  48, "2026-01-14 23:20:24", True, False),
    ("ANILLO OESTE", "CALLO - LA PUNTA",          "CALLAO",      "LA PUNTA",    48, "2026-01-14 22:34:45", True, False),
    ("ANILLO SUR",   "TELEPUERTO - ASIA",         "TELEPUERTO",  "ASIA",        24, "2026-01-14 02:16:00", True, False),
    ("ANILLO NORTE", "SANTA LUZMILA - LOS OLIVOS","SANTA LUZMILA","LOS OLIVOS", 48, "2026-01-12 18:38:00", True, False),
    ("ANILLO NORTE", "AEROPUERTO - LOS OLIVOS",   "AEROPUERTO",  "LOS OLIVOS",  48, "2026-01-12 16:40:00", True, False),
    ("ANILLO ESTE",  "ZARATE - SANTA ANITA",      "ZARATE",      "SANTA ANITA", 48, "2026-01-10 19:30:00", True, False),
    ("ANILLO NORTE", "ZARATE - INGENIERIA",       "ZARATE",      "INGENIERIA",  48, "2026-01-10 17:58:00", True, False),
    ("ANILLO SUR",   "TELPUERTO -POLO 2",         "POLO 2",      "VILLA EL SAVADOR", 24, "2026-01-09 23:12:00", True, False),
    ("ANILLO SUR",   "POLO 2 - POLO 1",           "P0L0 1",      "POLO 2",      24, "2026-01-09 18:01:00", True, False),
    ("ANILLO SUR",   "CHORRILLOS - OCHARAN",      "CHORRILLOS",  "OCHARAN",     96, "2026-01-08 18:39:00", True, False),
    ("ANILLO SUR",   "SAN JUAN - TELPUERTO",      "SAN JUAN",    "TELPUERTO",   96, "2026-01-07 22:31:00", True, False),
    # ---- Validados (A y B completos) ----
    ("ANILLO SUR",   "LURIN - TELEPUERTO - 48SM", "TELEPUERTO",  "LURIN",       48, "2026-01-14 03:15:26", True, True),
    ("ANILLO SUR",   "SAN JUAN - POLO 2",         "SAN JUAN",    "POLO 2",      96, "2026-01-08 00:14:39", True, True),
    ("ANILLO SUR",   "SAN JUAN - CHORRILLOS",     "SAN JUAN",    "CHORRILLOS",  48, "2026-01-07 20:04:48", True, True),
]


def seed_inicial():
    if Enlace.query.count() > 0:
        return
    for anillo, nombre, a, b, cap, fecha, a_ok, b_ok in SEED:
        e = Enlace(anillo=anillo, nombre=nombre, capacidad=cap,
                   origen_a=a, origen_b=b, a_completo=a_ok, b_completo=b_ok,
                   creado=datetime.strptime(fecha, "%Y-%m-%d %H:%M:%S"))
        for n in range(1, cap + 1):
            e.hilos.append(Hilo(numero=n,
                                descripcion_a="LIBRE" if a_ok else "",
                                descripcion_b="LIBRE" if b_ok else ""))
        db.session.add(e)
    db.session.commit()


with app.app_context():
    db.create_all()
    seed_inicial()

# ============================================================
# 3) PLANTILLAS
# ============================================================
TPL_DASH = """
<!doctype html><html lang="es"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Dashboard - Validación</title>
<style>
 body{margin:0;font-family:'Segoe UI',system-ui,sans-serif;background:#fff;color:#1f2937}
 header{display:flex;align-items:center;justify-content:space-between;
        padding:14px 24px;border-bottom:1px solid #e5e7eb}
 header h1{font-size:1.15rem;margin:0}
 .btn{display:inline-block;padding:6px 12px;border-radius:6px;border:1px solid #d1d5db;
      background:#fff;color:#374151;text-decoration:none;font-size:.78rem;cursor:pointer}
 .btn-azul{background:#2563eb;border-color:#2563eb;color:#fff}
 .btn-borde-azul{border-color:#2563eb;color:#2563eb}
 .btn-borde-verde{border-color:#059669;color:#059669}
 .btn-borde-naranja{border-color:#d97706;color:#d97706}
 .grid{display:grid;grid-template-columns:1fr 1fr;gap:24px;padding:24px;
       max-width:1550px;margin:0 auto;align-items:start}
 .card{background:#fff;border:1px solid #e5e7eb;border-radius:12px;
       box-shadow:0 1px 4px rgba(0,0,0,.06);padding:18px}
 .card h2{font-size:1rem;margin:0 0 2px;display:flex;justify-content:space-between}
 .badge{background:#dbeafe;color:#1d4ed8;border-radius:10px;padding:1px 9px;font-size:.75rem}
 .sub{color:#6b7280;font-size:.75rem;margin:0 0 12px}
 table{width:100%;border-collapse:collapse;font-size:.8rem}
 th{text-align:left;padding:6px 4px;border-bottom:2px solid #e5e7eb}
 td{padding:9px 4px;border-bottom:1px solid #f3f4f6;vertical-align:middle}
 .tramo{color:#9ca3af;font-size:.72rem}
 .falta{color:#d97706;font-size:.7rem;font-weight:600}
 .acc{display:flex;gap:5px;flex-wrap:wrap}
</style></head><body>
<header>
  <h1>🧵 Validación de Enlaces</h1>
  <a class="btn btn-azul" href="{{ url_for('nuevo') }}">+ Nuevo Enlace</a>
</header>
<div class="grid">

  <div class="card">
    <h2>Pendientes (Incompletos) <span class="badge">{{ pend|length }}</span></h2>
    <p class="sub">Enlaces con Extremo A y/o Extremo B pendientes de completar.</p>
    <table>
      <tr><th>Enlace</th><th>Cap.</th><th>Creado</th><th>Acciones</th></tr>
      {% for e in pend %}
      <tr>
        <td><strong>{{ e.anillo }} | {{ e.nombre }}</strong><br>
            <span class="tramo">{{ e.origen_a or '?' }} → {{ e.origen_b or '?' }}</span><br>
            <span class="falta">{{ e.pendiente_texto }}</span></td>
        <td>{{ e.capacidad }}</td>
        <td>{{ e.creado.strftime('%Y-%m-%d %H:%M:%S') }}</td>
        <td class="acc">
          {% if not e.a_completo %}
            <a class="btn btn-borde-naranja" href="{{ url_for('completar', id=e.id, lado='a') }}">Completar A</a>
          {% endif %}
          {% if not e.b_completo %}
            <a class="btn btn-borde-azul" href="{{ url_for('completar', id=e.id, lado='b') }}">Completar B</a>
          {% endif %}
          <a class="btn" href="{{ url_for('editar', id=e.id) }}">Editar</a>
          <a class="btn btn-borde-verde" href="{{ url_for('exportar', id=e.id) }}">Exportar</a>
        </td>
      </tr>
      {% endfor %}
    </table>
  </div>

  <div class="card">
    <h2>Validados <span class="badge">{{ valid|length }}</span></h2>
    <p class="sub">Enlaces con Extremo A y B completados (reporte comparativo).</p>
    <table>
      <tr><th>Enlace</th><th>Cap.</th><th>Creado</th><th>Acciones</th></tr>
      {% for e in valid %}
      <tr>
        <td><strong>{{ e.anillo }} | {{ e.nombre }}</strong><br>
            <span class="tramo">{{ e.origen_a }} ⇄ {{ e.origen_b }}</span></td>
        <td>{{ e.capacidad }}</td>
        <td>{{ e.creado.strftime('%Y-%m-%d %H:%M:%S') }}</td>
        <td class="acc">
          <a class="btn" href="{{ url_for('ver', id=e.id) }}">Ver</a>
          <a class="btn" href="{{ url_for('editar', id=e.id) }}">Editar</a>
          <a class="btn btn-borde-verde" href="{{ url_for('exportar', id=e.id) }}">Exportar</a>
        </td>
      </tr>
      {% endfor %}
    </table>
  </div>

</div></body></html>
"""

TPL_FORM = """
<!doctype html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{{ titulo }}</title>
<style>
 body{font-family:'Segoe UI',sans-serif;max-width:940px;margin:24px auto;color:#1f2937;padding:0 12px}
 label{display:block;font-size:.8rem;color:#6b7280;margin-top:10px}
 input,select,textarea{width:100%;padding:6px;border:1px solid #d1d5db;border-radius:6px;box-sizing:border-box}
 table{width:100%;border-collapse:collapse;font-size:.8rem;margin-top:14px}
 th,td{border:1px solid #e5e7eb;padding:5px}
 th{background:#f3f4f6}
 .btn{display:inline-block;margin-top:16px;padding:8px 16px;border-radius:6px;
      background:#2563eb;color:#fff;border:0;cursor:pointer;text-decoration:none}
 .rojo{background:#fee2e2}
 .ref{color:#6b7280}
</style></head><body>
<h2>{{ titulo }}</h2>
{{ cuerpo|safe }}
</body></html>
"""

# ============================================================
# 4) RUTAS
# ============================================================
@app.route("/")
def dashboard():
    todos = Enlace.query.order_by(Enlace.creado.desc()).all()
    pend = [e for e in todos if e.estado == "INCOMPLETO"]
    valid = [e for e in todos if e.estado == "VALIDADO"]
    return render_template_string(TPL_DASH, pend=pend, valid=valid)


@app.route("/nuevo", methods=["GET", "POST"])
def nuevo():
    """Crea solo la FICHA del enlace. A y B se completan después."""
    if request.method == "POST":
        cap = int(request.form.get("capacidad") or 24)
        e = Enlace(anillo=request.form.get("anillo", ""),
                   nombre=request.form.get("nombre", "").strip(),
                   capacidad=cap,
                   origen_a=request.form.get("origen_a", "").strip(),
                   origen_b=request.form.get("origen_b", "").strip())
        for n in range(1, cap + 1):
            e.hilos.append(Hilo(numero=n))
        db.session.add(e)
        db.session.commit()
        return redirect(url_for("dashboard"))

    cuerpo = """
    <p class="ref">Se crea la ficha del enlace. Luego usa <b>Completar A</b> y
    <b>Completar B</b> desde el dashboard para registrar cada extremo.</p>
    <form method="post">
      <label>Anillo</label>
      <select name="anillo">
        <option>ANILLO NORTE</option><option>ANILLO SUR</option>
        <option>ANILLO ESTE</option><option>ANILLO OESTE</option>
      </select>
      <label>Nombre del enlace (SITIO A - SITIO B)</label><input name="nombre" required>
      <label>Capacidad (hilos)</label>
      <select name="capacidad"><option>24</option><option selected>48</option><option>96</option></select>
      <label>Origen A (sugerido)</label><input name="origen_a">
      <label>Origen B (sugerido)</label><input name="origen_b">
      <button class="btn">Crear enlace</button>
      <a class="btn" style="background:#6b7280" href="/">Cancelar</a>
    </form>"""
    return render_template_string(TPL_FORM, titulo="Nuevo Enlace", cuerpo=cuerpo)


@app.route("/enlace/<id>/completar/<lado>", methods=["GET", "POST"])
def completar(id, lado):
    """Formulario para completar el Extremo A o el Extremo B."""
    if lado not in ("a", "b"):
        return redirect(url_for("dashboard"))
    e = Enlace.query.get_or_404(id)
    es_a = (lado == "a")

    if request.method == "POST":
        origen = request.form.get("origen", "").strip()
        if es_a:
            e.origen_a = origen or e.origen_a
            for h in e.hilos:
                h.descripcion_a = request.form.get(f"h{h.numero}", "").strip()
            e.a_completo = True
        else:
            e.origen_b = origen or e.origen_b
            for h in e.hilos:
                h.descripcion_b = request.form.get(f"h{h.numero}", "").strip()
            e.b_completo = True
        db.session.commit()
        # Si con esto quedó VALIDADO, ir directo al comparativo
        if e.estado == "VALIDADO":
            return redirect(url_for("ver", id=e.id))
        return redirect(url_for("dashboard"))

    # --- GET: armar formulario ---
    letra = "A" if es_a else "B"
    origen_actual = e.origen_a if es_a else e.origen_b
    # el otro lado se muestra como referencia si ya existe
    otro_ok = e.b_completo if es_a else e.a_completo
    otro_nom = e.origen_b if es_a else e.origen_a

    filas = []
    for h in e.hilos:
        ref = h.descripcion_b if es_a else h.descripcion_a
        val = h.descripcion_a if es_a else h.descripcion_b
        celda_ref = (f"<td class='ref'>{ref}</td>" if otro_ok else "")
        filas.append(f"<tr><td>{h.numero}</td>{celda_ref}"
                     f"<td><input name='h{h.numero}' value=\"{val}\" "
                     f"placeholder='LIBRE / ODF / SIN PIGTAIL...'></td></tr>")
    th_ref = (f"<th>Extremo {'B' if es_a else 'A'} — {otro_nom} (referencia)</th>"
              if otro_ok else "")

    cuerpo = f"""
    <p><b>{e.anillo} | {e.nombre}</b> · {e.capacidad} hilos ·
       Estado actual: {e.estado} ({e.pendiente_texto or 'completo'})</p>
    <form method="post">
      <label>Origen (Sitio {letra})</label>
      <input name="origen" value="{origen_actual}" required>
      <table>
        <tr><th>Fibra</th>{th_ref}<th>Descripción Extremo {letra}</th></tr>
        {''.join(filas)}
      </table>
      <button class="btn" style="background:{'#d97706' if es_a else '#2563eb'}">
        Guardar Extremo {letra}</button>
      <a class="btn" style="background:#6b7280" href="/">Cancelar</a>
    </form>"""
    return render_template_string(TPL_FORM,
                                  titulo=f"Completar Extremo {letra}", cuerpo=cuerpo)


@app.route("/enlace/<id>/editar", methods=["GET", "POST"])
def editar(id):
    e = Enlace.query.get_or_404(id)
    if request.method == "POST":
        e.nombre = request.form.get("nombre", e.nombre)
        e.origen_a = request.form.get("origen_a", e.origen_a)
        e.origen_b = request.form.get("origen_b", e.origen_b)
        for h in e.hilos:
            h.descripcion_a = request.form.get(f"a{h.numero}", h.descripcion_a)
            h.descripcion_b = request.form.get(f"b{h.numero}", h.descripcion_b)
        db.session.commit()
        return redirect(url_for("dashboard"))

    filas = "".join(
        f"<tr><td>{h.numero}</td>"
        f"<td><input name='a{h.numero}' value=\"{h.descripcion_a}\"></td>"
        f"<td><input name='b{h.numero}' value=\"{h.descripcion_b}\"></td></tr>"
        for h in e.hilos)
    cuerpo = f"""
    <form method="post">
      <label>Nombre</label><input name="nombre" value="{e.nombre}">
      <label>Origen A</label><input name="origen_a" value="{e.origen_a}">
      <label>Origen B</label><input name="origen_b" value="{e.origen_b}">
      <table><tr><th>Fibra</th><th>Extremo A</th><th>Extremo B</th></tr>{filas}</table>
      <button class="btn">Guardar cambios</button>
      <a class="btn" style="background:#6b7280" href="/">Cancelar</a>
    </form>"""
    return render_template_string(TPL_FORM, titulo="Editar Enlace", cuerpo=cuerpo)


@app.route("/enlace/<id>/ver")
def ver(id):
    e = Enlace.query.get_or_404(id)
    difs = sum(1 for h in e.hilos if not h.coincide)
    filas = "".join(
        f"<tr class='{'rojo' if not h.coincide else ''}'>"
        f"<td>{h.numero}</td><td>{h.descripcion_a}</td>"
        f"<td>{h.descripcion_b}</td><td>{'Sí' if h.coincide else 'No'}</td></tr>"
        for h in e.hilos)
    aviso = (f"<p style='color:#b45309'>⚠ {difs} hilo(s) con diferencias entre A y B.</p>"
             if difs else "<p style='color:#059669'>✔ Todos los hilos coinciden.</p>")
    cuerpo = f"""
    <p><b>{e.anillo} | {e.nombre}</b> · {e.capacidad} hilos ·
       {e.origen_a} ⇄ {e.origen_b} · Estado: {e.estado}</p>
    {aviso}
    <table><tr><th>Fibra</th><th>{e.origen_a or 'A'}</th>
           <th>{e.origen_b or 'B'}</th><th>Coincide</th></tr>{filas}</table>
    <a class="btn" style="background:#059669"
       href="{url_for('exportar', id=e.id)}">Exportar Excel</a>
    <a class="btn" style="background:#6b7280" href="/">Volver</a>"""
    return render_template_string(TPL_FORM, titulo="Reporte Comparativo", cuerpo=cuerpo)


@app.route("/enlace/<id>/export.xlsx")
def exportar(id):
    e = Enlace.query.get_or_404(id)
    wb = Workbook(); ws = wb.active; ws.title = "REPORTE"

    azul = PatternFill("solid", start_color="1F4E79")
    rojo = PatternFill("solid", start_color="F8CBAD")
    thin = Side(style="thin", color="BFBFBF")
    borde = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws["A1"] = f"{e.anillo} | {e.nombre}"
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws["A2"] = (f"Capacidad: {e.capacidad} · {e.origen_a or '?'} ⇄ {e.origen_b or '?'} · "
                f"Estado: {e.estado} · A:{'✔' if e.a_completo else 'pend.'} "
                f"B:{'✔' if e.b_completo else 'pend.'} · "
                f"Creado: {e.creado:%d/%m/%Y %H:%M}")
    ws["A2"].font = Font(size=9, color="595959")

    ws.append([]); ws.append(["FIBRA", f"DESCRIPCIÓN {e.origen_a or 'A'}",
                              f"DESCRIPCIÓN {e.origen_b or 'B'}", "COINCIDE"])
    hdr = ws.max_row
    for c in range(1, 5):
        cell = ws.cell(hdr, c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = azul
        cell.alignment = Alignment(horizontal="center")
        cell.border = borde

    for h in e.hilos:
        ws.append([h.numero, h.descripcion_a, h.descripcion_b,
                   "Sí" if h.coincide else "No"])
        r = ws.max_row
        for c in range(1, 5):
            ws.cell(r, c).border = borde
        if not h.coincide:
            ws.cell(r, 4).fill = rojo

    for col, w in zip("ABCD", [8, 50, 50, 10]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = f"A{hdr+1}"

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"{e.nombre.replace(' ', '_')}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument"
                              ".spreadsheetml.sheet")


if __name__ == "__main__":
    app.run(debug=True)
