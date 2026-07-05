# -*- coding: utf-8 -*-
"""Rutas principales: dashboard, CRUD de enlaces, papelera, reportes, salud."""
import logging

from flask import (Blueprint, render_template, request, redirect,
                   url_for, send_file, flash, abort, current_app)
from flask_login import login_required

from auth import solo_editor
from models import db, Enlace, Hilo, Sitio, AuditLog, auditar
from reports import excel_enlace, excel_avance

main_bp = Blueprint("main", __name__)
log = logging.getLogger("enlaces")

ANILLOS = ["ANILLO NORTE", "ANILLO SUR", "ANILLO ESTE", "ANILLO OESTE"]
TIPOS_CABLE = ["CABLE DE ENLACE SM", "ADSS", "OPGW", "FIGURA 8",
               "ARMADO", "DUCTO", "OTRO"]


# ----------------------------------------------------------------------
# Validaciones de servidor
# ----------------------------------------------------------------------
def validar_ficha(form, enlace_actual=None):
    """Valida los datos de cabecera. Devuelve (ok, errores, datos)."""
    errores = []
    nombre = (form.get("nombre") or form.get("nombre_manual") or "").strip()
    if enlace_actual and not nombre:
        nombre = enlace_actual.nombre
    if not nombre:
        errores.append("El nombre del enlace es obligatorio.")
    elif len(nombre) > 120:
        errores.append("El nombre no puede superar 120 caracteres.")
    else:
        q = Enlace.query.filter(Enlace.nombre == nombre,
                                Enlace.eliminado.is_(False))
        if enlace_actual:
            q = q.filter(Enlace.id != enlace_actual.id)
        if q.first():
            errores.append(f"Ya existe un enlace llamado '{nombre}'.")

    anillo = (form.get("anillo") or
              (enlace_actual.anillo if enlace_actual else "")).strip()
    if anillo not in ANILLOS:
        errores.append("Anillo inválido.")

    try:
        cap = int(form.get("capacidad") or
                  (enlace_actual.capacidad if enlace_actual else 24))
    except (TypeError, ValueError):
        cap = -1
    if cap not in Enlace.CAPACIDADES_VALIDAS:
        errores.append(f"Capacidad inválida "
                       f"(permitidas: {Enlace.CAPACIDADES_VALIDAS}).")

    tipo = (form.get("tipo_cable") or "CABLE DE ENLACE SM").strip()[:60]
    longitud = (form.get("longitud") or "").strip()[:40]
    return (not errores), errores, dict(nombre=nombre, anillo=anillo,
                                        capacidad=cap, tipo_cable=tipo,
                                        longitud=longitud)


def _get_enlace(id):
    e = db.session.get(Enlace, id)
    if not e or e.eliminado:
        abort(404)
    return e


# ----------------------------------------------------------------------
# Dashboard con buscador, filtro por anillo, semáforo y paginación
# ----------------------------------------------------------------------
@main_bp.route("/")
@login_required
def dashboard():
    q = (request.args.get("q") or "").strip()
    anillo = (request.args.get("anillo") or "").strip()
    page = max(1, request.args.get("page", 1, type=int))
    per_page = current_app.config["ITEMS_POR_PAGINA"]

    base = Enlace.query.filter(Enlace.eliminado.is_(False))
    if q:
        like = f"%{q}%"
        base = base.filter(db.or_(Enlace.nombre.ilike(like),
                                  Enlace.origen_a.ilike(like),
                                  Enlace.origen_b.ilike(like)))
    if anillo in ANILLOS:
        base = base.filter(Enlace.anillo == anillo)

    todos = base.order_by(Enlace.creado.desc()).all()
    pend_all = [e for e in todos if e.estado == "INCOMPLETO"]
    valid = [e for e in todos if e.estado == "VALIDADO"]

    total_pend = len(pend_all)
    paginas = max(1, -(-total_pend // per_page))   # ceil
    page = min(page, paginas)
    pend = pend_all[(page - 1) * per_page: page * per_page]

    return render_template("dashboard.html",
                           pend=pend, valid=valid, total_pend=total_pend,
                           page=page, paginas=paginas, q=q,
                           anillo=anillo, anillos=ANILLOS)


# ----------------------------------------------------------------------
# CRUD
# ----------------------------------------------------------------------
@main_bp.route("/nuevo", methods=["GET", "POST"])
@solo_editor
def nuevo():
    if request.method == "POST":
        ok, errores, d = validar_ficha(request.form)
        if not ok:
            for m in errores:
                flash(m, "error")
        else:
            try:
                e = Enlace(**d,
                           origen_a=(request.form.get("origen_a") or "").strip()[:80],
                           origen_b=(request.form.get("origen_b") or "").strip()[:80])
                for n in range(1, d["capacidad"] + 1):
                    e.hilos.append(Hilo(numero=n))
                e.tocar()
                db.session.add(e)
                auditar("CREAR", e.id, f"Enlace {e.nombre}")
                db.session.commit()
                log.info("Enlace creado: %s", e.nombre)
                flash("Enlace creado.", "ok")
                return redirect(url_for("main.dashboard"))
            except Exception:
                db.session.rollback()
                log.exception("Error creando enlace")
                flash("Error interno al crear el enlace.", "error")
    sitios = [s.nombre for s in Sitio.query.order_by(Sitio.nombre)]
    return render_template("nuevo.html", anillos=ANILLOS,
                           capacidades=Enlace.CAPACIDADES_VALIDAS,
                           sitios=sitios)


@main_bp.route("/enlace/<id>/completar/<lado>", methods=["GET", "POST"])
@solo_editor
def completar(id, lado):
    if lado not in ("a", "b"):
        abort(404)
    e = _get_enlace(id)
    es_a = lado == "a"
    if request.method == "POST":
        origen = (request.form.get("origen") or "").strip()[:80]
        if not origen:
            flash("El origen es obligatorio.", "error")
        else:
            try:
                sala = (request.form.get("sala") or "").strip()[:80]
                rack = (request.form.get("rack") or "").strip()[:40]
                pos = (request.form.get("posicion") or "").strip()[:40]
                if es_a:
                    e.origen_a, e.sala_a, e.rack_a, e.posicion_a = origen, sala, rack, pos
                    for h in e.hilos:
                        h.descripcion_a = (request.form.get(f"h{h.numero}") or "").strip()
                    e.a_completo = True
                else:
                    e.origen_b, e.sala_b, e.rack_b, e.posicion_b = origen, sala, rack, pos
                    for h in e.hilos:
                        h.descripcion_b = (request.form.get(f"h{h.numero}") or "").strip()
                    e.b_completo = True
                e.tocar()
                auditar(f"COMPLETAR_{lado.upper()}", e.id, e.nombre)
                db.session.commit()
                log.info("Extremo %s completado: %s", lado.upper(), e.nombre)
                if e.estado == "VALIDADO":
                    flash("¡Enlace VALIDADO! Ambos extremos completos.", "ok")
                    return redirect(url_for("main.ver", id=e.id))
                return redirect(url_for("main.dashboard"))
            except Exception:
                db.session.rollback()
                log.exception("Error completando extremo")
                flash("Error interno al guardar.", "error")
    return render_template("completar.html", e=e, es_a=es_a,
                           letra="A" if es_a else "B")


@main_bp.route("/enlace/<id>/editar", methods=["GET", "POST"])
@solo_editor
def editar(id):
    e = _get_enlace(id)
    if request.method == "POST":
        form = dict(request.form)
        sel = (form.get("nombre_sel") or "").strip()
        manual = (form.get("nombre_manual") or "").strip()
        form["nombre"] = manual or ("" if sel in ("", "__actual__") else sel)
        ok, errores, d = validar_ficha(form, enlace_actual=e)
        if not ok:
            for m in errores:
                flash(m, "error")
        else:
            try:
                nueva_cap = d["capacidad"]
                if nueva_cap != e.capacidad:
                    if nueva_cap < e.capacidad:
                        for h in list(e.hilos):
                            if h.numero > nueva_cap:
                                db.session.delete(h)
                    else:
                        for n in range(e.capacidad + 1, nueva_cap + 1):
                            e.hilos.append(Hilo(numero=n))
                e.nombre, e.anillo = d["nombre"], d["anillo"]
                e.capacidad, e.tipo_cable = nueva_cap, d["tipo_cable"]
                e.longitud = d["longitud"]
                for campo in ("origen_a", "sala_a", "rack_a", "posicion_a",
                              "origen_b", "sala_b", "rack_b", "posicion_b"):
                    if campo in request.form:
                        setattr(e, campo, (request.form.get(campo) or "").strip()[:80])
                for h in e.hilos:
                    if f"a{h.numero}" in request.form:
                        h.descripcion_a = request.form.get(f"a{h.numero}", "")
                    if f"b{h.numero}" in request.form:
                        h.descripcion_b = request.form.get(f"b{h.numero}", "")
                e.tocar()
                auditar("EDITAR", e.id, e.nombre)
                db.session.commit()
                log.info("Enlace editado: %s", e.nombre)
                flash("Cambios guardados.", "ok")
                return redirect(url_for("main.dashboard"))
            except Exception:
                db.session.rollback()
                log.exception("Error editando enlace")
                flash("Error interno al guardar cambios.", "error")
    sitios = [s.nombre for s in Sitio.query.order_by(Sitio.nombre)]
    return render_template("editar.html", e=e, anillos=ANILLOS,
                           tipos=TIPOS_CABLE, sitios=sitios,
                           capacidades=Enlace.CAPACIDADES_VALIDAS)


@main_bp.route("/enlace/<id>/ver")
@login_required
def ver(id):
    e = _get_enlace(id)
    difs = sum(1 for h in e.hilos if not h.coincide)
    historial = (AuditLog.query.filter_by(enlace_id=e.id)
                 .order_by(AuditLog.fecha.desc()).limit(10).all())
    return render_template("ver.html", e=e, difs=difs, historial=historial)


# ----------------------------------------------------------------------
# Borrado lógico + papelera
# ----------------------------------------------------------------------
@main_bp.route("/enlace/<id>/eliminar", methods=["POST"])
@solo_editor
def eliminar(id):
    e = _get_enlace(id)
    e.eliminado = True
    e.tocar()
    auditar("ELIMINAR", e.id, e.nombre)
    db.session.commit()
    log.info("Enlace enviado a papelera: %s", e.nombre)
    flash(f"'{e.nombre}' enviado a la papelera.", "ok")
    return redirect(url_for("main.dashboard"))


@main_bp.route("/papelera")
@login_required
def papelera():
    els = (Enlace.query.filter(Enlace.eliminado.is_(True))
           .order_by(Enlace.modificado_en.desc()).all())
    return render_template("papelera.html", enlaces=els)


@main_bp.route("/enlace/<id>/restaurar", methods=["POST"])
@solo_editor
def restaurar(id):
    e = db.session.get(Enlace, id)
    if not e or not e.eliminado:
        abort(404)
    e.eliminado = False
    e.tocar()
    auditar("RESTAURAR", e.id, e.nombre)
    db.session.commit()
    flash(f"'{e.nombre}' restaurado.", "ok")
    return redirect(url_for("main.papelera"))


# ----------------------------------------------------------------------
# Reportes / exportaciones
# ----------------------------------------------------------------------
@main_bp.route("/enlace/<id>/export.xlsx")
@login_required
def exportar(id):
    e = _get_enlace(id)
    buf, nombre = excel_enlace(e)
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype="application/vnd.openxmlformats-officedocument"
                              ".spreadsheetml.sheet")


@main_bp.route("/reporte-avance.xlsx")
@login_required
def reporte_avance():
    enlaces = Enlace.query.filter(Enlace.eliminado.is_(False)).all()
    buf, nombre = excel_avance(enlaces)
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype="application/vnd.openxmlformats-officedocument"
                              ".spreadsheetml.sheet")


# ----------------------------------------------------------------------
# Healthcheck (para Render / monitoreo)
# ----------------------------------------------------------------------
@main_bp.route("/health")
def health():
    try:
        db.session.execute(db.text("SELECT 1"))
        return {"status": "ok"}, 200
    except Exception:
        return {"status": "error"}, 500
