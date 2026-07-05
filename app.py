# -*- coding: utf-8 -*-
"""Punto de entrada. Arranque: gunicorn app:app"""
import logging
from datetime import datetime

from flask import Flask, render_template
from flask_migrate import Migrate
from flask_wtf import CSRFProtect

from jinja2 import ChoiceLoader, DictLoader

from config import Config
from plantillas import TEMPLATES
from models import db, Usuario, Sitio, Enlace, Hilo
from auth import auth_bp, login_manager
from routes import main_bp

csrf = CSRFProtect()
migrate = Migrate()

# --- Catálogo de sitios NORMALIZADO (corrige tipeos históricos) ---
SITIOS = ["AEROPUERTO", "ASIA", "AVIACION", "CALLAO", "CAMACHO", "CASAPALCA",
          "CHORRILLOS", "CHOQUE", "COLONIAL", "COTABAMBAS", "HIGUERETA",
          "HIJOS DE JERUSALEN", "INGENIERIA", "LA MILLA", "LA MOLINA",
          "LA PUNTA", "LOS OLIVOS", "LURIN", "MONTERRICO", "OCHARAN",
          "PAMPA LIBRE", "POLO 1", "POLO 2", "POP SAN MIGUEL", "SAN FELIPE 1",
          "SAN FELIPE 2", "SAN JUAN", "SAN LUIS", "SAN MIGUEL", "SANTA ANITA",
          "SANTA LUZMILA", "TELEPUERTO", "VENTANILLA", "VILLA EL SALVADOR",
          "ZARATE"]

SEED = [  # (anillo, nombre, a, b, cap, fecha, a_ok, b_ok)
    ("ANILLO ESTE", "SANTA ANITA - LA MOLINA", "SANTA ANITA", "LA MOLINA", 48, "2026-04-21 17:18:46", 1, 0),
    ("ANILLO SUR", "LURIN - TELEPUERTO - 96SM", "LURIN", "TELEPUERTO", 96, "2026-04-18 20:10:17", 1, 0),
    ("ANILLO ESTE", "AVIACION - HIGUERETA", "AVIACION", "HIGUERETA", 96, "2026-03-12 02:40:42", 1, 0),
    ("ANILLO ESTE", "CAMACHO - MONTERRICO", "MONTERRICO", "CAMACHO", 48, "2026-03-08 00:52:18", 1, 0),
    ("ANILLO ESTE", "SAN LUIS - MONTERRICO", "MONTERRICO", "SAN LUIS", 48, "2026-03-07 02:30:07", 1, 0),
    ("ANILLO ESTE", "HIGUERETA - SAN LUIS", "HIGUERETA", "SAN LUIS", 48, "2026-02-05 19:00:47", 1, 0),
    ("ANILLO ESTE", "HIGUERETA - TELEPUERTO", "HIGUERETA", "TELEPUERTO", 48, "2026-02-05 14:13:29", 1, 0),
    ("ANILLO SUR", "OCHARAN - HIGUERETA", "OCHARAN", "HIGUERETA", 96, "2026-02-04 21:05:33", 1, 0),
    ("ANILLO OESTE", "SAN MIGUEL - SAN FELIPE 1", "POP SAN MIGUEL", "SAN FELIPE 1", 48, "2026-01-21 20:40:26", 1, 0),
    ("ANILLO NORTE", "COTABAMBAS - ZARATE", "COTABAMBAS", "ZARATE", 96, "2026-01-20 22:26:54", 1, 0),
    ("ANILLO NORTE", "COLONIAL - COTABAMBAS", "COLONIAL", "COTABAMBAS", 48, "2026-01-20 17:16:31", 1, 0),
    ("ANILLO OESTE", "AEROPUERTO - CALLAO", "AEROPUERTO", "CALLAO", 96, "2026-01-19 20:28:32", 1, 0),
    ("ANILLO SUR", "AEROPUERTO - LA MILLA", "AEROPUERTO", "LA MILLA", 24, "2026-01-19 17:26:43", 1, 0),
    ("ANILLO SUR", "AEROPUERTO - PAMPA LIBRE", "AEROPUERTO", "PAMPA LIBRE", 24, "2026-01-19 14:28:30", 1, 0),
    ("ANILLO NORTE", "AEROPUERTO - COLONIAL", "AEROPUERTO", "COLONIAL", 48, "2026-01-15 16:10:38", 1, 0),
    ("ANILLO SUR", "AEROPUERTO - POLO 1", "AEROPUERTO", "POLO 1", 24, "2026-01-15 02:07:20", 1, 0),
    ("ANILLO OESTE", "LA PUNTA - SAN MIGUEL", "LA PUNTA", "SAN MIGUEL", 48, "2026-01-14 23:20:24", 1, 0),
    ("ANILLO OESTE", "CALLAO - LA PUNTA", "CALLAO", "LA PUNTA", 48, "2026-01-14 22:34:45", 1, 0),
    ("ANILLO SUR", "TELEPUERTO - ASIA", "TELEPUERTO", "ASIA", 24, "2026-01-14 02:16:00", 1, 0),
    ("ANILLO NORTE", "SANTA LUZMILA - LOS OLIVOS", "SANTA LUZMILA", "LOS OLIVOS", 48, "2026-01-12 18:38:00", 1, 0),
    ("ANILLO NORTE", "AEROPUERTO - LOS OLIVOS", "AEROPUERTO", "LOS OLIVOS", 48, "2026-01-12 16:40:00", 1, 0),
    ("ANILLO ESTE", "ZARATE - SANTA ANITA", "ZARATE", "SANTA ANITA", 48, "2026-01-10 19:30:00", 1, 0),
    ("ANILLO NORTE", "ZARATE - INGENIERIA", "ZARATE", "INGENIERIA", 48, "2026-01-10 17:58:00", 1, 0),
    ("ANILLO SUR", "TELEPUERTO - POLO 2", "POLO 2", "VILLA EL SALVADOR", 24, "2026-01-09 23:12:00", 1, 0),
    ("ANILLO SUR", "POLO 2 - POLO 1", "POLO 1", "POLO 2", 24, "2026-01-09 18:01:00", 1, 0),
    ("ANILLO SUR", "CHORRILLOS - OCHARAN", "CHORRILLOS", "OCHARAN", 96, "2026-01-08 18:39:00", 1, 0),
    ("ANILLO SUR", "SAN JUAN - TELEPUERTO", "SAN JUAN", "TELEPUERTO", 96, "2026-01-07 22:31:00", 1, 0),
    ("ANILLO SUR", "LURIN - TELEPUERTO - 48SM", "TELEPUERTO", "LURIN", 48, "2026-01-14 03:15:26", 1, 1),
    ("ANILLO SUR", "SAN JUAN - POLO 2", "SAN JUAN", "POLO 2", 96, "2026-01-08 00:14:39", 1, 1),
    ("ANILLO SUR", "SAN JUAN - CHORRILLOS", "SAN JUAN", "CHORRILLOS", 48, "2026-01-07 20:04:48", 1, 1),
]


def seed(app):
    with app.app_context():
        db.create_all()
        if Usuario.query.count() == 0:
            admin = Usuario(username="admin", rol="editor")
            admin.set_password("admin123")
            lector = Usuario(username="lector", rol="lector")
            lector.set_password("lector123")
            db.session.add_all([admin, lector])
        if Sitio.query.count() == 0:
            db.session.add_all([Sitio(nombre=s) for s in SITIOS])
        if Enlace.query.count() == 0:
            for anillo, nombre, a, b, cap, fecha, a_ok, b_ok in SEED:
                e = Enlace(anillo=anillo, nombre=nombre, capacidad=cap,
                           origen_a=a, origen_b=b,
                           a_completo=bool(a_ok), b_completo=bool(b_ok),
                           creado=datetime.strptime(fecha, "%Y-%m-%d %H:%M:%S"))
                for n in range(1, cap + 1):
                    e.hilos.append(Hilo(numero=n,
                                        descripcion_a="LIBRE" if a_ok else "",
                                        descripcion_b="LIBRE" if b_ok else ""))
                db.session.add(e)
        db.session.commit()


def create_app(config_class=Config):
    app = Flask(__name__)
    app.config.from_object(config_class)

    # Plantillas: usa templates/ si existe; si no, las incrustadas en plantillas.py.
    # Asi el deploy funciona subiendo SOLO archivos sueltos (sin carpetas).
    app.jinja_loader = ChoiceLoader([app.jinja_loader, DictLoader(TEMPLATES)])

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")

    db.init_app(app)
    migrate.init_app(app, db)
    csrf.init_app(app)
    login_manager.init_app(app)
    app.register_blueprint(auth_bp)
    app.register_blueprint(main_bp)

    @app.errorhandler(403)
    def forbidden(_):
        return render_template("error.html", codigo=403,
                               mensaje="Tu rol de lectura no permite esta acción."), 403

    @app.errorhandler(404)
    def not_found(_):
        return render_template("error.html", codigo=404,
                               mensaje="No encontrado."), 404

    if not app.config.get("TESTING"):
        seed(app)
    return app


app = create_app()

if __name__ == "__main__":
    app.run(debug=app.config["DEBUG"])
