# -*- coding: utf-8 -*-
"""Suite de pruebas automatizadas (pytest)."""
import io
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
import openpyxl

from app import create_app, SEED
from config import TestConfig
from models import db, Usuario, Sitio, Enlace, Hilo, AuditLog, normalizar
from datetime import datetime


@pytest.fixture
def app():
    app = create_app(TestConfig)
    with app.app_context():
        db.create_all()
        admin = Usuario(username="admin", rol="editor"); admin.set_password("admin123")
        lector = Usuario(username="lector", rol="lector"); lector.set_password("lector123")
        db.session.add_all([admin, lector, Sitio(nombre="SAN JUAN")])
        for anillo, nombre, a, b, cap, fecha, a_ok, b_ok in SEED:
            e = Enlace(anillo=anillo, nombre=nombre, capacidad=cap,
                       origen_a=a, origen_b=b,
                       a_completo=bool(a_ok), b_completo=bool(b_ok),
                       creado=datetime.strptime(fecha, "%Y-%m-%d %H:%M:%S"))
            for n in range(1, cap + 1):
                e.hilos.append(Hilo(numero=n, descripcion_a="LIBRE" if a_ok else "",
                                    descripcion_b="LIBRE" if b_ok else ""))
            db.session.add(e)
        db.session.commit()
    yield app


@pytest.fixture
def cli(app):
    return app.test_client()


def login(cli, user="admin", pwd="admin123"):
    return cli.post("/login", data={"username": user, "password": pwd},
                    follow_redirects=True)


def _un_pendiente(app):
    with app.app_context():
        return Enlace.query.filter_by(b_completo=False, eliminado=False).first().id


# ---------- Seguridad ----------
def test_dashboard_exige_login(cli):
    r = cli.get("/", follow_redirects=False)
    assert r.status_code == 302 and "/login" in r.headers["Location"]

def test_login_correcto(cli):
    r = login(cli)
    assert r.status_code == 200 and b"Pendientes" in r.data

def test_login_incorrecto(cli):
    r = cli.post("/login", data={"username": "admin", "password": "mal"},
                 follow_redirects=True)
    assert "incorrectos".encode() in r.data

def test_lector_no_puede_editar(cli, app):
    login(cli, "lector", "lector123")
    eid = _un_pendiente(app)
    assert cli.get(f"/enlace/{eid}/editar").status_code == 403
    assert cli.post(f"/enlace/{eid}/eliminar").status_code == 403

def test_lector_si_puede_ver_y_exportar(cli, app):
    login(cli, "lector", "lector123")
    eid = _un_pendiente(app)
    assert cli.get("/").status_code == 200
    assert cli.get(f"/enlace/{eid}/export.xlsx").status_code == 200


# ---------- Validaciones de servidor ----------
def test_nombre_duplicado_rechazado(cli, app):
    login(cli)
    r = cli.post("/nuevo", data={"anillo": "ANILLO SUR",
                                 "nombre": "SAN JUAN - CHORRILLOS",
                                 "capacidad": "24"}, follow_redirects=True)
    assert "Ya existe".encode() in r.data

def test_capacidad_invalida_rechazada(cli):
    login(cli)
    r = cli.post("/nuevo", data={"anillo": "ANILLO SUR", "nombre": "X - Y",
                                 "capacidad": "77"}, follow_redirects=True)
    assert "Capacidad inv".encode() in r.data

def test_anillo_invalido_rechazado(cli):
    login(cli)
    r = cli.post("/nuevo", data={"anillo": "ANILLO FALSO", "nombre": "X - Z",
                                 "capacidad": "24"}, follow_redirects=True)
    assert "Anillo inv".encode() in r.data


# ---------- Flujo A/B ----------
def test_flujo_completo_a_y_b(cli, app):
    login(cli)
    cli.post("/nuevo", data={"anillo": "ANILLO SUR", "nombre": "PRUEBA - FLUJO",
                             "capacidad": "12"})
    with app.app_context():
        e = Enlace.query.filter_by(nombre="PRUEBA - FLUJO").first()
        eid = e.id
        assert e.estado == "INCOMPLETO" and len(e.hilos) == 12
    cli.post(f"/enlace/{eid}/completar/a",
             data={"origen": "PRUEBA", "h1": "ODF-5"})
    cli.post(f"/enlace/{eid}/completar/b",
             data={"origen": "FLUJO", "h1": "odf  5"})
    with app.app_context():
        e = db.session.get(Enlace, eid)
        assert e.estado == "VALIDADO"
        assert e.hilos[0].coincide          # 'ODF-5' == 'odf  5' (normalizado)

def test_normalizar():
    assert normalizar("ODF-5") == normalizar("odf  5") == normalizar("ODF_5")
    assert normalizar("LIBRE") != normalizar("OCUPADO")


# ---------- Capacidad dinámica ----------
def test_reducir_y_aumentar_capacidad(cli, app):
    login(cli)
    eid = _un_pendiente(app)
    cli.post(f"/enlace/{eid}/editar",
             data={"nombre_sel": "__actual__", "capacidad": "12",
                   "tipo_cable": "ADSS"})
    with app.app_context():
        e = db.session.get(Enlace, eid)
        anillo = e.anillo
        assert e.capacidad == 12 and len(e.hilos) == 12
    cli.post(f"/enlace/{eid}/editar",
             data={"nombre_sel": "__actual__", "anillo": anillo,
                   "capacidad": "48", "tipo_cable": "ADSS"})
    with app.app_context():
        e = db.session.get(Enlace, eid)
        assert e.capacidad == 48 and len(e.hilos) == 48


# ---------- Papelera ----------
def test_borrado_logico_y_restauracion(cli, app):
    login(cli)
    eid = _un_pendiente(app)
    cli.post(f"/enlace/{eid}/eliminar")
    with app.app_context():
        assert db.session.get(Enlace, eid).eliminado is True
    assert cli.get(f"/enlace/{eid}/ver").status_code == 404   # oculto
    r = cli.get("/papelera")
    assert r.status_code == 200
    cli.post(f"/enlace/{eid}/restaurar")
    with app.app_context():
        assert db.session.get(Enlace, eid).eliminado is False


# ---------- Auditoría ----------
def test_auditoria_registra_acciones(cli, app):
    login(cli)
    cli.post("/nuevo", data={"anillo": "ANILLO NORTE",
                             "nombre": "AUDIT - TEST", "capacidad": "24"})
    with app.app_context():
        logs = AuditLog.query.filter_by(accion="CREAR").all()
        assert any("AUDIT - TEST" in (l.detalle or "") for l in logs)
        assert logs[-1].usuario == "admin"


# ---------- Dashboard: filtros y paginación ----------
def test_buscador(cli):
    login(cli)
    r = cli.get("/?q=CHORRILLOS")
    assert b"CHORRILLOS" in r.data and b"AEROPUERTO - CALLAO" not in r.data

def test_filtro_anillo(cli):
    login(cli)
    r = cli.get("/?anillo=ANILLO+OESTE")
    assert b"ANILLO OESTE" in r.data and "ANILLO ESTE | AVIACION".encode() not in r.data

def test_paginacion(cli):
    login(cli)
    r1 = cli.get("/")
    assert "Página 1 de".encode() in r1.data
    r2 = cli.get("/?page=2")
    assert "Página 2 de".encode() in r2.data


# ---------- Reportes ----------
def test_reporte_avance(cli):
    login(cli)
    r = cli.get("/reporte-avance.xlsx")
    assert r.status_code == 200
    wb = openpyxl.load_workbook(io.BytesIO(r.data))
    assert wb.sheetnames == ["RESUMEN", "PENDIENTES", "VALIDADOS"]

def test_export_enlace_excluye_eliminados(cli, app):
    login(cli)
    eid = _un_pendiente(app)
    cli.post(f"/enlace/{eid}/eliminar")
    assert cli.get(f"/enlace/{eid}/export.xlsx").status_code == 404


# ---------- Salud ----------
def test_healthcheck_publico(cli):
    r = cli.get("/health")
    assert r.status_code == 200 and r.get_json()["status"] == "ok"
