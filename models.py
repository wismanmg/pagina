# -*- coding: utf-8 -*-
import re
import uuid
from datetime import datetime

from flask_login import UserMixin, current_user
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()


def normalizar(texto):
    """Normaliza descripciones para comparar A vs B:
    mayúsculas, guiones/underscores como espacio, espacios colapsados.
    Así 'ODF 5' == 'ODF-5' == 'odf  5'."""
    t = (texto or "").upper()
    t = re.sub(r"[-_/]+", " ", t)
    return " ".join(t.split())


class Usuario(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(60), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    rol = db.Column(db.String(10), default="lector")   # 'editor' | 'lector'

    def set_password(self, pwd):
        self.password_hash = generate_password_hash(pwd)

    def check_password(self, pwd):
        return check_password_hash(self.password_hash, pwd)

    @property
    def es_editor(self):
        return self.rol == "editor"


class Sitio(db.Model):
    """Catálogo normalizado de sitios (corrige tipeos históricos)."""
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(80), unique=True, nullable=False)


class Enlace(db.Model):
    id = db.Column(db.String(36), primary_key=True,
                   default=lambda: str(uuid.uuid4()))
    anillo = db.Column(db.String(30), nullable=False, index=True)
    nombre = db.Column(db.String(120), nullable=False)
    capacidad = db.Column(db.Integer, nullable=False, default=24)
    tipo_cable = db.Column(db.String(60), default="CABLE DE ENLACE SM")
    longitud = db.Column(db.String(40), default="")
    origen_a = db.Column(db.String(80), default="")
    sala_a = db.Column(db.String(80), default="")
    rack_a = db.Column(db.String(40), default="")
    posicion_a = db.Column(db.String(40), default="")
    origen_b = db.Column(db.String(80), default="")
    sala_b = db.Column(db.String(80), default="")
    rack_b = db.Column(db.String(40), default="")
    posicion_b = db.Column(db.String(40), default="")
    a_completo = db.Column(db.Boolean, default=False)
    b_completo = db.Column(db.Boolean, default=False)
    eliminado = db.Column(db.Boolean, default=False, index=True)  # borrado lógico
    creado = db.Column(db.DateTime, default=datetime.utcnow)
    modificado_por = db.Column(db.String(60), default="")
    modificado_en = db.Column(db.DateTime)
    hilos = db.relationship("Hilo", backref="enlace",
                            cascade="all, delete-orphan",
                            order_by="Hilo.numero", lazy="select")

    CAPACIDADES_VALIDAS = (12, 24, 48, 96, 144)

    @property
    def estado(self):
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

    @property
    def dias(self):
        return (datetime.utcnow() - self.creado).days

    @property
    def semaforo(self):
        """verde <30d, ambar 30-89d, rojo >=90d"""
        d = self.dias
        return "verde" if d < 30 else ("ambar" if d < 90 else "rojo")

    def tocar(self):
        """Registra quién y cuándo modificó."""
        self.modificado_en = datetime.utcnow()
        try:
            self.modificado_por = current_user.username
        except Exception:
            self.modificado_por = "sistema"


class Hilo(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    enlace_id = db.Column(db.String(36), db.ForeignKey("enlace.id"), index=True)
    numero = db.Column(db.Integer)
    descripcion_a = db.Column(db.Text, default="")
    descripcion_b = db.Column(db.Text, default="")

    @property
    def coincide(self):
        return normalizar(self.descripcion_a) == normalizar(self.descripcion_b)


class AuditLog(db.Model):
    """Historial: quién hizo qué y cuándo."""
    id = db.Column(db.Integer, primary_key=True)
    fecha = db.Column(db.DateTime, default=datetime.utcnow)
    usuario = db.Column(db.String(60))
    accion = db.Column(db.String(30))       # CREAR|EDITAR|COMPLETAR_A|...
    enlace_id = db.Column(db.String(36))
    detalle = db.Column(db.String(300))


def auditar(accion, enlace_id="", detalle=""):
    try:
        user = current_user.username
    except Exception:
        user = "sistema"
    db.session.add(AuditLog(usuario=user, accion=accion,
                            enlace_id=enlace_id, detalle=detalle[:300]))
