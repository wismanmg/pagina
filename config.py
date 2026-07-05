# -*- coding: utf-8 -*-
"""Configuración basada en variables de entorno (nunca valores fijos)."""
import os


class Config:
    SECRET_KEY = os.environ.get("SECRET_KEY", "cambia-esta-clave-en-produccion")
    # PostgreSQL en producción (Render la inyecta como DATABASE_URL);
    # SQLite como respaldo para desarrollo local.
    _db = os.environ.get("DATABASE_URL", "sqlite:///enlaces.db")
    # Render entrega postgres:// pero SQLAlchemy 2 exige postgresql://
    SQLALCHEMY_DATABASE_URI = _db.replace("postgres://", "postgresql://", 1)
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    DEBUG = os.environ.get("DEBUG", "0") == "1"
    ITEMS_POR_PAGINA = int(os.environ.get("ITEMS_POR_PAGINA", "10"))


class TestConfig(Config):
    TESTING = True
    SQLALCHEMY_DATABASE_URI = "sqlite://"     # en memoria
    WTF_CSRF_ENABLED = False                  # los tests postean sin token
