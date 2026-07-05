# -*- coding: utf-8 -*-
"""Autenticación (Flask-Login) y control de rol editor/lector."""
from functools import wraps

from flask import Blueprint, render_template, request, redirect, url_for, flash, abort
from flask_login import LoginManager, login_user, logout_user, login_required, current_user

from models import db, Usuario

auth_bp = Blueprint("auth", __name__)
login_manager = LoginManager()
login_manager.login_view = "auth.login"
login_manager.login_message = "Debes iniciar sesión para continuar."


@login_manager.user_loader
def load_user(uid):
    return db.session.get(Usuario, int(uid))


def solo_editor(f):
    """Decorator: la vista exige rol editor (los lectores reciben 403)."""
    @wraps(f)
    def wrapper(*a, **kw):
        if not current_user.is_authenticated:
            return login_manager.unauthorized()
        if not current_user.es_editor:
            abort(403)
        return f(*a, **kw)
    return wrapper


@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        u = Usuario.query.filter_by(
            username=request.form.get("username", "").strip().lower()).first()
        if u and u.check_password(request.form.get("password", "")):
            login_user(u)
            destino = request.args.get("next") or url_for("main.dashboard")
            return redirect(destino)
        flash("Usuario o contraseña incorrectos.", "error")
    return render_template("login.html")


@auth_bp.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("auth.login"))
