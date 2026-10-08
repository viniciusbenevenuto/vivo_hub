"""routes/auth.py — Rotas de autenticação e registro de usuários."""
import logging
import sqlite3

from flask import (
    Blueprint, current_app, flash, redirect, render_template,
    request, session, url_for,
)
from werkzeug.security import check_password_hash, generate_password_hash

from auth import ROLE_HOME_ENDPOINTS, admin_required
from db import get_db

logger = logging.getLogger(__name__)
bp = Blueprint("auth", __name__)


@bp.get("/login")
def login():
    return render_template("login.html")


@bp.post("/login")
def login_post():
    email    = (request.form.get("email") or "").strip().lower()
    password = request.form.get("password", "")
    db       = get_db()
    user     = db.execute("SELECT * FROM users WHERE email = ?", (email,)).fetchone()

    if not user or not check_password_hash(user["password_hash"], password):
        flash("Credenciais inválidas.", "danger")
        return redirect(url_for("auth.login"))

    session.clear()
    session["user_id"] = user["id"]
    session["email"]   = user["email"]
    session["role"]    = user["role"]
    return redirect(url_for(ROLE_HOME_ENDPOINTS[user["role"]]))


@bp.get("/logout")
def logout():
    session.clear()
    flash("Sessão encerrada.", "info")
    return redirect(url_for("auth.login"))


@bp.get("/admin_login")
def admin_login():
    return render_template("admin_login.html")


@bp.post("/admin_login")
def admin_login_post():
    if request.form.get("code", "") == current_app.config["ADMIN_CODE"]:
        session["is_admin"] = True
        flash("Login de administrador efetuado.", "success")
        return redirect(url_for("auth.register"))
    flash("Código de administrador inválido.", "danger")
    return redirect(url_for("auth.admin_login"))


@bp.get("/register")
@admin_required
def register():
    return render_template("register.html")


@bp.post("/register")
@admin_required
def register_post():
    email    = (request.form.get("email") or "").strip().lower()
    password = request.form.get("password", "")
    role     = (request.form.get("role") or "").strip().lower()

    if not email or not password or role not in ("engenharia", "atacado"):
        flash("Preencha todos os campos corretamente.", "danger")
        return redirect(url_for("auth.register"))

    db = get_db()
    try:
        db.execute(
            "INSERT INTO users (email, password_hash, role) VALUES (?, ?, ?)",
            (email, generate_password_hash(password), role),
        )
        db.commit()
        flash("Usuário criado com sucesso.", "success")
    except sqlite3.IntegrityError:
        flash("Este e-mail já está cadastrado.", "warning")

    return redirect(url_for("auth.register"))

