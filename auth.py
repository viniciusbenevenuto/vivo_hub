"""auth.py — Decoradores de autenticação e headers de segurança."""
from functools import wraps

from flask import Flask, Response, flash, redirect, session, url_for

# Endpoint da central de cada perfil (role) da aplicação.
ROLE_HOME_ENDPOINTS = {
    "atacado": "central.central_atacado",
    "engenharia": "central.central_engenharia",
}


def register_security_headers(app: Flask) -> None:
    @app.after_request
    def add_headers(resp: Response) -> Response:
        resp.headers.setdefault("X-Content-Type-Options", "nosniff")
        resp.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
        resp.headers.setdefault("X-XSS-Protection", "0")
        return resp


def login_required(view):
    """Exige usuário autenticado; redireciona para o login caso contrário."""
    @wraps(view)
    def wrapped(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("auth.login"))
        return view(*args, **kwargs)
    return wrapped


def role_required(required_role: str):
    """Exige que o usuário logado tenha o perfil (role) informado."""
    def decorator(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            if "user_id" not in session:
                return redirect(url_for("auth.login"))
            if session.get("role") != required_role:
                flash("Acesso negado para esta área.", "danger")
                home = ROLE_HOME_ENDPOINTS.get(session.get("role"))
                return redirect(url_for(home) if home else url_for("auth.login"))
            return view(*args, **kwargs)
        return wrapped
    return decorator


def admin_required(view):
    """Exige sessão de administrador (código validado em /admin_login)."""
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("is_admin"):
            flash("Área restrita ao administrador.", "danger")
            return redirect(url_for("auth.admin_login"))
        return view(*args, **kwargs)
    return wrapped
