"""routes/central.py — Rotas das centrais (landing pages por role)."""
from flask import Blueprint, redirect, render_template, session, url_for

from auth import ROLE_HOME_ENDPOINTS, login_required, role_required

bp = Blueprint("central", __name__)


@bp.get("/")
@login_required
def index():
    home = ROLE_HOME_ENDPOINTS.get(session.get("role"))
    return redirect(url_for(home) if home else url_for("auth.login"))


@bp.get("/central_atacado")
@login_required
@role_required("atacado")
def central_atacado():
    return render_template("central_atacado.html")


@bp.get("/central_engenharia")
@login_required
@role_required("engenharia")
def central_engenharia():
    return render_template("central_engenharia.html")
