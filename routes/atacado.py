"""routes/atacado.py — CRUD de formulários do perfil Atacado + versionamento."""
import logging

from flask import (
    Blueprint, abort, flash, redirect, render_template,
    request, session, url_for,
)

from auth import login_required, role_required
from config import BOOLEAN_FIELDS, ENG_ITX_NAME, MAX_TABLE_ROWS, TEXT_FIELDS
from db import get_db
from forms import extract_form_payload, validate_cn_uf_rows, validate_table_rows
from helpers import display_name_from_email, group_forms_by_operadora, utc_now_str
from operadoras import rn1_is_valid, rn1_nome_match
from queries import build_list_query, get_status_counters

logger = logging.getLogger(__name__)
bp = Blueprint("atacado", __name__, url_prefix="/atacado_formularios")

# Colunas JSON gravadas junto com os campos de texto/booleanos.
_JSON_COLS = ["escopo_flags_json", "dados_vivo_json", "dados_operadora_json"]
_FORM_COLS = list(TEXT_FIELDS) + list(BOOLEAN_FIELDS) + _JSON_COLS


def _apply_fixed_responsaveis(payload: dict) -> None:
    """Campos de responsáveis não são editáveis: valores definidos no servidor.

    - Gestão ITX Atacado = usuário logado (responsavel_atacado/responsavel_vivo)
    - Eng de ITX e Aprovado por = ENG_ITX_NAME (definido em config.py)
    """
    gestao = display_name_from_email(session.get("email", ""))
    payload["responsavel_atacado"]    = gestao
    payload["responsavel_vivo"]       = gestao
    payload["responsavel_engenharia"] = ENG_ITX_NAME
    payload["aprovado_por"]           = ENG_ITX_NAME


def _validate_payload(db, payload: dict) -> list[str]:
    """Valida campos restritos contra o Anexo 5. Retorna lista de erros.

    Em rascunho, campos vazios são tolerados; valores preenchidos nunca
    podem estar fora da tabela de operadoras.
    """
    errors: list[str] = []
    strict = payload.get("status") == "enviado"

    rn1  = (payload.get("rn1") or "").strip()
    nome = (payload.get("nome_operadora") or "").strip()
    tmr  = (payload.get("tmr") or "").strip()

    if rn1:
        if not rn1_is_valid(rn1):
            errors.append("RN1 deve conter apenas números (máx. 5 dígitos).")
        elif not db.execute(
            "SELECT 1 FROM operadoras WHERE rn1 = ? LIMIT 1", (rn1,)
        ).fetchone():
            errors.append(f"RN1 {rn1} não consta no Anexo 5.")
    elif strict:
        errors.append("Informe o RN1 da operadora.")

    if nome:
        if not db.execute(
            "SELECT 1 FROM operadoras WHERE nome = ? LIMIT 1", (nome,)
        ).fetchone():
            errors.append(f'Operadora "{nome}" não consta no Anexo 5.')
        elif rn1 and rn1_is_valid(rn1) and not rn1_nome_match(db, rn1, nome):
            errors.append(f'O RN1 {rn1} não pertence à operadora "{nome}".')
    elif strict:
        errors.append("Informe o nome da operadora.")

    if tmr and not (tmr.isdigit() and len(tmr) <= 2):
        errors.append("TMR deve conter apenas números (máx. 2 dígitos).")

    asn = (payload.get("asn") or "").strip()
    if asn and not (asn.isdigit() and 5 <= len(asn) <= 6):
        errors.append("ASN deve conter apenas números (5 a 6 dígitos).")

    errors += validate_cn_uf_rows(payload.get("dados_operadora_json", "[]"), "Dados Operadora")
    errors += validate_cn_uf_rows(payload.get("dados_vivo_json", "[]"), "Dados VIVO")

    return errors


def _get_own_form(db, form_id: int):
    """Retorna o formulário se pertencer ao usuário logado; senão 404."""
    form = db.execute(
        "SELECT * FROM atacado_forms WHERE id = ? AND owner_id = ?",
        (form_id, session["user_id"]),
    ).fetchone()
    if not form:
        abort(404)
    return form


# ---------------------------------------------------------------------------
# LISTAGEM — agrupada por operadora, com versões
# ---------------------------------------------------------------------------
@bp.get("")
@login_required
@role_required("atacado")
def form_list():
    db     = get_db()
    uid    = session["user_id"]
    q      = (request.args.get("q") or "").strip()
    status = (request.args.get("status") or "").strip()

    sql, params = build_list_query(
        "SELECT * FROM atacado_forms",
        search_term=q or None,
        status_filter=status or None,
        owner_id=uid,
        sort_key="operadora_version",
    )
    forms    = db.execute(sql, params).fetchall()
    grupos   = group_forms_by_operadora(forms)
    counters = get_status_counters(db, "WHERE owner_id = ?", (uid,))

    return render_template(
        "atacado_formularios.html",
        grupos=grupos,
        counters=counters,
        q=q,
        status=status,
    )


# ---------------------------------------------------------------------------
# CRIAR
# ---------------------------------------------------------------------------
@bp.get("/new")
@login_required
@role_required("atacado")
def form_new():
    return render_template("formulario_atacado.html", form=None)


@bp.post("/new")
@login_required
@role_required("atacado")
def form_create():
    payload = extract_form_payload()
    payload["owner_id"]               = session["user_id"]
    payload["status"]                 = (payload.get("status") or "rascunho").lower()
    payload["engenharia_params_json"] = "{}"
    payload["version"]                = 1
    payload["parent_id"]              = None
    truncated = validate_table_rows(payload)
    _apply_fixed_responsaveis(payload)

    db = get_db()
    errors = _validate_payload(db, payload)
    if errors:
        for err in errors:
            flash(err, "warning")
        return redirect(url_for("atacado.form_new"))

    cols = ["owner_id", "version", "parent_id"] + _FORM_COLS + ["engenharia_params_json"]
    db.execute(
        f"INSERT INTO atacado_forms ({','.join(cols)}) "
        f"VALUES ({','.join(['?'] * len(cols))})",
        [payload.get(c) for c in cols],
    )
    db.commit()

    msg = "Formulário criado."
    if truncated:
        msg += f" Tabelas limitadas a {MAX_TABLE_ROWS} linhas."
    flash(msg, "success")
    return redirect(url_for("atacado.form_list"))


# ---------------------------------------------------------------------------
# NOVA VERSÃO
# ---------------------------------------------------------------------------
@bp.post("/<int:form_id>/new_version")
@login_required
@role_required("atacado")
def form_new_version(form_id: int):
    db   = get_db()
    orig = _get_own_form(db, form_id)

    root_id = orig["parent_id"] or orig["id"]
    max_ver = db.execute(
        "SELECT MAX(version) FROM atacado_forms "
        "WHERE owner_id = ? AND (id = ? OR parent_id = ?)",
        (session["user_id"], root_id, root_id),
    ).fetchone()[0] or 1
    next_ver = max_ver + 1

    row_dict = dict(orig)
    db.execute(
        f"INSERT INTO atacado_forms "
        f"(owner_id, version, parent_id, status, {','.join(_FORM_COLS)}, engenharia_params_json) "
        f"VALUES (?, ?, ?, 'rascunho', {','.join(['?'] * len(_FORM_COLS))}, '{{}}')",
        [session["user_id"], next_ver, root_id] + [row_dict.get(c) for c in _FORM_COLS],
    )
    db.commit()

    flash(f"Versão {next_ver} criada a partir do formulário #{form_id}.", "success")
    return redirect(url_for("atacado.form_list"))


# ---------------------------------------------------------------------------
# EDITAR
# ---------------------------------------------------------------------------
@bp.get("/<int:form_id>")
@login_required
@role_required("atacado")
def form_edit(form_id: int):
    form = _get_own_form(get_db(), form_id)
    return render_template("formulario_atacado.html", form=form)


@bp.post("/<int:form_id>")
@login_required
@role_required("atacado")
def form_update(form_id: int):
    db = get_db()
    _get_own_form(db, form_id)

    payload               = extract_form_payload()
    payload["status"]     = (payload.get("status") or "rascunho").lower()
    payload["updated_at"] = utc_now_str()
    truncated             = validate_table_rows(payload)
    _apply_fixed_responsaveis(payload)

    errors = _validate_payload(db, payload)
    if errors:
        for err in errors:
            flash(err, "warning")
        return redirect(url_for("atacado.form_edit", form_id=form_id))

    assignments = ", ".join([f"{col} = ?" for col in _FORM_COLS] + ["updated_at = ?"])
    params = (
        [payload.get(col) for col in _FORM_COLS]
        + [payload["updated_at"], form_id, session["user_id"]]
    )
    db.execute(
        f"UPDATE atacado_forms SET {assignments} WHERE id = ? AND owner_id = ?",
        params,
    )
    db.commit()

    msg = "Formulário atualizado."
    if truncated:
        msg += f" Tabelas limitadas a {MAX_TABLE_ROWS} linhas."
    flash(msg, "success")
    return redirect(url_for("atacado.form_list"))


# ---------------------------------------------------------------------------
# DELETAR
# ---------------------------------------------------------------------------
@bp.post("/<int:form_id>/delete")
@login_required
@role_required("atacado")
def form_delete(form_id: int):
    db = get_db()
    db.execute(
        "DELETE FROM atacado_forms WHERE id = ? AND owner_id = ?",
        (form_id, session["user_id"]),
    )
    db.commit()
    flash("Formulário excluído.", "info")
    return redirect(url_for("atacado.form_list"))
