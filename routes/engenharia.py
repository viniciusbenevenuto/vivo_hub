"""routes/engenharia.py — Formulários, validação de engenharia e geração de Excel."""
import json
import logging
import os
import tempfile
from io import BytesIO

from flask import (
    Blueprint, abort, flash, redirect, render_template,
    request, send_file, session, url_for,
)

from auth import login_required, role_required
from config import ENG_ITX_NAME
from db import get_db
from forms import parse_json_dict, truncate_json_list, validate_cn_uf_rows
from helpers import group_forms_by_operadora, safe_filename, utc_now_str
from params_import import MAX_UPLOAD_BYTES, ParamsImportError, ler_parametros
from queries import build_list_query, get_status_counters

logger = logging.getLogger(__name__)
bp = Blueprint("engenharia", __name__)

_XLSX_MIMETYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ---------------------------------------------------------------------------
# LISTAGEM
# ---------------------------------------------------------------------------
@bp.get("/engenharia_formularios")
@login_required
@role_required("engenharia")
def form_list():
    db     = get_db()
    q      = (request.args.get("q") or "").strip()
    status = (request.args.get("status") or "").strip()

    sql, params = build_list_query(
        "SELECT * FROM atacado_forms",
        search_term=q or None,
        status_filter=status or None,
        sort_key="operadora_version",
    )
    forms    = db.execute(sql, params).fetchall()
    grupos   = group_forms_by_operadora(forms)
    counters = get_status_counters(db)

    has_download = bool(session.get("_excel_tmp"))
    ipam_ok      = session.pop("_ipam_ok",  None)
    ipam_msg     = session.pop("_ipam_msg", "")

    return render_template(
        "engenharia_formularios.html",
        grupos=grupos,
        counters=counters,
        q=q,
        status=status,
        download_url=url_for("engenharia.excel_download") if has_download else "",
        ipam_ok=ipam_ok,
        ipam_msg=ipam_msg,
    )


# ---------------------------------------------------------------------------
# VISUALIZAR / VALIDAR
# ---------------------------------------------------------------------------
@bp.route("/engenharia_formularios/<int:form_id>", methods=["GET", "POST"])
@login_required
@role_required("engenharia")
def form_view(form_id: int):
    db   = get_db()
    form = db.execute(
        "SELECT * FROM atacado_forms WHERE id = ?", (form_id,)
    ).fetchone()
    if not form:
        abort(404)

    if request.method == "POST":
        # Eng de ITX é fixo, definido em config.ENG_ITX_NAME
        resp_eng  = ENG_ITX_NAME
        eng_json  = parse_json_dict(request.form.get("engenharia_params_json", "") or "{}")
        vivo_json = truncate_json_list(request.form.get("dados_vivo_json", "[]"), "[]")
        op_json   = truncate_json_list(request.form.get("dados_operadora_json", "[]"), "[]")

        cn_errors = (validate_cn_uf_rows(vivo_json, "Dados VIVO")
                     + validate_cn_uf_rows(op_json, "Dados Operadora"))
        if cn_errors:
            for err in cn_errors:
                flash(err, "warning")
            return redirect(url_for("engenharia.form_view", form_id=form_id))

        # Banda / Canais / CAPS por tipo de tráfego (seção 8)
        metricas = parse_json_dict(
            request.form.get("trafego_metricas_json", "") or "{}"
        )

        db.execute(
            "UPDATE atacado_forms "
            "SET engenharia_params_json=?, responsavel_engenharia=?, "
            "dados_vivo_json=?, dados_operadora_json=?, "
            "trafego_metricas_json=?, updated_at=? "
            "WHERE id=?",
            (eng_json, resp_eng, vivo_json, op_json, metricas,
             utc_now_str(), form_id),
        )
        db.commit()
        flash("Validação da Engenharia salva.", "success")
        return redirect(url_for("engenharia.form_list"))

    # Respostas já importadas do PTI devolvido pela operadora (se houver)
    op_params = None
    if "operadora_params_json" in form.keys() and form["operadora_params_json"]:
        try:
            op_params = json.loads(form["operadora_params_json"])
        except (json.JSONDecodeError, TypeError):
            logger.warning("operadora_params_json inválido no PTI #%s", form_id)

    return render_template(
        "formulario_atacado.html", form=form, readonly=True, op_params=op_params
    )


def _set_form_status(form_id: int, new_status: str) -> str:
    """Atualiza o status do PTI e retorna o nome da operadora (ou 404)."""
    db = get_db()
    form = db.execute(
        "SELECT id, nome_operadora FROM atacado_forms WHERE id = ?", (form_id,)
    ).fetchone()
    if not form:
        abort(404)
    db.execute(
        "UPDATE atacado_forms SET status=?, updated_at=? WHERE id=?",
        (new_status, utc_now_str(), form_id),
    )
    db.commit()
    return form["nome_operadora"] or ""


@bp.post("/engenharia_formularios/<int:form_id>/reprovar")
@login_required
@role_required("engenharia")
def reprovar(form_id: int):
    nome = _set_form_status(form_id, "reprovado")
    flash(f"PTI #{form_id} — {nome} reprovado.", "warning")
    return redirect(url_for("engenharia.form_list"))


@bp.post("/engenharia_formularios/<int:form_id>/validar")
@login_required
@role_required("engenharia")
def validar(form_id: int):
    nome = _set_form_status(form_id, "aprovado")
    flash(f"PTI #{form_id} — {nome} aprovado.", "success")
    return redirect(url_for("engenharia.form_list"))


@bp.post("/engenharia_formularios/<int:form_id>/importar_parametros")
@login_required
@role_required("engenharia")
def importar_parametros(form_id: int):
    """Importa o PTI devolvido pela operadora com a aba de parâmetros preenchida."""
    db = get_db()
    if not db.execute(
        "SELECT id FROM atacado_forms WHERE id = ?", (form_id,)
    ).fetchone():
        abort(404)

    arquivo = request.files.get("arquivo_pti")
    if not arquivo or not arquivo.filename:
        flash("Selecione o arquivo .xlsx devolvido pela operadora.", "warning")
        return redirect(url_for("engenharia.form_view", form_id=form_id))

    if not arquivo.filename.lower().endswith((".xlsx", ".xlsm")):
        flash("Formato inválido. Envie o arquivo .xlsx do PTI.", "warning")
        return redirect(url_for("engenharia.form_view", form_id=form_id))

    arquivo.stream.seek(0, os.SEEK_END)
    tamanho = arquivo.stream.tell()
    arquivo.stream.seek(0)
    if tamanho > MAX_UPLOAD_BYTES:
        flash("Arquivo muito grande (máximo 10 MB).", "warning")
        return redirect(url_for("engenharia.form_view", form_id=form_id))

    try:
        resultado = ler_parametros(arquivo.stream)
    except ParamsImportError as exc:
        flash(str(exc), "danger")
        return redirect(url_for("engenharia.form_view", form_id=form_id))

    db.execute(
        "UPDATE atacado_forms "
        "SET operadora_params_json=?, operadora_params_em=?, updated_at=? "
        "WHERE id=?",
        (json.dumps(resultado, ensure_ascii=False), utc_now_str(),
         utc_now_str(), form_id),
    )
    db.commit()

    flash(
        f"Parâmetros da operadora importados: {resultado['respondidos']} "
        f"de {resultado['total']} respondidos.",
        "success",
    )
    return redirect(url_for("engenharia.form_view", form_id=form_id))


def _get_all_versions(db, form) -> list:
    """Retorna todas as versões relacionadas ao PTI, ordenadas por versão."""
    rows: list = []
    visited: set[int] = set()

    # Subir até a raiz pelo parent_id
    cur = form
    while cur and cur["id"] not in visited:
        visited.add(cur["id"])
        rows.append(cur)
        if not cur["parent_id"]:
            break
        cur = db.execute(
            "SELECT * FROM atacado_forms WHERE id = ?", (cur["parent_id"],)
        ).fetchone()

    # Descer: buscar todos os filhos/netos a partir da raiz
    def _collect_children(parent_id: int) -> None:
        for child in db.execute(
            "SELECT * FROM atacado_forms WHERE parent_id = ?", (parent_id,)
        ).fetchall():
            if child["id"] not in visited:
                visited.add(child["id"])
                rows.append(child)
                _collect_children(child["id"])

    root_id = rows[-1]["id"] if rows else form["id"]
    _collect_children(root_id)

    rows.sort(key=lambda r: int(r["version"] or 1))
    return rows


# ---------------------------------------------------------------------------
# DOWNLOAD DO EXCEL (intermediário para permitir flash antes do download)
# ---------------------------------------------------------------------------
@bp.get("/excel_download")
@login_required
def excel_download():
    tmp_path = session.pop("_excel_tmp",  None)
    fname    = session.pop("_excel_name", "PTI.xlsx")
    if not tmp_path or not os.path.exists(tmp_path):
        flash("Arquivo não encontrado. Gere o Excel novamente.", "warning")
        return redirect(url_for("engenharia.form_list"))

    with open(tmp_path, "rb") as fh:
        data = fh.read()
    os.unlink(tmp_path)

    return send_file(
        BytesIO(data),
        mimetype=_XLSX_MIMETYPE,
        as_attachment=True,
        download_name=fname,
    )


# ---------------------------------------------------------------------------
# EXPORTAR EXCEL COMPLETO (PTI)
# ---------------------------------------------------------------------------
@bp.get("/formularios/<int:form_id>/excel_index")
@login_required
def exportar_excel(form_id: int):
    try:
        from excel.builder import PTIWorkbookBuilder
    except ImportError:
        flash("openpyxl não instalado.", "warning")
        return redirect(url_for("central.index"))

    db   = get_db()
    form = db.execute(
        "SELECT * FROM atacado_forms WHERE id = ?", (form_id,)
    ).fetchone()
    if not form:
        abort(404)
    if session.get("role") == "atacado" and form["owner_id"] != session.get("user_id"):
        abort(403)

    builder = PTIWorkbookBuilder(form, all_versions=_get_all_versions(db, form))
    wb      = builder.build()

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    tmp.close()

    nome  = safe_filename(form["nome_operadora"] or "Operadora")
    ver   = form["version"] if "version" in form.keys() else 1
    fname = f"PTI_{nome}_v{ver}_ID{form['id']}.xlsx"

    session["_excel_tmp"]  = tmp.name
    session["_excel_name"] = fname

    # Resultado IPAM:
    #   nenhuma tentativa → não aplicável; falhas → erro com detalhes;
    #   reservas feitas → sucesso; tentou sem resultado → alerta.
    if builder.ipam_tentativas == 0:
        session["_ipam_ok"], session["_ipam_msg"] = None, ""
    elif builder.ipam_falhas:
        session["_ipam_ok"], session["_ipam_msg"] = False, "; ".join(builder.ipam_falhas)
    elif builder.ipam_reservas:
        session["_ipam_ok"], session["_ipam_msg"] = True, ""
    else:
        session["_ipam_ok"] = False
        session["_ipam_msg"] = "Reserva não concluída — verifique o pool no phpIPAM"

    return redirect(url_for("engenharia.form_list"))
