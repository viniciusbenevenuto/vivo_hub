"""routes/api.py — API REST: CNs e SIP Router SP."""
import logging
from dataclasses import asdict

from flask import Blueprint, current_app, jsonify, request

from auth import login_required
from db import get_db
from operadoras import list_operadoras
from siprouter_sp import query_siprouter_sp

logger = logging.getLogger(__name__)
bp = Blueprint("api", __name__, url_prefix="/api")


# ---------------------------------------------------------------------------
# CNs
# ---------------------------------------------------------------------------
@bp.get("/cns")
@login_required
def api_cns():
    db   = get_db()
    rows = db.execute(
        "SELECT codigo, COALESCE(nome,'') AS nome, COALESCE(uf,'') AS uf "
        "FROM cns WHERE ativo=1 ORDER BY codigo ASC"
    ).fetchall()
    return jsonify([
        {"codigo": r["codigo"], "nome": r["nome"], "uf": r["uf"]}
        for r in rows
    ])


# ---------------------------------------------------------------------------
# OPERADORAS (Anexo 5 — ABR Telecom)
# ---------------------------------------------------------------------------
@bp.get("/operadoras")
@login_required
def api_operadoras():
    """Lista compacta de operadoras para o autopreenchimento do formulário.

    Retorno: [{rn1, nome, uf, eot}, ...] — uma linha por combinação
    operadora × UF do Anexo 5.
    """
    return jsonify(list_operadoras(get_db()))


# ---------------------------------------------------------------------------
# SBC VIVO POR CAPACIDADE (sbc/analyzer.py)
# ---------------------------------------------------------------------------
@bp.get("/sbc/sugestoes")
@login_required
def sbc_sugestoes():
    """SBCs do CN (ou da UF) ordenados por capacidade/saúde.

    Parâmetros: cn (preferencial) ou uf.
    Retorno: SBCSuggestionResponse — cada SBC traz saude
    (disponivel/moderado/critico), CAPS, tendência, score e recomendado.
    Sem planilha de medições, `sbcs` vem vazio com uma mensagem.
    """
    analyzer = current_app.config["SBC_ANALYZER"]
    cn = (request.args.get("cn") or "").strip()
    uf = (request.args.get("uf") or "").strip().upper()

    if cn:
        if not (cn.isdigit() and len(cn) <= 2):
            return jsonify({"error": "'cn' deve ter até 2 dígitos."}), 400
        resultado = analyzer.suggest_for_cn(cn)
    elif uf:
        resultado = analyzer.suggest_for_uf(uf)
    else:
        return jsonify({"error": "Informe 'cn' ou 'uf'."}), 400

    return jsonify(asdict(resultado))


# ---------------------------------------------------------------------------
# SIPROUTER SP
# ---------------------------------------------------------------------------
@bp.get("/siprouter/query")
@login_required
def siprouter_query():
    """
    Consulta a tabela siprouter_sp.

    Parâmetros obrigatórios:
        cn  — Código Nacional (inteiro, ex.: 11)
        rn1 — RN1 da operadora (ex.: 55131)

    Parâmetros opcionais:
        scm — 1 para filtrar por SCM, 0 (padrão)
        av  — 1 para filtrar por AV,  0 (padrão)

    Retorno JSON:
        { found, cn, rn1_consultado, resultados, dados_vivo, mensagem }
    """
    cn_raw = request.args.get("cn",  "").strip()
    rn1    = request.args.get("rn1", "").strip()
    scm    = request.args.get("scm", "0").strip() == "1"
    av     = request.args.get("av",  "0").strip() == "1"

    if not cn_raw or not rn1:
        return jsonify({"error": "Parâmetros 'cn' e 'rn1' são obrigatórios."}), 400

    try:
        cn = int(cn_raw)
    except ValueError:
        return jsonify({"error": f"'cn' deve ser inteiro. Recebido: {cn_raw!r}"}), 400

    result = query_siprouter_sp(get_db(), cn=cn, rn1=rn1, scm=scm, av=av)
    return jsonify(result)
