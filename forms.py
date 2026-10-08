"""forms.py — Extração e validação de dados de formulários Flask."""
import json
import logging

from flask import request

from config import (
    ALLOWED_SCOPE_FLAGS,
    BOOLEAN_FIELDS,
    CN_METADATA,
    JSON_FIELDS,
    MAX_TABLE_ROWS,
    TEXT_FIELDS,
)
from helpers import parse_bool_field, truncate_json_list

logger = logging.getLogger(__name__)


# =============================================================================
# ESCOPO / FLAGS
# =============================================================================
def extract_scope_flags_from_request() -> str:
    """
    Lê flags de escopo do request.

    Tentativa 1: checkboxes com name="escopo_flags".
    Tentativa 2: hidden input "escopo_flags_json" atualizado pelo JavaScript.
    """
    raw = request.form.getlist("escopo_flags") or []

    if not raw:
        json_raw = request.form.get("escopo_flags_json", "[]")
        try:
            parsed = json.loads(json_raw) if isinstance(json_raw, str) else []
            if isinstance(parsed, list):
                raw = [str(x).strip() for x in parsed if str(x).strip()]
        except (json.JSONDecodeError, TypeError):
            raw = []

    valid = [flag for flag in raw if flag in ALLOWED_SCOPE_FLAGS]
    return json.dumps(list(dict.fromkeys(valid)), ensure_ascii=False)


def parse_json_dict(raw) -> str:
    """Normaliza um JSON de objeto; retorna '{}' se inválido."""
    if not raw:
        return "{}"
    try:
        parsed = json.loads(raw) if isinstance(raw, str) else raw
        if isinstance(parsed, dict):
            return json.dumps(parsed, ensure_ascii=False)
    except (json.JSONDecodeError, TypeError):
        pass
    return "{}"


# =============================================================================
# PAYLOAD COMPLETO
# =============================================================================
def extract_form_payload() -> dict:
    """Extrai todos os campos do formulário de PTI a partir do request."""
    payload: dict = {}
    for field in TEXT_FIELDS:
        payload[field] = (request.form.get(field) or "").strip()
    for field in BOOLEAN_FIELDS:
        payload[field] = parse_bool_field(request.form.get(field))
    payload["escopo_flags_json"] = extract_scope_flags_from_request()
    for field in JSON_FIELDS:
        if field == "escopo_flags_json":
            continue
        raw = request.form.get(field, "")
        payload[field] = (
            parse_json_dict(raw)
            if field == "engenharia_params_json"
            else truncate_json_list(raw, "[]")
        )
    return payload


def validate_cn_uf_rows(rows_json, label: str) -> list[str]:
    """Verifica que o CN de cada linha pertence à UF informada (Anexo/CN_METADATA).

    Retorna a lista de erros; linhas sem CN ou sem UF são ignoradas.
    """
    errors: list[str] = []
    try:
        rows = json.loads(rows_json) if isinstance(rows_json, str) else rows_json
    except (json.JSONDecodeError, TypeError):
        return errors
    if not isinstance(rows, list):
        return errors
    for i, row in enumerate(rows, 1):
        if not isinstance(row, dict):
            continue
        cn = str(row.get("cn", "")).strip()
        uf = str(row.get("uf", "")).strip().upper()
        if not cn or not uf:
            continue
        meta = CN_METADATA.get(cn.zfill(2))
        if meta and meta[1] != uf:
            errors.append(
                f"{label} (linha {i}): CN {cn} pertence a {meta[1]}, não a {uf}."
            )
    return errors


def validate_table_rows(payload: dict) -> bool:
    """Trunca tabelas que excedam MAX_TABLE_ROWS. Retorna True se truncou."""
    truncated = False
    for key in ("dados_vivo_json", "dados_operadora_json"):
        try:
            rows = json.loads(payload[key])
            if isinstance(rows, list) and len(rows) > MAX_TABLE_ROWS:
                payload[key] = json.dumps(rows[:MAX_TABLE_ROWS], ensure_ascii=False)
                truncated = True
        except (json.JSONDecodeError, TypeError, KeyError):
            payload[key] = "[]"
    return truncated
