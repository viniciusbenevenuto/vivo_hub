"""helpers.py — Funções utilitárias, context processors e template filters."""
import json
import logging
from datetime import datetime, timezone
from typing import Optional

from flask import Flask, session

from config import ENG_ITX_NAME, MAX_TABLE_ROWS
from db import get_db
from params_programacao import PARAMS_SECOES, TOTAL_RESPONDIVEIS, param_key
from traffic_types import TRAFFIC_TYPES

logger = logging.getLogger(__name__)


# =============================================================================
# UTILITÁRIOS GERAIS
# =============================================================================
def row_get(row, key: str, default=""):
    """Acessa uma linha de DB com fallback seguro."""
    try:
        value = row[key]
        return default if value is None else value
    except (KeyError, IndexError):
        return default


def safe_filename(name: str) -> str:
    """Remove caracteres inválidos de um nome de arquivo."""
    name = (name or "").strip().replace(" ", "_")
    for ch in ("..", "/", "\\", ":", "*", "?", '"', "<", ">", "|"):
        name = name.replace(ch, "_")
    return name or "Operadora"


def utc_now_str() -> str:
    """Timestamp UTC no formato ISO usado nas colunas updated_at."""
    return (
        datetime.now(timezone.utc)
        .replace(tzinfo=None)
        .isoformat(timespec="seconds")
    )


def parse_db_datetime(value) -> Optional[datetime]:
    """Converte string ISO ou similar para datetime."""
    if not value:
        return None
    try:
        return datetime.fromisoformat(str(value).replace("T", " "))
    except (ValueError, TypeError):
        return None


def truncate_json_list(raw, default: str = "[]") -> str:
    """Limita uma lista JSON a MAX_TABLE_ROWS itens."""
    if not raw:
        return default
    try:
        parsed = json.loads(raw) if isinstance(raw, str) else raw
        if not isinstance(parsed, list):
            return default
        return json.dumps(parsed[:MAX_TABLE_ROWS], ensure_ascii=False)
    except (json.JSONDecodeError, TypeError):
        return default


def parse_bool_field(value) -> int:
    """Converte valor de checkbox/select em 0/1."""
    return 1 if str(value).lower() in ("on", "1", "true", "sim") else 0


def display_name_from_email(email: str) -> str:
    """'joao.silva@vivo.com' → 'Joao Silva'."""
    local = (email or "").split("@")[0]
    return local.replace(".", " ").replace("_", " ").title()


def group_forms_by_operadora(forms) -> dict[str, list]:
    """Agrupa linhas de atacado_forms pelo nome da operadora."""
    grupos: dict[str, list] = {}
    for form in forms:
        operadora = (form["nome_operadora"] or "Sem operadora").strip()
        grupos.setdefault(operadora, []).append(form)
    return grupos


# =============================================================================
# CONTEXT PROCESSORS E TEMPLATE FILTERS
# =============================================================================
def register_context_processors(app: Flask) -> None:
    @app.context_processor
    def inject_cn_codes():
        db = get_db()
        rows = db.execute(
            "SELECT codigo, COALESCE(nome,'') AS nome, COALESCE(uf,'') AS uf "
            "FROM cns WHERE ativo = 1 ORDER BY codigo ASC"
        ).fetchall()
        return {
            "CN_CODES": [r["codigo"] for r in rows],
            "CN_FULL": [
                {"codigo": r["codigo"], "nome": r["nome"], "uf": r["uf"]}
                for r in rows
            ],
            "ENG_ITX_NAME": ENG_ITX_NAME,
            "CURRENT_USER_NAME": display_name_from_email(session.get("email", "")),
            "MAX_TABLE_ROWS": MAX_TABLE_ROWS,
            "TRAFFIC_TYPES": TRAFFIC_TYPES,
            # Parâmetros de Programação (respondidos pela operadora)
            "PARAMS_SECOES": PARAMS_SECOES,
            "PARAMS_TOTAL": TOTAL_RESPONDIVEIS,
            "param_key": param_key,
        }


def register_template_filters(app: Flask) -> None:
    @app.template_filter("date_br")
    def date_br_filter(value) -> str:
        if not value:
            return ""
        if hasattr(value, "strftime"):
            try:
                return value.strftime("%d/%m/%Y %H:%M")
            except (ValueError, OSError):
                return str(value)
        dt = parse_db_datetime(value)
        return dt.strftime("%d/%m/%Y %H:%M") if dt else str(value)
