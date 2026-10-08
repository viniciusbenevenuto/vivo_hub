"""queries.py — Construtores de queries SQL reutilizáveis."""
import sqlite3
from typing import Optional

# Chaves de ordenação aceitas por build_list_query (whitelist — nunca
# interpolar valores vindos do usuário diretamente no ORDER BY).
_SORT_MAP: dict[str, str] = {
    "created_at":        "created_at ASC",
    "-created_at":       "created_at DESC",
    "nome_operadora":    "nome_operadora COLLATE NOCASE ASC",
    "-nome_operadora":   "nome_operadora COLLATE NOCASE DESC",
    "id":                "id ASC",
    "-id":               "id DESC",
    "operadora_version": "nome_operadora COLLATE NOCASE ASC, version ASC",
}


def build_list_query(
    base_sql: str,
    *,
    search_term: Optional[str] = None,
    status_filter: Optional[str] = None,
    owner_id: Optional[int] = None,
    sort_key: str = "-created_at",
) -> tuple[str, list]:
    """Monta SELECT com filtros opcionais e ordenação segura."""
    conditions: list[str] = []
    params: list = []

    if owner_id is not None:
        conditions.append("owner_id = ?")
        params.append(owner_id)
    if search_term:
        conditions.append("nome_operadora LIKE ?")
        params.append(f"%{search_term}%")
    if status_filter:
        conditions.append("LOWER(COALESCE(status, '')) = LOWER(?)")
        params.append(status_filter)

    if conditions:
        base_sql += " WHERE " + " AND ".join(conditions)

    base_sql += f" ORDER BY {_SORT_MAP.get(sort_key, 'id DESC')}"
    return base_sql, params


def get_status_counters(
    db: sqlite3.Connection,
    extra_where: str = "",
    extra_params: tuple = (),
) -> dict[str, int]:
    """Conta formulários por status, com filtro WHERE opcional."""
    connector = "AND" if "WHERE" in extra_where else "WHERE"

    def _count(cond: str = "") -> int:
        sql = f"SELECT COUNT(*) AS c FROM atacado_forms {extra_where}"
        if cond:
            sql += f" {connector} {cond}"
        return db.execute(sql, extra_params).fetchone()["c"]

    return {
        "total":      _count(),
        "rascunho":   _count("LOWER(status) = 'rascunho'"),
        "enviado":    _count("LOWER(status) = 'enviado'"),
        "em_revisao": _count("LOWER(status) = 'em revisão'"),
        "aprovado":   _count("LOWER(status) = 'aprovado'"),
    }
