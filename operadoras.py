"""operadoras.py — Base de operadoras do Anexo 5 (ABR Telecom).

Fonte oficial: https://www.abrtelecom.com.br/padronizacao (Anexo 5 —
Codificação das Prestadoras). O XLSX é importado para a tabela `operadoras`,
que alimenta o autopreenchimento restrito (RN1 ⇄ Nome ⇄ UF ⇄ EOT) do
formulário de PTI.

Somente linhas com RN1 numérico de até 5 dígitos são importadas — é o
identificador usado pelo formulário.
"""
import logging
import os
import sqlite3
from typing import Final

logger = logging.getLogger(__name__)

# Layout do Anexo 5: cabeçalho na linha 8 (índice 7); colunas relevantes.
_HEADER_ROW_INDEX: Final[int] = 7
_COLUMNS: Final[dict[str, int]] = {
    "eot": 1,
    "nome": 2,
    "razao_social": 3,
    "csp": 4,
    "tipo_servico": 5,
    "modalidade": 6,
    "area_prestacao": 7,
    "holding": 8,
    "uf": 13,
    "regiao": 14,
    "concessao": 15,
    "rn1": 16,
    "spid": 17,
}

_SCHEMA_SQL: Final[str] = """
    CREATE TABLE IF NOT EXISTS operadoras (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        eot TEXT, nome TEXT NOT NULL, razao_social TEXT,
        csp TEXT, tipo_servico TEXT, modalidade TEXT,
        area_prestacao TEXT, holding TEXT, uf TEXT,
        regiao TEXT, concessao TEXT,
        rn1 TEXT NOT NULL, spid TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_operadoras_rn1  ON operadoras (rn1);
    CREATE INDEX IF NOT EXISTS idx_operadoras_nome ON operadoras (nome);
    CREATE INDEX IF NOT EXISTS idx_operadoras_uf   ON operadoras (uf);
"""


def rn1_is_valid(rn1: str) -> bool:
    """RN1 válido: somente dígitos, no máximo 5."""
    return rn1.isdigit() and len(rn1) <= 5


def _cell(row: tuple, idx: int) -> str:
    value = row[idx] if idx < len(row) else None
    return str(value).strip() if value is not None else ""


def parse_anexo5(path: str) -> list[dict]:
    """Lê o XLSX do Anexo 5 e retorna as linhas com RN1 válido."""
    from openpyxl import load_workbook

    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        rows: list[dict] = []
        for i, raw in enumerate(ws.iter_rows(values_only=True)):
            if i <= _HEADER_ROW_INDEX:
                continue
            record = {field: _cell(raw, idx) for field, idx in _COLUMNS.items()}
            if not record["nome"] or not rn1_is_valid(record["rn1"]):
                continue
            record["uf"] = record["uf"].upper()
            record["eot"] = record["eot"].upper()
            rows.append(record)
        return rows
    finally:
        wb.close()


def import_anexo5(db: sqlite3.Connection, path: str) -> int:
    """Substitui o conteúdo da tabela operadoras pelo XLSX informado."""
    rows = parse_anexo5(path)
    if not rows:
        raise ValueError(f"Nenhuma linha válida encontrada em: {path}")

    fields = list(_COLUMNS.keys())
    db.executescript(_SCHEMA_SQL)
    db.execute("DELETE FROM operadoras")
    db.executemany(
        f"INSERT INTO operadoras ({','.join(fields)}) "
        f"VALUES ({','.join(['?'] * len(fields))})",
        [tuple(r[f] for f in fields) for r in rows],
    )
    db.commit()
    logger.info("operadoras: %d registros importados de %s", len(rows), path)
    return len(rows)


def init_operadoras(db: sqlite3.Connection, default_xlsx: str) -> None:
    """Cria a tabela e importa o Anexo 5 embarcado se ela estiver vazia."""
    db.executescript(_SCHEMA_SQL)
    count = db.execute("SELECT COUNT(*) FROM operadoras").fetchone()[0]
    if count:
        logger.debug("operadoras: tabela já preenchida (%d registros).", count)
        return
    if not os.path.exists(default_xlsx):
        logger.warning(
            "operadoras: tabela vazia e Anexo 5 não encontrado em %s. "
            "Execute 'flask import-anexo5 <caminho>'.", default_xlsx,
        )
        return
    import_anexo5(db, default_xlsx)


def list_operadoras(db: sqlite3.Connection) -> list[dict]:
    """Linhas compactas para o autopreenchimento do formulário.

    modalidade é normalizada (espaços especiais, caixa) para permitir a
    segmentação Local × Longa Distância dos campos EOT no frontend.
    """
    rows = db.execute(
        "SELECT rn1, nome, uf, eot, tipo_servico, modalidade "
        "FROM operadoras ORDER BY nome COLLATE NOCASE, uf"
    ).fetchall()
    result = []
    for r in rows:
        rec = dict(r)
        rec["modalidade"] = " ".join(rec["modalidade"].replace("\xa0", " ").split())
        result.append(rec)
    return result


def rn1_nome_match(db: sqlite3.Connection, rn1: str, nome: str) -> bool:
    """True se o par RN1 + Nome existe no Anexo 5."""
    return db.execute(
        "SELECT 1 FROM operadoras WHERE rn1 = ? AND nome = ? LIMIT 1",
        (rn1, nome),
    ).fetchone() is not None
