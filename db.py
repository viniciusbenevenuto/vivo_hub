"""db.py — Conexão, schema e seeds do banco SQLite."""
import logging
import sqlite3

from flask import Flask, current_app, g

from config import ANEXO5_XLSX_PATH, CN_METADATA, CN_SEED_RAW
from operadoras import init_operadoras
from siprouter_sp import init_siprouter_sp

logger = logging.getLogger(__name__)

# Colunas adicionadas após a criação inicial da tabela (migrações leves).
_ATACADO_FORMS_MIGRATIONS: tuple[tuple[str, str], ...] = (
    ("owner_id",               "INTEGER NOT NULL DEFAULT 0"),
    ("engenharia_params_json", "TEXT"),
    ("dados_vivo_json",        "TEXT"),
    ("dados_operadora_json",   "TEXT"),
    ("escopo_text",            "TEXT"),
    ("escopo_flags_json",      "TEXT"),
    ("responsavel_atacado",    "TEXT"),
    ("responsavel_engenharia", "TEXT"),
    ("version",                "INTEGER NOT NULL DEFAULT 1"),
    ("parent_id",              "INTEGER DEFAULT NULL"),
    ("scm",                    "INTEGER DEFAULT 0"),
    ("av",                     "INTEGER DEFAULT 0"),
    ("updated_at",             "TEXT"),
    # Banda / Canais / CAPS por tipo de tráfego (preenchido pela Engenharia)
    ("trafego_metricas_json",  "TEXT"),
    # Respostas da operadora importadas do PTI devolvido (params_import.py)
    ("operadora_params_json",  "TEXT"),
    ("operadora_params_em",    "TEXT"),
    # Seção Escopo — "Projeto padrão"
    ("projeto_padrao",         "INTEGER DEFAULT 0"),
)

_SCHEMA_SQL = """
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        email TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        role TEXT NOT NULL CHECK (role IN ('engenharia', 'atacado')),
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS atacado_forms (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        owner_id INTEGER NOT NULL,
        status TEXT DEFAULT 'rascunho',
        nome_operadora TEXT, rn1 TEXT,
        csp INTEGER DEFAULT 0, servicos_especiais INTEGER DEFAULT 0,
        cng INTEGER DEFAULT 0, atendimento TEXT, redes TEXT,
        qual TEXT, tmr TEXT,
        responsavel_operadora TEXT, responsavel_vivo TEXT,
        sbc_ativo INTEGER DEFAULT 0, ip_reservado INTEGER DEFAULT 0,
        vivo_reserva INTEGER DEFAULT 0, asn TEXT,
        escopo_text TEXT, escopo_flags_json TEXT, dados_vivo_json TEXT,
        dados_operadora_json TEXT,
        operadora_ciente INTEGER DEFAULT 0, responsavel_infra TEXT,
        lcr_nacional INTEGER DEFAULT 0, white_list INTEGER DEFAULT 0,
        prefixos_liberados_abr INTEGER DEFAULT 0,
        premissas_ok INTEGER DEFAULT 0, aprovado_por TEXT,
        engenharia_params_json TEXT,
        responsavel_atacado TEXT, responsavel_engenharia TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(owner_id) REFERENCES users(id)
    );
    CREATE TABLE IF NOT EXISTS exports (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        form_id INTEGER NOT NULL,
        filename TEXT NOT NULL, filepath TEXT NOT NULL,
        size_bytes INTEGER DEFAULT 0,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(form_id) REFERENCES atacado_forms(id)
    );
"""


# =============================================================================
# CONEXÃO
# =============================================================================
def get_db() -> sqlite3.Connection:
    """Retorna a conexão da requisição atual, criando-a se necessário."""
    if "db" not in g:
        g.db = sqlite3.connect(current_app.config["DATABASE"])
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON;")
        g.db.execute("PRAGMA journal_mode = WAL;")
        g.db.execute("PRAGMA synchronous = NORMAL;")
        g.db.execute("PRAGMA cache_size = -64000;")
    return g.db


def register_db_hooks(app: Flask) -> None:
    """Fecha a conexão ao final de cada requisição."""
    @app.teardown_appcontext
    def close_db(exception=None):
        db = g.pop("db", None)
        if db is not None:
            db.close()


def init_db(app: Flask) -> None:
    """Cria/migra o schema e aplica os seeds. Executado uma vez no startup."""
    with app.app_context():
        db = get_db()
        db.executescript(_SCHEMA_SQL)
        for column, coldef in _ATACADO_FORMS_MIGRATIONS:
            _ensure_column(db, "atacado_forms", column, coldef)
        db.commit()
        seed_cns(db)
        apply_cn_metadata(db)
        init_siprouter_sp(db)
        init_operadoras(db, ANEXO5_XLSX_PATH)
        logger.info("Schema do banco verificado/atualizado.")


def _ensure_column(
    db: sqlite3.Connection, table: str, column: str, coldef: str
) -> None:
    existing = {
        row["name"]
        for row in db.execute(f"PRAGMA table_info({table});").fetchall()
    }
    if column not in existing:
        db.execute(f"ALTER TABLE {table} ADD COLUMN {column} {coldef};")


# =============================================================================
# SEEDS
# =============================================================================
def _parse_cn_seed(raw: str) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for token in raw.split():
        code = token.strip().zfill(2)
        if code.isdigit() and code not in seen:
            seen.add(code)
            result.append(code)
    return result


def seed_cns(db: sqlite3.Connection) -> None:
    """Cria a tabela de CNs e insere os códigos se estiver vazia."""
    db.execute("""
        CREATE TABLE IF NOT EXISTS cns (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT UNIQUE NOT NULL, nome TEXT, uf TEXT,
            ativo INTEGER NOT NULL DEFAULT 1
        )
    """)
    if db.execute("SELECT COUNT(*) AS c FROM cns").fetchone()["c"] == 0:
        codes = _parse_cn_seed(CN_SEED_RAW)
        db.executemany(
            "INSERT OR IGNORE INTO cns (codigo, ativo) VALUES (?, 1)",
            [(code,) for code in codes],
        )
        db.commit()


def apply_cn_metadata(db: sqlite3.Connection) -> None:
    """Atualiza nome/UF de cada CN a partir de CN_METADATA."""
    for code, (nome, uf) in CN_METADATA.items():
        db.execute(
            """
            INSERT INTO cns (codigo, nome, uf, ativo) VALUES (?, ?, ?, 1)
            ON CONFLICT(codigo) DO UPDATE
                SET nome=excluded.nome, uf=excluded.uf, ativo=1
            """,
            (code, nome, uf),
        )
    db.commit()
