"""
siprouter_sp.py — Base de dados do siprouter de São Paulo.

Responsabilidades:
  - Criar a tabela siprouter_sp no banco SQLite da aplicação.
  - Inserir os 180 registros oficiais SE a tabela estiver vazia.
  - Expor query_siprouter_sp() como única interface de consulta.

Restrições absolutas:
  - NÃO sugere SBCs.
  - NÃO utiliza localidade para busca.
  - NÃO consulta fontes externas.
  - NÃO infere dados.
  - NÃO altera registros após a carga inicial.

Os 180 registros (CDSIP_SPO_PL: 90 + CDSIP_SPO_JG: 90) seguem um padrão
fixo por CN (11–19) e são gerados deterministicamente por _build_seed_data().
A saída é idêntica, campo a campo, à listagem oficial original.
"""
import logging
import sqlite3
from typing import Final

logger = logging.getLogger(__name__)

# =============================================================================
# BASE OFICIAL — regras de geração dos 180 registros
# (elemento, bloco_ip, vlan, vrf, descricao, cn, rn1)
# =============================================================================

# Terceiro octeto do bloco IP de cada elemento (10.<cn>.<octeto>.x/28).
_ELEMENT_THIRD_OCTET: Final[dict[str, int]] = {
    "CDSIP_SPO_PL": 130,
    "CDSIP_SPO_JG": 131,
}

# Código VRF de cada CN (formato final: "<código>:Interconexao_SIP_<cn>_New_H").
_VRF_CODE_BY_CN: Final[dict[int, str]] = {
    11: "V75054", 12: "V75055", 13: "V75057",
    14: "V80898", 15: "V79865", 16: "V79216",
    17: "V79886", 18: "V79894", 19: "V79903",
}

# Rotas de cada CN, na ordem oficial (define VLAN e último octeto do bloco).
_ROUTE_PLAN: Final[tuple[tuple[str, str], ...]] = (
    ("ESPELHINHOS", "Diversos"),
    ("OI",          "55131/55114"),
    ("TIM",         "55123/55141/55341"),
    ("CLARO",       "55121/55321"),
    ("ALGAR",       "55112/55224/55312"),
    ("DATORA",      "55181"),
    ("AV",          "Diversos"),
    ("SCM",         "Diversos"),
    ("RESERVA",     "Diversos"),
    ("RESERVA",     "Diversos"),
)

_FIRST_CN: Final[int] = 11
_LAST_CN: Final[int] = 19
_BASE_VLAN: Final[int] = 402          # VLAN da 1ª rota do CN 11
_VLANS_PER_CN: Final[int] = len(_ROUTE_PLAN)


def _vrf_for(cn: int, rota: str) -> str:
    # Exceção histórica da base oficial: o bloco ESPELHINHOS do CN 16
    # usa a VRF do CN 11.
    vrf_cn = _FIRST_CN if (cn == 16 and rota == "ESPELHINHOS") else cn
    return f"{_VRF_CODE_BY_CN[vrf_cn]}:Interconexao_SIP_{vrf_cn}_New_H"


def _build_seed_data() -> tuple[tuple, ...]:
    rows: list[tuple] = []
    for elemento, third_octet in _ELEMENT_THIRD_OCTET.items():
        for cn in range(_FIRST_CN, _LAST_CN + 1):
            vlan_base = _BASE_VLAN + (cn - _FIRST_CN) * _VLANS_PER_CN
            for idx, (rota, rn1) in enumerate(_ROUTE_PLAN):
                rows.append((
                    elemento,
                    f"10.{cn}.{third_octet}.{idx * 16}/28",
                    vlan_base + idx,
                    _vrf_for(cn, rota),
                    f"ROTAS CN {cn} {rota}",
                    cn,
                    rn1,
                ))
    return tuple(rows)


_SEED_DATA: Final[tuple[tuple, ...]] = _build_seed_data()


# =============================================================================
# INICIALIZAÇÃO — cria tabela e carrega seed se vazia
# =============================================================================
def init_siprouter_sp(db: sqlite3.Connection) -> None:
    """Garante que a tabela siprouter_sp existe e está preenchida."""
    db.execute("""
        CREATE TABLE IF NOT EXISTS siprouter_sp (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            elemento  TEXT    NOT NULL,
            bloco_ip  TEXT    NOT NULL,
            vlan      INTEGER NOT NULL,
            vrf       TEXT    NOT NULL,
            descricao TEXT    NOT NULL,
            cn        INTEGER NOT NULL,
            rn1       TEXT    NOT NULL
        )
    """)

    count = db.execute("SELECT COUNT(*) FROM siprouter_sp").fetchone()[0]
    if count == 0:
        db.executemany(
            "INSERT INTO siprouter_sp "
            "(elemento, bloco_ip, vlan, vrf, descricao, cn, rn1) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            _SEED_DATA,
        )
        db.commit()
        logger.info("siprouter_sp: %d registros inseridos.", len(_SEED_DATA))
    else:
        logger.debug("siprouter_sp: tabela já preenchida (%d registros).", count)


# =============================================================================
# HELPERS INTERNOS
# =============================================================================
def _rn1_matches(table_rn1: str, user_rn1: str) -> bool:
    """
    Retorna True se o RN1 do usuário está contido no campo rn1 da tabela.

    Regras:
      - "Diversos" sempre corresponde a qualquer RN1.
      - Para valores como "55131/55114", o RN1 do usuário deve ser um dos
        elementos da lista separada por "/".
    """
    if table_rn1 == "Diversos":
        return True
    return user_rn1.strip() in [part.strip() for part in table_rn1.split("/")]


def _desc_is_espelho(descricao: str) -> bool:
    return "ESPELH" in descricao.upper()


def _desc_is_scm(descricao: str) -> bool:
    return "SCM" in descricao.upper()


def _desc_is_av(descricao: str) -> bool:
    return "AV" in descricao.upper()


def _desc_is_operadora(descricao: str) -> bool:
    """Registro de operadora: não é ESPELHO, RESERVA, SCM nem AV."""
    d = descricao.upper()
    return not any(kw in d for kw in ("ESPELH", "RESERVA", "SCM", "AV"))


def _pair_pl_jg(rows: list[dict], descricao: str) -> dict | None:
    """Retorna o par PL/JG de uma descrição, ou None se não houver registros."""
    pl = next(
        (r for r in rows
         if r["elemento"] == "CDSIP_SPO_PL" and r["descricao"] == descricao),
        None,
    )
    jg = next(
        (r for r in rows
         if r["elemento"] == "CDSIP_SPO_JG" and r["descricao"] == descricao),
        None,
    )
    if not pl and not jg:
        return None

    ref = pl or jg
    return {
        "descricao": ref["descricao"],
        "vlan":      ref["vlan"],
        "vrf":       ref["vrf"],
        "rn1":       ref["rn1"],
        "pl_bloco":  pl["bloco_ip"] if pl else None,
        "jg_bloco":  jg["bloco_ip"] if jg else None,
    }


# =============================================================================
# QUERY PÚBLICA — retorna SEMPRE um único bloco (par PL/JG)
# =============================================================================
def query_siprouter_sp(
    db: sqlite3.Connection,
    cn: int,
    rn1: str,
    scm: bool = False,
    av: bool = False,
) -> dict:
    """
    Consulta a tabela siprouter_sp e retorna EXATAMENTE UM bloco (par PL/JG).

    Prioridade de seleção:
      1. SCM marcado  → bloco SCM do CN informado.
      2. AV marcado   → bloco AV do CN informado.
      3. Nenhum flag  → bloco cuja lista de RN1 contenha o RN1 informado
                        (operadoras específicas).
         Fallback     → se nenhuma operadora casar com o RN1, retorna o
                        bloco ESPELHINHOS do CN informado.

    Retorno:
        {
          "found":          bool,
          "cn":             int,
          "rn1_consultado": str,
          "descricao":      str,
          "vlan":           int | None,
          "vrf":            str | None,
          "pl_bloco":       str | None,
          "jg_bloco":       str | None,
          "mensagem":       str,
        }
    """
    def _not_found(msg: str) -> dict:
        return {
            "found": False, "cn": cn, "rn1_consultado": rn1,
            "descricao": None, "vlan": None, "vrf": None,
            "pl_bloco": None, "jg_bloco": None, "mensagem": msg,
        }

    def _found(par: dict, mensagem: str) -> dict:
        return {
            **par, "found": True, "cn": cn,
            "rn1_consultado": rn1, "mensagem": mensagem,
        }

    rows = db.execute(
        "SELECT elemento, bloco_ip, vlan, vrf, descricao, cn, rn1 "
        "FROM siprouter_sp WHERE cn = ? ORDER BY vlan ASC",
        (cn,),
    ).fetchall()

    if not rows:
        return _not_found(f"Não existe BLOCO IP disponível para CN {cn}.")

    rows = [dict(r) for r in rows]

    # 1 e 2. Flags SCM/AV → bloco dedicado do CN
    for flag, predicate, label in (
        (scm, _desc_is_scm, "SCM"),
        (av, _desc_is_av, "AV"),
    ):
        if not flag:
            continue
        flag_rows = [r for r in rows if predicate(r["descricao"])]
        if not flag_rows:
            return _not_found(f"Nenhum bloco {label} encontrado para CN {cn}.")
        par = _pair_pl_jg(rows, flag_rows[0]["descricao"])
        if not par:
            return _not_found(
                f"Par PL/JG não encontrado para bloco {label} — CN {cn}."
            )
        return _found(par, f"Bloco {label} · CN {cn}.")

    # 3. Sem flags → tentar casar RN1 com operadora específica
    operadora_rows = [
        r for r in rows
        if _desc_is_operadora(r["descricao"]) and _rn1_matches(r["rn1"], rn1)
    ]
    if operadora_rows:
        desc = operadora_rows[0]["descricao"]
        par = _pair_pl_jg(rows, desc)
        if par:
            return _found(par, f"{desc} · CN {cn}.")

    # 4. Fallback → bloco ESPELHINHOS do CN
    espelho_rows = [r for r in rows if _desc_is_espelho(r["descricao"])]
    if not espelho_rows:
        return _not_found(f"Nenhum bloco encontrado para CN {cn} / RN1 {rn1}.")
    desc = espelho_rows[0]["descricao"]
    par = _pair_pl_jg(rows, desc)
    if not par:
        return _not_found(f"Par PL/JG não encontrado para ESPELHINHOS — CN {cn}.")
    return _found(par, f"{desc} · CN {cn} (fallback ESPELHINHOS).")
