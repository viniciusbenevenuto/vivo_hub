"""params_import.py — Leitura do PTI devolvido pela operadora.

Fluxo: a Engenharia gera o PTI e o envia à operadora, que preenche a aba
"Parâmetros de Programação" (a própria coluna de valores e a coluna
"Cumpre?") e devolve o arquivo. A Engenharia recarrega esse arquivo aqui e
a Seção 9 do formulário passa a mostrar o que foi respondido.

A lista de parâmetros vem de params_programacao.py — a mesma que gera a aba.
Este módulo apenas LÊ o arquivo; nada é gravado no banco aqui.
"""
import logging
import unicodedata
from typing import Any

from params_programacao import PARAMS_BY_KEY, TOTAL_RESPONDIVEIS, param_key

logger = logging.getLogger(__name__)

SHEET_NAME = "Parâmetros de Programação"

# Layout da aba (ver excel/builder.py::_build_params_sheet)
COL_ITEM      = 2   # B — nome do parâmetro
COL_VALOR     = 3   # C — valor requerido pela VIVO
COL_RESPOSTA  = 4   # D — valor informado pela operadora
COL_GRAU      = 5   # E — Mandatório / Default
COL_CUMPRE    = 6   # F — ✔ / ❌ da operadora
COL_OBS       = 7   # G — observação

# A lista suspensa oferece ✔ / ❌, mas operadoras costumam devolver variações.
_SIM = {"✔", "✓", "✅", "sim", "s", "yes", "y", "ok", "x", "atende",
        "cumpre", "true", "1"}
_NAO = {"❌", "✘", "✗", "❎", "nao", "n", "no", "-", "--", "nao cumpre",
        "nao atende", "false", "0"}

MAX_UPLOAD_BYTES = 10 * 1024 * 1024   # 10 MB


class ParamsImportError(Exception):
    """Arquivo inválido ou fora do formato esperado."""


def _norm(value: Any) -> str:
    """Texto limpo: sem espaços duplos, sem seletores de emoji."""
    txt = str(value or "").replace("\xa0", " ")
    txt = "".join(c for c in txt if unicodedata.category(c) != "Cf")
    return " ".join(txt.split())


def _sem_acento(txt: str) -> str:
    base = unicodedata.normalize("NFKD", txt)
    return "".join(c for c in base if not unicodedata.combining(c))


def interpretar_resposta(valor: Any) -> bool | None:
    """'✔'/'Sim' → True, '❌'/'Não' → False, vazio/desconhecido → None."""
    txt = _norm(valor).lower()
    if not txt:
        return None
    if txt in _SIM:
        return True
    if txt in _NAO or _sem_acento(txt) in _NAO:
        return False
    return None


def ler_parametros(fonte) -> dict:
    """Extrai as respostas da aba "Parâmetros de Programação".

    `fonte` é um caminho ou um objeto de arquivo (upload do Flask).

    Retorno:
        {
          "respostas":   {key: True/False},          # só o que foi respondido
          "valores":     {key: "texto da operadora"},
          "detalhes":    {key: {secao, item, valor, resposta, ...}},
          "total":       parâmetros encontrados no arquivo,
          "respondidos": quantos vieram com ✔ ou ❌,
          "esperados":   quantos deveriam ser respondidos,
          "desconhecidos": itens do arquivo fora da lista oficial,
        }
    """
    from openpyxl import load_workbook

    try:
        wb = load_workbook(fonte, data_only=True, read_only=True)
    except Exception as exc:                       # arquivo corrompido/não-xlsx
        raise ParamsImportError(
            "Não foi possível abrir o arquivo. Envie o .xlsx do PTI."
        ) from exc

    try:
        if SHEET_NAME not in wb.sheetnames:
            raise ParamsImportError(
                f'A aba "{SHEET_NAME}" não foi encontrada no arquivo. '
                "Confirme se é o PTI gerado por este sistema."
            )
        ws = wb[SHEET_NAME]

        respostas: dict[str, bool] = {}
        valores: dict[str, str] = {}
        detalhes: dict[str, dict] = {}
        desconhecidos: list[str] = []
        secao = ""
        total = 0

        for row in ws.iter_rows(min_row=6, values_only=True):
            def cell(idx: int) -> str:
                pos = idx - 1
                return _norm(row[pos]) if pos < len(row) else ""

            item = cell(COL_ITEM)
            if not item:
                continue

            # Cabeçalho de colunas da seção
            if cell(COL_CUMPRE).lower() in ("cumpre?", "cumpre"):
                continue

            # Título de seção: apenas a coluna do item preenchida
            if not any(cell(c) for c in (COL_VALOR, COL_RESPOSTA, COL_GRAU,
                                         COL_CUMPRE, COL_OBS)):
                secao = item
                continue

            key = param_key(secao, item)
            if key not in PARAMS_BY_KEY:
                desconhecidos.append(f"{secao} · {item}")
                continue

            total += 1
            resposta = interpretar_resposta(cell(COL_CUMPRE))
            valor_op = cell(COL_RESPOSTA)
            if resposta is not None:
                respostas[key] = resposta
            if valor_op:
                valores[key] = valor_op
            detalhes[key] = {
                "secao":    secao,
                "item":     item,
                "valor":    cell(COL_VALOR),
                "grau":     cell(COL_GRAU),
                "obs":      cell(COL_OBS),
                "resposta": resposta,
                "valor_operadora": valor_op,
            }

        if not total:
            raise ParamsImportError(
                f'A aba "{SHEET_NAME}" não contém parâmetros reconhecidos. '
                "Confirme se o arquivo é o PTI gerado por este sistema."
            )

        if desconhecidos:
            logger.warning(
                "Parâmetros fora da lista oficial ignorados (%d): %s",
                len(desconhecidos), desconhecidos[:5],
            )
        logger.info(
            "Parâmetros importados: %d respondidos de %d encontrados.",
            len(respostas), total,
        )
        return {
            "respostas":     respostas,
            "valores":       valores,
            "detalhes":      detalhes,
            "total":         total,
            "respondidos":   len(respostas),
            "esperados":     TOTAL_RESPONDIVEIS,
            "desconhecidos": desconhecidos,
        }
    finally:
        wb.close()
