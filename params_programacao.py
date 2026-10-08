"""params_programacao.py — Lista oficial dos Parâmetros de Programação.

Fonte única de verdade para três lugares:
  1. aba "Parâmetros de Programação" do PTI (excel/builder.py);
  2. Seção 9 do formulário (templates/formulario_atacado.html);
  3. leitura do PTI devolvido pela operadora (params_import.py).

Quem responde é a OPERADORA: ela recebe o PTI, preenche a coluna com o
próprio valor e marca "Cumpre?" (✔ / ❌). A Engenharia recarrega o arquivo
no sistema e a Seção 9 mostra o que foi respondido.

Para alterar a lista, edite apenas este arquivo.

Layout de cada seção na planilha (colunas B..G):
    B: item        C: valor requerido (VIVO)   D: resposta da operadora
    E: grau        F: Cumpre?                  G: observação
"""
import re
import unicodedata
from typing import Final

# Marca usada como "ainda não respondido" na coluna Cumpre?
NAO_RESPONDIDO: Final[str] = "❌"

# ---------------------------------------------------------------------------
# ESTRUTURA
#   secao      — título da seção (linha roxa na planilha)
#   col_item   — rótulo da 1ª coluna (varia: Item / Prioridade / Método / RFC)
#   col_valor  — rótulo da 2ª coluna ("" quando a seção não tem valor requerido)
#   itens      — (item, valor, grau, observação, responder)
#                responder=False → linha informativa, sem "Cumpre?"
# ---------------------------------------------------------------------------
PARAMS_SECOES: Final[tuple[dict, ...]] = (
    {
        "secao": "Interconexão Nacional",
        "col_item": "Item", "col_valor": "TELEFÔNICA VIVO",
        "itens": (
            ("Fabricante / Fornecedor", "", "Mandatório", "", True),
            ("Modelo / Versão SW", "Elementos Rede Fixa I, Fixa II e Móvel", "Mandatório", "", True),
            ("Nós SIP", "", "Mandatório", "", True),
        ),
    },
    {
        "secao": "Tipo de Serviço",
        "col_item": "Item", "col_valor": "TELEFÔNICA VIVO",
        "itens": (
            ("Voz",  "SIM", "Mandatório", "", True),
            ("Fax",  "SIM", "Mandatório", "", True),
            ("DTMF", "SIM", "Mandatório", "", True),
        ),
    },
    {
        "secao": "Protocolo",
        "col_item": "Item", "col_valor": "TELEFÔNICA VIVO",
        "itens": (
            ("Tipo de Protocolo", "SIP-I (Q.1912.5)", "Mandatório", "", True),
        ),
    },
    {
        "secao": "Atributos SIP",
        "col_item": "Item", "col_valor": "Valor Requerido",
        "itens": (
            ("SIP Version", "2", "Mandatório", "", True),
            ("Protocolo de Transporte", "UDP", "Mandatório", "", True),
            ("Tipo de Interface", "Gateway to Gateway (NNI)", "Mandatório", "", True),
            ("Endereço IP Sinalização SIP", "A ser definido", "Mandatório", "", True),
            ("Porta Agente SIP", "5060", "Mandatório", "", True),
            ("Endereço IP RTP", "IP ADDR do RTP", "Mandatório", "", True),
            ("RTP Port", "1024 a 65000", "Mandatório", "", True),
            ("FW ou SBC antes da Core", "SIM", "Mandatório", "", True),
            ("P-Charging Vector", "Suporta, enviamos ou não", "Mandatório", "", True),
            ("SIP Domain", "IP do SIP Agent ou domínio", "Mandatório", "", True),
            ("FROM", "Envia DOMAIN/IP", "Mandatório", "", True),
            ("TO", "Envia DOMAIN/IP", "Mandatório", "", True),
            ("VIA", "Envia DOMAIN/IP", "Mandatório", "", True),
            ("Contact", "Envia IP ou Nome de Domínio", "Mandatório", "", True),
            ("Keep Alive da Rota", "OPTIONS (preferencial)", "Mandatório", "", True),
            ("Confirmação OPTIONS", "200 OK", "Mandatório", "", True),
            ("SDP no INVITE", "SIM", "Mandatório", "", True),
            ("PRACK", "Suporta", "Mandatório", "", True),
            ("Resposta PRACK", "200 OK", "Mandatório", "", True),
            ("Formato INVITE", "Enbloc", "Mandatório", "", True),
            ("Desligamento para Reroteamento", "Causa #47", "Mandatório", "", True),
            ("Causas Especiais de Desligamento", "N.A.", "Mandatório", "", True),
            ("ISUP Encapsulado", "Q.1912.5", "Mandatório", "", True),
            ("Identificação Originador Rede Fixa", "ISUP Encapsulado", "Mandatório", "ISUPBR (Região III)", True),
            ("Identificação Originador Rede Móvel", "ISUP Encapsulado", "Mandatório", "ISUPBR / opcional ITU92", True),
            ("Domínio FQDN", "Designação da rota", "Default", "", True),
        ),
    },
    {
        "secao": "Codec Rede Móvel",
        "col_item": "Prioridade", "col_valor": "Codec",
        "itens": (
            ("1ª opção", "AMR Narrowband", "Mandatório", "", True),
            ("2ª opção", "G.711A 20 ms (8)", "Mandatório", "Não aceita G.711U", True),
        ),
    },
    {
        "secao": "Codec Rede Fixa",
        "col_item": "Prioridade", "col_valor": "Codec",
        "itens": (
            ("1ª opção", "G.711A 20 ms (8)", "Mandatório", "Não aceita G.711U", True),
            ("2ª opção", "G.729A 20 ms (18)", "Mandatório", "", True),
        ),
    },
    {
        "secao": "DTMF",
        "col_item": "Item", "col_valor": "Valor Requerido",
        "itens": (
            ("DTMF (1ª opção)", "RFC 2833 OutBand Payload 100", "Mandatório", "", True),
            ("DTMF (2ª opção)", "Inband (G.711)", "", "", True),
            ("Payloads permitidos", "96 a 125", "Mandatório", "", True),
        ),
    },
    {
        "secao": "FAX",
        "col_item": "Item", "col_valor": "Valor Requerido",
        "itens": (
            ("T.38", "Não suportado", "", "", False),
            ("Detecção UPSPEED", "G.711A", "Mandatório", "Dados necessários no SDP", True),
            ("Re-INVITE para Fax", "Utilização de G.711", "Mandatório", "gpmd=8 / vbd=yes / ecan:fb on", True),
        ),
    },
    {
        "secao": "POS",
        "col_item": "Item", "col_valor": "Valor Requerido",
        "itens": (
            ("Detecção UPSPEED", "G.711A", "", "Dados necessários no SDP", False),
            ("Re-INVITE para POS", "Utilização de G.711", "Mandatório", "gpmd=8 / vbd=yes / ecan:fb off", True),
        ),
    },
    {
        "secao": "Encaminhamento Entrante",
        "col_item": "Item", "col_valor": "Valor Requerido",
        "itens": (
            ("A-Number (Calling Number)", "E.164 Nacional", "Mandatório", "", True),
            ("P-Asserted Identity", "SIM (RFC 3325) E.164", "Mandatório", "", True),
            ("NOA Calling Number", "SUB / NAT / UNKW", "Mandatório", "", True),
            ("B-Number (Called Number)", "E.164 Roteamento", "Mandatório", "", True),
            ("NOA Called Number", "SUB / NAT / UNKW", "Mandatório", "", True),
        ),
    },
    {
        "secao": "Encaminhamento Sainte",
        "col_item": "Item", "col_valor": "Valor Requerido",
        "itens": (
            ("A-Number (Calling Number)", "E.164 Nacional", "Mandatório", "", True),
            ("P-Asserted Identity", "SIM (RFC 3325) E.164", "Mandatório", "", True),
            ("NOA Calling Number", "SUB / NAT / UNKW", "Mandatório", "", True),
            ("B-Number (Called Number)", "E.164 Roteamento", "Mandatório", "", True),
            ("NOA Called Number", "SUB / NAT / UNKW", "Mandatório", "", True),
        ),
    },
    {
        "secao": "RFCs Obrigatórias / Suportadas",
        "col_item": "RFC", "col_valor": "Descrição",
        "itens": (
            ("RFC 3261", "SIP Base", "", "", True),
            ("RFC 3262", "100rel / PRACK", "", "", True),
            ("RFC 3264", "SDP / Hold", "", "", True),
            ("RFC 3266", "SDP", "", "", True),
            ("RFC 2327", "SDP", "", "", True),
            ("RFC 4566", "SDP", "", "", True),
            ("RFC 4028", "SIP Timers", "", "", True),
            ("RFC 3389", "Comfort Noise", "", "", True),
            ("RFC 2833", "Transporte DTMF", "", "", True),
            ("RFC 6337", "Modelo Offer/Answer SIP", "", "", True),
            ("RFC 3960", "Early Media e Ring Back Tone", "",
             "SIP 180 Ringing puro; o Proxy deverá gerar o Ring Back Tone localmente", True),
        ),
    },
    {
        "secao": "Alteração de SDP",
        "col_item": "Método", "col_valor": "",
        "itens": (
            ("UPDATE",    "", "Mandatório", "Antes do atendimento", True),
            ("RE-INVITE", "", "Mandatório", "Após atendimento", True),
        ),
    },
    {
        "secao": "Negociação de Codec",
        "col_item": "Requisito", "col_valor": "",
        "itens": (
            ("Re-INVITE obrigatório para definição de codec quando o destino retornar múltiplos codecs",
             "", "Mandatório", "", True),
            ("Ptime/Maxptime deve ser múltiplo inteiro de 20 ms para dispositivos móveis (3GPP 26.114)",
             "", "Mandatório", "", True),
            ("RTP com marcação CS5, DSCP 46 ou EF", "", "Mandatório", "Ajustar QoS", True),
        ),
    },
)


def _slug(texto: str) -> str:
    """'Atributos SIP' → 'atributos_sip' (sem acento, minúsculo)."""
    txt = unicodedata.normalize("NFKD", str(texto or ""))
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    txt = re.sub(r"[^a-zA-Z0-9]+", "_", txt).strip("_").lower()
    return txt


def param_key(secao: str, item: str) -> str:
    """Chave estável de um parâmetro.

    Inclui a seção porque há itens de mesmo nome em seções diferentes
    (ex.: "A-Number" aparece em Encaminhamento Entrante e Sainte).
    """
    return f"{_slug(secao)}.{_slug(item)}"


def _build_lista() -> tuple[dict, ...]:
    lista: list[dict] = []
    for bloco in PARAMS_SECOES:
        for item, valor, grau, obs, responder in bloco["itens"]:
            lista.append({
                "key":       param_key(bloco["secao"], item),
                "secao":     bloco["secao"],
                "item":      item,
                "valor":     valor,
                "grau":      grau,
                "obs":       obs,
                "responder": responder,
            })
    return tuple(lista)


# Lista achatada, na ordem do documento
PARAMS: Final[tuple[dict, ...]] = _build_lista()

# key -> parâmetro
PARAMS_BY_KEY: Final[dict[str, dict]] = {p["key"]: p for p in PARAMS}

# Quantos a operadora precisa responder
TOTAL_RESPONDIVEIS: Final[int] = sum(1 for p in PARAMS if p["responder"])
