"""config.py — Constantes e configurações da aplicação VIVOHUB.

Este módulo não importa Flask; é importado por todos os outros.

Variáveis de ambiente reconhecidas:
    SECRET_KEY          — chave de sessão do Flask (obrigatória em produção)
    ADMIN_CODE          — código de acesso à área administrativa
    SBC_DIR             — diretório com as planilhas XLSX de SBC
    DIAGRAM_IMAGE_PATH  — caminho da imagem base do diagrama de interligação
    IPAM_BASE_URL / IPAM_USER / IPAM_PASS — acesso à API phpIPAM
"""
import os
from typing import Final

from traffic_types import TRAFFIC_TYPE_NAMES

_BASE_DIR: Final[str] = os.path.dirname(os.path.abspath(__file__))

# =============================================================================
# APLICAÇÃO
# =============================================================================
MAX_TABLE_ROWS: Final[int] = 5

# ATENÇÃO: valores padrão apenas para desenvolvimento local.
# Em produção, defina ADMIN_CODE via variável de ambiente.
DEFAULT_ADMIN_CODE: Final[str] = "v28B112004"

# ┌─────────────────────────────────────────────────────────────────────┐
# │ COLOQUE AQUI O NOME DO ENG DE ITX.                                  │
# │ Este nome preenche automaticamente os campos "Eng de ITX" e         │
# │ "Aprovado por" de todos os formulários.                             │
# └─────────────────────────────────────────────────────────────────────┘
ENG_ITX_NAME: Final[str] = "Vinicius"

# Anexo 5 (ABR Telecom) — fonte da tabela de operadoras/RN1.
# Atualização: https://www.abrtelecom.com.br/padronizacao
# Para reimportar: flask import-anexo5 [caminho-do-xlsx]
ANEXO5_XLSX_PATH: Final[str] = os.environ.get(
    "ANEXO5_XLSX_PATH", os.path.join(_BASE_DIR, "data", "anexo5.xlsx")
)


def _resolve_diagram_image() -> str:
    """Resolve o caminho da imagem do diagrama em localizações conhecidas."""
    # 1. Variável de ambiente (Docker / produção)
    env_path = os.environ.get("DIAGRAM_IMAGE_PATH", "")
    if env_path and os.path.exists(env_path):
        return env_path

    parent = os.path.dirname(_BASE_DIR)
    names = [
        "Mídia (3).jpg",
        "Midia (3).jpg",
        "20251007_140942_0000.png",
        "diagram.png",
        "diagram.jpg",
    ]

    # 2. Dentro do projeto (static/, templates/)
    for name in names:
        for folder in ("static", "templates"):
            path = os.path.join(_BASE_DIR, folder, name)
            if os.path.exists(path):
                return path

    # 3. Pasta irmã VIVOHUB (estrutura: Desktop/vivo_hub + Desktop/VIVOHUB)
    for sibling in ("VIVOHUB", "vivohub", "VIVOHub", "VIVO_HUB"):
        for name in names:
            for folder in ("templates", "static"):
                path = os.path.join(parent, sibling, folder, name)
                if os.path.exists(path):
                    return path

    # 4. Fallback — o builder do Excel loga erro se o arquivo não existir
    return os.path.join(_BASE_DIR, "static", "diagram.png")


DEFAULT_DIAGRAM_IMAGE: Final[str] = _resolve_diagram_image()

DEFAULT_SBC_DATA_DIR: Final[str] = os.environ.get(
    "SBC_DIR", os.path.join(_BASE_DIR, "dados-sbcs")
)

# =============================================================================
# CAMPOS DO FORMULÁRIO
# =============================================================================
BOOLEAN_FIELDS: Final[tuple[str, ...]] = (
    "csp", "servicos_especiais", "cng", "scm", "av", "projeto_padrao",
    "sbc_ativo", "ip_reservado", "vivo_reserva",
    "operadora_ciente", "lcr_nacional", "white_list",
    "prefixos_liberados_abr", "premissas_ok",
)

TEXT_FIELDS: Final[tuple[str, ...]] = (
    "nome_operadora", "rn1", "atendimento", "redes", "qual", "tmr",
    "responsavel_operadora", "responsavel_vivo", "asn",
    "responsavel_infra", "aprovado_por",
    "status", "escopo_text",
    "responsavel_atacado", "responsavel_engenharia",
)

JSON_FIELDS: Final[tuple[str, ...]] = (
    "escopo_flags_json", "dados_vivo_json",
    "dados_operadora_json", "engenharia_params_json",
)

# Tipos de tráfego oficiais (planilha "No de A Formato de Entrega por Rota").
# Para alterar a lista ou os textos, edite traffic_types.py.
_LEGACY_SCOPE_FLAGS: Final[tuple[str, ...]] = (
    # Tipos antigos — mantidos apenas para não invalidar PTIs já salvos.
    # Não aparecem mais como opção no formulário.
    "LC", "LD15 + CNG", "LDS/CSP + CNG", "Transporte", "VC1", "Concentração",
)

ALLOWED_SCOPE_FLAGS: Final[frozenset[str]] = frozenset(
    TRAFFIC_TYPE_NAMES + _LEGACY_SCOPE_FLAGS
)

# =============================================================================
# SBC
# =============================================================================
SBC_CACHE_TTL_SECONDS: Final[int] = 300

SBC_STATUS_TO_HEALTH: Final[dict[str, str]] = {
    "normal":   "disponivel",
    "atenção":  "moderado",
    "atencao":  "moderado",
    "crítico":  "critico",
    "critico":  "critico",
}

SBC_STATUS_BASE_SCORE: Final[dict[str, int]] = {
    "disponivel": 70,
    "moderado":   45,
    "critico":    15,
}

# =============================================================================
# IPAM
# =============================================================================
# ATENÇÃO: em produção, credenciais devem vir exclusivamente de variáveis
# de ambiente. Os padrões abaixo existem para desenvolvimento local.
IPAM_BASE_URL: Final[str] = os.getenv("IPAM_BASE_URL", "http://10.113.144.242")
IPAM_USER: Final[str] = os.getenv("IPAM_USER", "40418843")
IPAM_PASS: Final[str] = os.getenv("IPAM_PASS", "230581Bs.@@")
IPAM_SECTION_ID: Final[int] = 186
IPAM_REQUEST_TIMEOUT: Final[int] = 15  # segundos

IPAM_MASK_POOL_MAP: Final[dict[str, str]] = {
    "24": "POOL 1",
    "28": "POOL 3",
    "29": "POOL 4",
}
IPAM_DEFAULT_POOL: Final[str] = "POOL 5"

# =============================================================================
# CNs — SEED E METADATA
# =============================================================================
CN_SEED_RAW: Final[str] = """
68 82 97 92 96 77 75 74 73 71 88 85 61 27 28 64 62 61 98 99
34 37 31 35 32 38 33 67 66 65 91 94 93 83 81 87 86 89
43 44 45 46 41 24 22 21 84 69 95 51 53 54 55 47 48 49 79
18 14 15 16 13 19 17 11 12 63
"""

CN_METADATA: Final[dict[str, tuple[str, str]]] = {
    "11": ("São Paulo", "SP"),          "12": ("São José dos Campos", "SP"),
    "13": ("Santos", "SP"),             "14": ("Bauru", "SP"),
    "15": ("Sorocaba", "SP"),           "16": ("Ribeirão Preto", "SP"),
    "17": ("São José do Rio Preto", "SP"), "18": ("Presidente Prudente", "SP"),
    "19": ("Campinas", "SP"),
    "21": ("Rio de Janeiro", "RJ"),     "22": ("Campos dos Goytacazes", "RJ"),
    "24": ("Volta Redonda", "RJ"),      "27": ("Vitória", "ES"),
    "28": ("Cachoeiro de Itapemirim", "ES"),
    "31": ("Belo Horizonte", "MG"),     "32": ("Juiz de Fora", "MG"),
    "33": ("Governador Valadares", "MG"), "34": ("Uberlândia", "MG"),
    "35": ("Poços de Caldas", "MG"),    "37": ("Divinópolis", "MG"),
    "38": ("Montes Claros", "MG"),
    "41": ("Curitiba", "PR"),           "42": ("Ponta Grossa", "PR"),
    "43": ("Londrina", "PR"),           "44": ("Maringá", "PR"),
    "45": ("Foz do Iguaçu", "PR"),      "46": ("Francisco Beltrão", "PR"),
    "47": ("Joinville", "SC"),          "48": ("Florianópolis", "SC"),
    "49": ("Chapecó", "SC"),
    "51": ("Porto Alegre", "RS"),       "53": ("Pelotas", "RS"),
    "54": ("Caxias do Sul", "RS"),      "55": ("Santa Maria", "RS"),
    "61": ("Brasília", "DF"),           "62": ("Goiânia", "GO"),
    "63": ("Palmas", "TO"),             "64": ("Rio Verde", "GO"),
    "65": ("Cuiabá", "MT"),             "66": ("Rondonópolis", "MT"),
    "67": ("Campo Grande", "MS"),
    "71": ("Salvador", "BA"),           "73": ("Ilhéus", "BA"),
    "74": ("Juazeiro", "BA"),           "75": ("Feira de Santana", "BA"),
    "77": ("Vitória da Conquista", "BA"),
    "79": ("Aracaju", "SE"),            "81": ("Recife", "PE"),
    "82": ("Maceió", "AL"),             "83": ("João Pessoa", "PB"),
    "84": ("Natal", "RN"),              "85": ("Fortaleza", "CE"),
    "86": ("Teresina", "PI"),           "87": ("Petrolina", "PE"),
    "88": ("Juazeiro do Norte", "CE"),  "89": ("Picos", "PI"),
    "91": ("Belém", "PA"),              "92": ("Manaus", "AM"),
    "93": ("Santarém", "PA"),           "94": ("Marabá", "PA"),
    "95": ("Boa Vista", "RR"),          "96": ("Macapá", "AP"),
    "97": ("Coari", "AM"),
    "98": ("São Luís", "MA"),           "99": ("Imperatriz", "MA"),
    "68": ("Rio Branco", "AC"),         "69": ("Porto Velho", "RO"),
}

# =============================================================================
# UF → REGIONAL E VIZINHOS
# =============================================================================
UF_TO_REGIONAL: Final[dict[str, str]] = {
    "AC": "NORTE",  "AM": "NORTE",  "AP": "NORTE",  "PA": "NORTE",
    "RO": "NORTE",  "RR": "NORTE",  "TO": "NORTE",
    "AL": "NORDESTE", "BA": "NORDESTE", "CE": "NORDESTE",
    "MA": "NORDESTE", "PB": "NORDESTE", "PE": "NORDESTE",
    "PI": "NORDESTE", "RN": "NORDESTE", "SE": "NORDESTE",
    "DF": "CENTRO-OESTE", "GO": "CENTRO-OESTE",
    "MS": "CENTRO-OESTE", "MT": "CENTRO-OESTE",
    "ES": "SUDESTE", "MG": "SUDESTE", "RJ": "SUDESTE", "SP": "SUDESTE",
    "PR": "SUL",    "RS": "SUL",    "SC": "SUL",
}

UF_NEIGHBORS: Final[dict[str, list[str]]] = {
    "AC": ["RO", "AM"],
    "AL": ["PE", "SE", "BA"],
    "AM": ["PA", "RR", "AC", "RO", "MT"],
    "AP": ["PA"],
    "BA": ["SE", "AL", "PE", "PI", "MG", "GO", "TO", "MA"],
    "CE": ["RN", "PB", "PE", "PI"],
    "DF": ["GO", "MG"],
    "ES": ["MG", "RJ", "BA"],
    "GO": ["DF", "MG", "MS", "MT", "TO", "BA"],
    "MA": ["PI", "TO", "PA"],
    "MG": ["SP", "RJ", "ES", "BA", "GO", "DF", "MS"],
    "MS": ["PR", "SP", "MG", "GO", "MT"],
    "MT": ["MS", "GO", "TO", "PA", "AM", "RO"],
    "PA": ["MA", "TO", "MT", "AM", "AP", "RR"],
    "PB": ["PE", "RN", "CE"],
    "PE": ["PB", "AL", "BA", "CE", "PI"],
    "PI": ["MA", "CE", "PE", "BA", "TO"],
    "PR": ["SP", "SC", "MS"],
    "RJ": ["SP", "MG", "ES"],
    "RN": ["PB", "CE"],
    "RO": ["MT", "AM", "AC"],
    "RR": ["AM", "PA"],
    "RS": ["SC"],
    "SC": ["PR", "RS"],
    "SE": ["AL", "BA"],
    "SP": ["RJ", "MG", "PR", "MS"],
    "TO": ["MA", "PI", "BA", "GO", "MT", "PA"],
}
