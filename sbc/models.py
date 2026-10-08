"""sbc/models.py — Dataclasses de dados SBC."""
from dataclasses import dataclass


@dataclass
class SBCMeasurement:
    """Uma única medição de um SBC (uma linha do XLSX)."""
    semana: str
    dia: str
    cidade: str
    uf: str
    regional: str
    sbc: str
    caps: int
    status: str
    modelo: str
    servico: str
    responsavel: str
    prazo: str


@dataclass
class SBCAnalysisResult:
    """Resultado da análise de um SBC individual."""
    nome: str
    cidade: str
    uf: str
    regional: str
    modelo: str
    servicos: list[str]
    caps_avg: float
    caps_max: int
    caps_min: int
    caps_ultimo: int
    caps_tendencia: str
    total_medicoes: int
    status_fonte: str
    saude: str
    score: int
    recomendado: bool
    motivo: str
    responsavel: str
    prazo: str
    dia_mais_recente: str
    semana: str


@dataclass
class SBCSuggestionResponse:
    """Resposta completa da API de sugestão de SBC."""
    cn: str
    uf: str
    cidade: str
    regional: str
    source_file: str
    source_modificado_em: str
    data_medicao: str
    total_sbcs: int
    sbcs: list[dict]
    fallback_usado: bool
    fallback_origem: str
    mensagem: str
    todos_criticos: bool = False
    total_disponiveis: int = 0
    total_criticos: int = 0
