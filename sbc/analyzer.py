"""sbc/analyzer.py — Motor de análise e recomendação de SBCs."""
import logging
import os
from collections import defaultdict
from dataclasses import asdict
from typing import Optional

from config import (
    CN_METADATA,
    SBC_STATUS_BASE_SCORE,
    SBC_STATUS_TO_HEALTH,
    UF_NEIGHBORS,
    UF_TO_REGIONAL,
)
from sbc.models import SBCAnalysisResult, SBCMeasurement, SBCSuggestionResponse
from sbc.processor import SBCDataProcessor

logger = logging.getLogger(__name__)


class SBCAnalyzer:
    """Motor de análise e recomendação de SBCs."""

    def __init__(self, processor: SBCDataProcessor) -> None:
        self.processor = processor

    # ------------------------------------------------------------------
    # Classificação de saúde e status
    # ------------------------------------------------------------------
    def _classify_health(self, status: str) -> str:
        key = status.strip().lower()
        result = SBC_STATUS_TO_HEALTH.get(key)
        if result:
            return result
        if "crít" in key or "crit" in key:
            return "critico"
        if "aten" in key:
            return "moderado"
        if "norm" in key or "ok" in key or "disp" in key:
            return "disponivel"
        return "moderado"

    def _worst_status(self, statuses: list[str]) -> str:
        severity_map: dict[str, int] = {
            "normal": 0,
            "atenção": 1, "atencao": 1, "atencion": 1,
            "crítico": 2, "critico": 2,
        }
        canonical = {0: "Normal", 1: "Atenção", 2: "Crítico"}
        worst_score = -1
        for s in statuses:
            key = s.strip().lower()
            s_score = severity_map.get(key, -1)
            if s_score < 0:
                if "crít" in key or "crit" in key:
                    s_score = 2
                elif "aten" in key:
                    s_score = 1
                else:
                    s_score = 0
            if s_score > worst_score:
                worst_score = s_score
        return canonical.get(max(worst_score, 0), "Normal")

    # ------------------------------------------------------------------
    # Métricas
    # ------------------------------------------------------------------
    def _calc_tendencia(self, caps_values: list[int]) -> str:
        if len(caps_values) < 2:
            return "estavel"
        mid = len(caps_values) // 2
        avg_first = sum(caps_values[:mid]) / mid if mid else 0
        avg_second = (
            sum(caps_values[mid:]) / len(caps_values[mid:])
            if caps_values[mid:] else 0
        )
        if avg_first == 0:
            return "estavel"
        variation = (avg_second - avg_first) / avg_first
        if variation > 0.10:
            return "subindo"
        if variation < -0.10:
            return "descendo"
        return "estavel"

    def _calculate_score(
        self,
        saude: str,
        tendencia: str,
        caps_values: list[int],
        num_servicos: int,
        total_medicoes: int,
    ) -> int:
        base = SBC_STATUS_BASE_SCORE.get(saude, 45)
        bonus = 0
        if caps_values:
            mn, mx = min(caps_values), max(caps_values)
            avg = sum(caps_values) / len(caps_values)
            if avg > 0 and (mx - mn) / avg < 0.10:
                bonus += 10
        if total_medicoes >= 5:
            bonus += 5
        if num_servicos > 1:
            bonus += 3
        penalty = 10 if tendencia == "subindo" else 0
        return max(0, min(100, int(base + bonus - penalty)))

    def _generate_reason(self, result: SBCAnalysisResult) -> str:
        labels = {
            "disponivel": "capacidade disponível",
            "moderado":   "ocupação moderada, monitorar",
            "critico":    "capacidade crítica, evitar",
        }
        parts = [
            f"Status XLSX: {result.status_fonte} — {labels.get(result.saude, result.saude)}",
            (
                f"CAPS avg {result.caps_avg} "
                f"(mín {result.caps_min}, máx {result.caps_max}, "
                f"último {result.caps_ultimo})"
            ),
        ]
        if result.caps_tendencia and result.caps_tendencia != "estavel":
            icon = "📈" if result.caps_tendencia == "subindo" else "📉"
            parts.append(f"Tendência: {icon} {result.caps_tendencia}")
        else:
            parts.append("Tendência: → estável")
        if result.total_medicoes > 1:
            parts.append(f"{result.total_medicoes} medições")
        if len(result.servicos) > 1:
            parts.append(f"Serviços: {', '.join(result.servicos)}")
        if result.cidade:
            parts.append(f"Local: {result.cidade}/{result.uf}")
        return " | ".join(parts)

    # ------------------------------------------------------------------
    # Agregação
    # ------------------------------------------------------------------
    def _aggregate_sbc_measurements(
        self, measurements: list[SBCMeasurement]
    ) -> list[SBCAnalysisResult]:
        groups: dict[str, list[SBCMeasurement]] = defaultdict(list)
        for m in measurements:
            groups[m.sbc].append(m)

        results: list[SBCAnalysisResult] = []
        for sbc_name, sbc_m in groups.items():
            caps_values = [m.caps for m in sbc_m]
            caps_avg = sum(caps_values) / len(caps_values) if caps_values else 0
            caps_tendencia = self._calc_tendencia(caps_values)
            worst_status = self._worst_status([m.status for m in sbc_m])
            saude = self._classify_health(worst_status)
            servicos = list(dict.fromkeys(m.servico for m in sbc_m if m.servico))
            first = sbc_m[0]
            dias = [m.dia for m in sbc_m if m.dia]
            score = self._calculate_score(
                saude, caps_tendencia, caps_values, len(servicos), len(sbc_m)
            )
            result = SBCAnalysisResult(
                nome=sbc_name,
                cidade=first.cidade, uf=first.uf, regional=first.regional,
                modelo=first.modelo, servicos=servicos,
                caps_avg=round(caps_avg, 1),
                caps_max=max(caps_values) if caps_values else 0,
                caps_min=min(caps_values) if caps_values else 0,
                caps_ultimo=caps_values[-1] if caps_values else 0,
                caps_tendencia=caps_tendencia,
                total_medicoes=len(sbc_m),
                status_fonte=worst_status, saude=saude, score=score,
                recomendado=False, motivo="",
                responsavel=next((m.responsavel for m in sbc_m if m.responsavel), ""),
                prazo=next((m.prazo for m in sbc_m if m.prazo), ""),
                dia_mais_recente=dias[-1] if dias else "",
                semana=first.semana,
            )
            results.append(result)

        results.sort(key=lambda r: r.score, reverse=True)
        for i, r in enumerate(results):
            r.recomendado = (i == 0)
            r.motivo = self._generate_reason(r)
        return results

    # ------------------------------------------------------------------
    # API pública
    # ------------------------------------------------------------------
    def suggest_for_cn(self, cn: str) -> SBCSuggestionResponse:
        cn = str(cn).strip().zfill(2)
        meta = CN_METADATA.get(cn)
        if not meta:
            return SBCSuggestionResponse(
                cn=cn, uf="", cidade="", regional="",
                source_file="", source_modificado_em="", data_medicao="",
                total_sbcs=0, sbcs=[], fallback_usado=False, fallback_origem="",
                mensagem=f"CN {cn} não encontrado no cadastro.",
            )
        cidade, uf = meta
        result = self.suggest_for_uf(uf)
        result.cn = cn
        result.cidade = cidade
        return result

    def suggest_for_uf(self, uf: str) -> SBCSuggestionResponse:
        uf = uf.strip().upper()
        if len(uf) != 2:
            return SBCSuggestionResponse(
                cn="", uf=uf, cidade="", regional="",
                source_file="", source_modificado_em="", data_medicao="",
                total_sbcs=0, sbcs=[], fallback_usado=False, fallback_origem="",
                mensagem=f"UF inválida: '{uf}'.",
            )

        regional = UF_TO_REGIONAL.get(uf, "")
        cidade = next(
            (c for c, u in CN_METADATA.values() if u == uf), uf
        )
        cn_found = next(
            (code for code, (_, u) in CN_METADATA.items() if u == uf), ""
        )

        measurements, _ = self.processor.get_data()
        file_info = self.processor.get_file_info()

        if not measurements:
            return SBCSuggestionResponse(
                cn=cn_found, uf=uf, cidade=cidade, regional=regional,
                source_file=file_info.get("filename", ""),
                source_modificado_em=file_info.get("modified_at", ""),
                data_medicao="", total_sbcs=0, sbcs=[],
                fallback_usado=False, fallback_origem="",
                mensagem="Nenhum dado de SBC disponível.",
            )

        uf_measurements = [m for m in measurements if m.uf == uf]
        fallback_usado = False
        fallback_origem = ""

        if not uf_measurements:
            for neighbor_uf in UF_NEIGHBORS.get(uf, []):
                uf_measurements = [m for m in measurements if m.uf == neighbor_uf]
                if uf_measurements:
                    fallback_usado = True
                    fallback_origem = f"UF vizinha: {neighbor_uf}"
                    break

        if not uf_measurements and regional:
            regional_ufs = [u for u, r in UF_TO_REGIONAL.items() if r == regional]
            uf_measurements = [m for m in measurements if m.uf in regional_ufs]
            if uf_measurements:
                fallback_usado = True
                fallback_origem = f"Regional: {regional}"

        if not uf_measurements:
            return SBCSuggestionResponse(
                cn=cn_found, uf=uf, cidade=cidade, regional=regional,
                source_file=file_info.get("filename", ""),
                source_modificado_em=file_info.get("modified_at", ""),
                data_medicao="", total_sbcs=0, sbcs=[],
                fallback_usado=False, fallback_origem="",
                mensagem=f"Nenhum SBC para {uf}, vizinhos ou regional {regional}.",
            )

        all_results = self._aggregate_sbc_measurements(uf_measurements)
        nao_criticos = sorted(
            [r for r in all_results if r.saude != "critico"],
            key=lambda r: r.score, reverse=True,
        )
        criticos = sorted(
            [r for r in all_results if r.saude == "critico"],
            key=lambda r: r.score, reverse=True,
        )
        ordered = nao_criticos + criticos

        for i, r in enumerate(ordered):
            r.recomendado = (i == 0)
            r.motivo = self._generate_reason(r)

        n_disp = len(nao_criticos)
        n_crit = len(criticos)
        msg_parts = [f"{len(ordered)} SBC(s) para {uf} ({cidade})"]
        if n_disp:
            msg_parts.append(f"{n_disp} disponível(is)")
        if n_crit:
            msg_parts.append(f"{n_crit} crítico(s)")
        if fallback_usado:
            msg_parts.append(f"Dados de {fallback_origem}")

        return SBCSuggestionResponse(
            cn=cn_found, uf=uf, cidade=cidade, regional=regional,
            source_file=file_info.get("filename", ""),
            source_modificado_em=file_info.get("modified_at", ""),
            data_medicao=all_results[0].dia_mais_recente if all_results else "",
            total_sbcs=len(ordered),
            sbcs=[asdict(r) for r in ordered],
            fallback_usado=fallback_usado, fallback_origem=fallback_origem,
            todos_criticos=(n_disp == 0 and n_crit > 0),
            total_disponiveis=n_disp, total_criticos=n_crit,
            mensagem=" · ".join(msg_parts),
        )

    def get_overview(self) -> dict:
        measurements, aggregated = self.processor.get_data()
        file_info = self.processor.get_file_info()
        if not measurements:
            return {
                "file_info": file_info, "total_sbcs": 0,
                "por_regional": {}, "por_saude": {},
                "por_status_fonte": {}, "dados_agregados": len(aggregated),
            }
        all_results = self._aggregate_sbc_measurements(measurements)
        por_regional: dict[str, int] = defaultdict(int)
        por_saude:    dict[str, int] = defaultdict(int)
        por_status:   dict[str, int] = defaultdict(int)
        for r in all_results:
            por_regional[r.regional] += 1
            por_saude[r.saude] += 1
            por_status[r.status_fonte] += 1
        return {
            "file_info": file_info, "total_sbcs": len(all_results),
            "por_regional": dict(por_regional), "por_saude": dict(por_saude),
            "por_status_fonte": dict(por_status),
            "dados_agregados": len(aggregated),
            "sbcs": [asdict(r) for r in all_results],
        }

    def health_check(self) -> dict:
        file_info = self.processor.get_file_info()
        file_path = self.processor.find_latest_file()
        return {
            "data_dir_exists": os.path.isdir(self.processor.data_dir),
            "data_dir": self.processor.data_dir,
            "latest_file": os.path.basename(file_path) if file_path else None,
            "file_info": file_info,
            "status": "ok" if file_path else "no_file_found",
        }
