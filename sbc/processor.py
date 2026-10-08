"""sbc/processor.py — Lê, limpa e cacheia dados de SBC a partir de planilhas XLSX."""
import glob
import logging
import os
import time
from datetime import datetime
from typing import Optional

from config import SBC_CACHE_TTL_SECONDS
from sbc.models import SBCMeasurement

logger = logging.getLogger(__name__)

# Importação opcional
try:
    from openpyxl import load_workbook
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


class SBCDataProcessor:
    """
    Lê, limpa e cacheia dados de SBC a partir de planilhas XLSX.

    Estrutura esperada:
      SEMANA | DIA | CIDADE | UF | REGIONAL | SBC | CAPS | STATUS |
      MOD/FORNEC | SERVIÇO | RESPONSÁVEL | PRAZO
    """

    COLUMN_MAP: dict[str, str] = {
        "SEMANA": "semana", "ANO SEMANA": "semana",
        "DIA": "dia",
        "CIDADE": "cidade", "UF": "uf", "REGIONAL": "regional",
        "SBC": "sbc", "CAPS": "caps",
        "ST": "status", "STATUS": "status",
        "MOD/ FORNEC": "modelo", "MOD/FORNEC": "modelo",
        "MOD / FORNEC": "modelo", "MODELO": "modelo",
        "FORNECEDOR": "modelo", "FABRICANTE": "modelo",
        "SERVIÇO": "servico", "SERVICO": "servico",
        "RESPONSÁVEL": "responsavel", "RESPONSAVEL": "responsavel",
        "PRAZO": "prazo",
    }

    def __init__(self, data_dir: str, cache_ttl: int = SBC_CACHE_TTL_SECONDS) -> None:
        self.data_dir = data_dir
        self.cache_ttl = cache_ttl
        self._cache_data: list[SBCMeasurement] = []
        self._cache_aggregated: list[dict] = []
        self._cache_timestamp: float = 0.0
        self._cache_file_path: str = ""
        self._cache_file_mtime: float = 0.0

    # ------------------------------------------------------------------
    # Arquivo
    # ------------------------------------------------------------------
    def find_latest_file(self) -> Optional[str]:
        if not os.path.isdir(self.data_dir):
            logger.warning("Diretório de SBC não encontrado: %s", self.data_dir)
            return None
        files: list[str] = []
        for pattern in ("*.xlsx", "*.XLSX", "*.xls", "*.XLS"):
            files.extend(glob.glob(os.path.join(self.data_dir, pattern)))
        if not files:
            logger.warning("Nenhum XLSX encontrado em: %s", self.data_dir)
            return None
        files.sort(key=lambda f: os.path.getmtime(f), reverse=True)
        latest = files[0]
        logger.info(
            "Arquivo SBC mais recente: %s (%d total)",
            os.path.basename(latest), len(files),
        )
        return latest

    # ------------------------------------------------------------------
    # Leitura e normalização
    # ------------------------------------------------------------------
    def _normalize_column_name(self, raw_name: str) -> str:
        clean = raw_name.strip().upper()
        return self.COLUMN_MAP.get(
            clean, clean.lower().replace(" ", "_").replace("/", "_")
        )

    def _parse_int(self, value, default: int = 0) -> int:
        try:
            cleaned = str(value).strip().replace(",", ".").replace(" ", "")
            return int(float(cleaned)) if cleaned else default
        except (ValueError, TypeError):
            return default

    def _read_xlsx(
        self, filepath: str
    ) -> tuple[list[SBCMeasurement], list[dict]]:
        measurements: list[SBCMeasurement] = []
        aggregated: list[dict] = []

        if not OPENPYXL_AVAILABLE:
            logger.error("openpyxl não disponível.")
            return [], []

        try:
            wb = load_workbook(filepath, read_only=True, data_only=True)
            ws = wb.active
            rows_iter = ws.iter_rows(values_only=True)

            try:
                header_raw = next(rows_iter)
            except StopIteration:
                logger.error("XLSX vazio: %s", filepath)
                wb.close()
                return [], []

            headers = [
                self._normalize_column_name(
                    str(v).strip() if v is not None else ""
                )
                for v in header_raw
            ]
            logger.info("Colunas detectadas: %s", headers)

            required = {
                "semana", "dia", "cidade", "uf", "regional", "sbc",
                "caps", "status", "modelo", "servico", "responsavel", "prazo",
            }
            missing = required - set(headers)
            if missing:
                logger.warning("Colunas ausentes: %s", sorted(missing))

            for row_num, row_values in enumerate(rows_iter, start=2):
                row: dict[str, str] = {}
                for idx, cell_val in enumerate(row_values):
                    if idx < len(headers):
                        row[headers[idx]] = (
                            str(cell_val).strip() if cell_val is not None else ""
                        )

                sbc_name = row.get("sbc", "").strip()
                uf = row.get("uf", "").strip().upper()

                if sbc_name and uf:
                    measurements.append(SBCMeasurement(
                        semana=row.get("semana", ""),
                        dia=row.get("dia", ""),
                        cidade=row.get("cidade", "").strip().title(),
                        uf=uf,
                        sbc=sbc_name,
                        regional=row.get("regional", "").strip().upper(),
                        caps=self._parse_int(row.get("caps", "0")),
                        status=row.get("status", "Normal").strip(),
                        modelo=row.get("modelo", "").strip(),
                        servico=row.get("servico", "").strip(),
                        responsavel=row.get("responsavel", "").strip(),
                        prazo=row.get("prazo", "").strip(),
                    ))
                else:
                    aggregated.append({
                        "row_num": row_num, "raw_data": row,
                        "nota": "Linha sem SBC/UF (dado agregado ou total)",
                    })

            wb.close()
            logger.info(
                "XLSX lido: %d medições, %d agregadas",
                len(measurements), len(aggregated),
            )

        except FileNotFoundError:
            logger.error("XLSX não encontrado: %s", filepath)
        except Exception:
            logger.exception("Erro ao ler XLSX %s", filepath)

        return measurements, aggregated

    # ------------------------------------------------------------------
    # Cache
    # ------------------------------------------------------------------
    def get_data(
        self, force_reload: bool = False
    ) -> tuple[list[SBCMeasurement], list[dict]]:
        now = time.time()
        cache_expired = (now - self._cache_timestamp) > self.cache_ttl

        if force_reload or not self._cache_data or cache_expired:
            latest_file = self.find_latest_file()
            if not latest_file:
                return [], []

            current_mtime = os.path.getmtime(latest_file)
            if (
                latest_file == self._cache_file_path
                and current_mtime == self._cache_file_mtime
                and self._cache_data
                and not force_reload
            ):
                self._cache_timestamp = now
                return self._cache_data, self._cache_aggregated

            logger.info("Carregando do disco: %s", os.path.basename(latest_file))
            ext = os.path.splitext(latest_file)[1].lower()
            if ext in (".xlsx", ".xls"):
                measurements, aggregated = self._read_xlsx(latest_file)
            else:
                logger.warning("Formato não suportado: %s", ext)
                return [], []

            self._cache_data = measurements
            self._cache_aggregated = aggregated
            self._cache_timestamp = now
            self._cache_file_path = latest_file
            self._cache_file_mtime = current_mtime

        return self._cache_data, self._cache_aggregated

    def get_file_info(self) -> dict:
        if self._cache_file_path and os.path.exists(self._cache_file_path):
            mtime = datetime.fromtimestamp(self._cache_file_mtime)
            return {
                "filename": os.path.basename(self._cache_file_path),
                "modified_at": mtime.strftime("%d/%m/%Y %H:%M"),
                "total_measurements": len(self._cache_data),
                "total_aggregated": len(self._cache_aggregated),
                "cache_age_seconds": int(time.time() - self._cache_timestamp),
            }
        return {"filename": None, "modified_at": None, "total_measurements": 0}
