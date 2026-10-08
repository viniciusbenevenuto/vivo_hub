"""excel/builder.py — Builder do Workbook Excel do PTI."""
import json
import logging
import os
import re
from collections import defaultdict
from datetime import datetime
from typing import Optional

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

try:
    from PIL import Image as PILImage
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

from config import (
    DEFAULT_DIAGRAM_IMAGE,
    IPAM_DEFAULT_POOL,
    IPAM_MASK_POOL_MAP,
    IPAM_PASS,
    IPAM_USER,
    MAX_TABLE_ROWS,
    UF_NEIGHBORS,
    UF_TO_REGIONAL,
)
from excel.image_utils import (
    fit_size_keep_aspect,
    parse_xy_percent,
    render_labels_on_image,
)
from excel.styles import ExcelStylePalette
from helpers import parse_db_datetime, row_get
from ipam import IPAMClient
from params_programacao import NAO_RESPONDIDO, PARAMS_SECOES
from traffic_types import TRAFFIC_BY_NAME

logger = logging.getLogger(__name__)

# SBC VIVO dos PTIs gravados na época do SipRouter SP (antes da escolha
# do SBC pela capacidade). Usado só quando há bloco_ip e nenhum SBC escolhido.
SIPROUTER_SBC_LEGADO = "CDSIP_SPO_PL/CDSIP_SPO_JG"


class PTIWorkbookBuilder:
    """Builder para Workbook Excel do PTI."""

    def __init__(self, form_row, sbc_analyzer=None, all_versions=None) -> None:
        if not PIL_AVAILABLE:
            raise RuntimeError("Pillow não instalado.")

        self.form = form_row
        self.sbc_analyzer = sbc_analyzer
        self.all_versions = all_versions or [form_row]
        self.s = ExcelStylePalette()
        self.wb = Workbook()
        self.nome = (row_get(form_row, "nome_operadora") or "Operadora").strip()
        self.resp_atk = (
            row_get(form_row, "responsavel_atacado")
            or row_get(form_row, "responsavel_vivo")
            or ""
        ).strip()
        self.resp_eng = (row_get(form_row, "responsavel_engenharia", "") or "").strip()
        self.asn_op   = (row_get(form_row, "asn", "") or "").strip()

        cdt = parse_db_datetime(row_get(form_row, "created_at", ""))
        self.created_str = cdt.strftime("%d/%m/%Y") if cdt else ""

        self.traffic      = self._extract_traffic()
        self.escopo_text  = (row_get(form_row, "escopo_text", "") or "").strip()

        if self.traffic:
            flags_str = " | ".join(self.traffic)
            self.escopo_completo = (
                f"{self.escopo_text} [{flags_str}]" if self.escopo_text else flags_str
            )
        else:
            self.escopo_completo = self.escopo_text

        self.vivo_rows = self._extract_vivo_rows()
        self.op_rows   = self._extract_op_rows()

        # Resultado das reservas IPAM — preenchido por reserva_redes_ipam(),
        # chamada no início de build().
        self.ipam_reservas:   dict[str, str] = {}
        self.ipam_falhas:     list[str] = []
        self.ipam_tentativas: int = 0   # quantos blocos tentou reservar (não pulados)

        cns_vivo = self._unique_field(self.vivo_rows, "cn")
        cns_op   = self._unique_field(self.op_rows,   "cn")
        self.cns_unicos = list(dict.fromkeys(cns_vivo + cns_op))

        areas_vivo = self._unique_field(self.vivo_rows, "cidade")
        areas_op   = self._unique_field(self.op_rows,   "cidade")
        self.areas_locais = list(dict.fromkeys(areas_vivo + areas_op))

        self.rn1              = (row_get(form_row, "rn1", "") or "").strip()
        self.csp              = bool(int(row_get(form_row, "csp", 0) or 0))
        self.cng              = bool(int(row_get(form_row, "cng", 0) or 0))
        self.servicos_especiais = bool(int(row_get(form_row, "servicos_especiais", 0) or 0))

        all_rows = self.vivo_rows + self.op_rows
        eot_lc = [r.get("eto_lc", "") for r in all_rows if r.get("eto_lc", "").strip()]
        eot_ld = [r.get("eot_ld", "") for r in all_rows if r.get("eot_ld", "").strip()]
        self.eot_local = ", ".join(eot_lc) if eot_lc else ""
        self.eot_ld    = ", ".join(eot_ld) if eot_ld else ""

        # Banda / Canais / CAPS informados pela Engenharia (seção 8)
        self.trafego_metricas = self._extract_metricas()

    def _extract_metricas(self) -> dict:
        """{tipo de tráfego: {banda_exist, banda_plan, canais_*, caps_*}}."""
        raw = row_get(self.form, "trafego_metricas_json", "{}")
        try:
            parsed = json.loads(raw) if isinstance(raw, str) else raw
            return parsed if isinstance(parsed, dict) else {}
        except (json.JSONDecodeError, TypeError):
            return {}

    # =========================================================================
    # RESERVA IPAM
    # =========================================================================
    def reserva_redes_ipam(self) -> None:
        """Reserva um bloco IP no phpIPAM para cada CN preenchido no PTI."""
        vivo_rows = self.vivo_rows
        if not IPAM_USER or not IPAM_PASS:
            logger.warning("Credenciais IPAM não configuradas. Reserva ignorada.")
            self.ipam_falhas.append("Credenciais IPAM não configuradas — reserva não realizada.")
            return
        rn1_val   = (row_get(self.form, "rn1", "") or "").strip()
        nome_op   = (row_get(self.form, "nome_operadora", "") or "Operadora").strip()
        num = max(len(vivo_rows), len(self.op_rows))
        logger.info("IPAM | Iniciando reserva: %d bloco(s) a processar", num)
        for i in range(num):
            vr        = vivo_rows[i]    if i < len(vivo_rows)    else {}
            opr       = self.op_rows[i] if i < len(self.op_rows) else {}
            cn_usuario = vr.get("cn", "") or opr.get("cn", "")
            mask_alvo  = opr.get("mask", "") or vr.get("mask", "")
            if not cn_usuario or not mask_alvo:
                logger.warning("IPAM | Bloco %d pulado — cn=%r mask_op=%r mask_vr=%r",
                               i, cn_usuario, opr.get("mask",""), vr.get("mask",""))
                continue
            # Descrição: RN1-XXXX - Nome da Operadora
            descricao = f"{rn1_val} - {nome_op}" if rn1_val else nome_op
            nome_pool = IPAM_MASK_POOL_MAP.get(mask_alvo, IPAM_DEFAULT_POOL)
            logger.info("Reserva IPAM | CN: %s | /%s | Pool: %s | Desc: %s",
                        cn_usuario, mask_alvo, nome_pool, descricao)
            self.ipam_tentativas += 1
            cliente = IPAMClient(IPAM_USER, IPAM_PASS)
            try:
                cliente.autenticar()
                id_cn   = cliente.buscar_id_cn(cn_usuario)
                id_pool = cliente.buscar_id_pool(id_cn, nome_pool)
                cliente.buscar_rede_pai(id_cn, id_pool, mask_alvo)
                nova_rede = cliente.reservar_rede(id_pool, mask_alvo, descricao)
                # nova_rede é um dict ou string retornado pela API; extrair o CIDR
                if isinstance(nova_rede, dict):
                    subnet  = nova_rede.get("subnet", "") or nova_rede.get("network", "") or ""
                    mask_nr = nova_rede.get("mask", "") or nova_rede.get("prefix", "") or mask_alvo
                    cidr    = f"{subnet}/{mask_nr}" if subnet else ""
                else:
                    cidr = str(nova_rede).strip() if nova_rede else ""
                if cidr and cn_usuario not in self.ipam_reservas:
                    self.ipam_reservas[cn_usuario] = cidr
                logger.info("Reserva concluída: %s → CIDR armazenado: %s", nova_rede, cidr)
            except (ValueError, RuntimeError) as exc:
                logger.error("Falha IPAM CN '%s': %s", cn_usuario, exc)
                self.ipam_falhas.append(f"CN {cn_usuario}: {exc}")

    # =========================================================================
    # EXTRAÇÃO DE DADOS
    # =========================================================================
    def _extract_vivo_rows(self) -> list[dict]:
        # 'mask' mantido como campo legado para compatibilidade com dados antigos
        # (o campo foi movido para Dados Operadora, mas registros antigos ainda têm mask aqui)
        keys = ("ref", "data", "escopo", "localidade", "cn", "bloco_ip", "id_vivo",
                "mask", "endereco_link", "cidade", "uf", "lat", "long",
                "sbc_vivo")   # SBC VIVO escolhido pela capacidade
        return self._extract_rows("dados_vivo_json", keys)

    def _extract_op_rows(self) -> list[dict]:
        keys = ("ref", "localidade", "eto_lc", "eot_ld", "cn", "sbc", "faixa_ip",
                "mask", "id_op", "endereco_link", "cidade", "uf", "lat", "long")
        return self._extract_rows("dados_operadora_json", keys)

    def _extract_rows(self, field: str, keys: tuple) -> list[dict]:
        raw = row_get(self.form, field, "[]")
        try:
            d = json.loads(raw) if isinstance(raw, str) else raw
            if not isinstance(d, list):
                return []
            rows = []
            for item in d:
                if not isinstance(item, dict):
                    continue
                row = {k: str(item.get(k, "")).strip() for k in keys}
                if any(row.values()):
                    rows.append(row)
            return rows[:MAX_TABLE_ROWS]
        except Exception:
            return []

    def _extract_traffic(self) -> list[str]:
        raw = row_get(self.form, "escopo_flags_json", "[]")
        try:
            p = json.loads(raw) if isinstance(raw, str) else raw
            return [str(x).strip() for x in (p if isinstance(p, list) else []) if str(x).strip()]
        except Exception:
            return []

    def _unique_field(self, rows: list[dict], field: str) -> list[str]:
        seen: list[str] = []
        for r in rows:
            v = r.get(field, "").strip()
            if v and v not in seen:
                seen.append(v)
        return seen

    # =========================================================================
    # HELPERS DE PLANILHA
    # =========================================================================
    def _set_column_widths(self, ws, widths: dict) -> None:
        for col, width in widths.items():
            ws.column_dimensions[get_column_letter(col)].width = width

    # ------------------------------------------------------------------
    # BORDAS
    # Regra do PTI: títulos, cabeçalhos e a tabela inteira levam contorno
    # EXTERNO grosso (medium); as divisões internas ficam finas (thin).
    # ------------------------------------------------------------------
    @staticmethod
    def _set_side(cell, *, top=None, bottom=None, left=None, right=None) -> None:
        """Sobrescreve apenas os lados informados, preservando os demais."""
        b = cell.border
        cell.border = Border(
            top=top or b.top, bottom=bottom or b.bottom,
            left=left or b.left, right=right or b.right,
        )

    def _thin_box(self, ws, r1: int, c1: int, r2: int, c2: int) -> None:
        """Preenche todo o intervalo com bordas finas."""
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(row=r, column=c).border = self.s.box_border

    def _thick_outline(self, ws, r1: int, c1: int, r2: int, c2: int) -> None:
        """Aplica contorno externo grosso ao intervalo, sem tocar no interior."""
        m = Side(style="medium", color="000000")
        for c in range(c1, c2 + 1):
            self._set_side(ws.cell(row=r1, column=c), top=m)
            self._set_side(ws.cell(row=r2, column=c), bottom=m)
        for r in range(r1, r2 + 1):
            self._set_side(ws.cell(row=r, column=c1), left=m)
            self._set_side(ws.cell(row=r, column=c2), right=m)

    def _write_brand_header(self, ws, row: int, sc: int, ec: int) -> None:
        """Faixa superior "VIVO — <operadora>" presente em todas as abas."""
        ws.merge_cells(start_row=row, start_column=sc, end_row=row, end_column=ec)
        c = ws.cell(row=row, column=sc, value=f"VIVO — {self.nome}")
        c.font = self.s.font_title
        c.alignment = self.s.align_center
        c.fill = self.s.fill_primary
        self._thick_outline(ws, row, sc, row, ec)   # título: contorno grosso
        ws.row_dimensions[row].height = 24.0

    def _write_section_title(self, ws, row: int, sc: int, ec: int, title: str) -> None:
        ws.merge_cells(start_row=row, start_column=sc, end_row=row, end_column=ec)
        c = ws.cell(row=row, column=sc, value=title)
        c.font = self.s.font_subtitle
        c.alignment = self.s.align_center
        c.fill = self.s.fill_secondary
        self._thick_outline(ws, row, sc, row, ec)   # título: contorno grosso
        ws.row_dimensions[row].height = 25.0

    def _setup_page(self, ws, *, ls: bool = False) -> None:
        """Configura impressão A4 (retrato por padrão; ls=True → paisagem)."""
        ws.sheet_view.showGridLines = False
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = (
            ws.ORIENTATION_LANDSCAPE if ls else ws.ORIENTATION_PORTRAIT
        )
        ws.print_options.horizontalCentered = True
        ws.page_margins.left = ws.page_margins.right = 0.4
        ws.page_margins.top  = ws.page_margins.bottom = 0.5

    def _diagram_cfg(self) -> dict:
        f = self.form

        def _fl(k, d):
            try:
                return float(row_get(f, k, d))
            except (ValueError, TypeError):
                return d

        def _it(k, d):
            try:
                return int(float(row_get(f, k, d)))
            except (ValueError, TypeError):
                return d

        ns = str(row_get(f, "roteador_op_no_stroke", "")).lower() in ("1", "true", "yes", "sim")
        rs = 0.0 if ns else 0.009
        ov = row_get(f, "roteador_op_stroke_width_pct", None)
        if ov not in (None, ""):
            try:
                rs = float(ov)
            except (ValueError, TypeError):
                pass

        return {
            "img":  row_get(f, "diagram_image_path") or DEFAULT_DIAGRAM_IMAGE,
            "vxy":  parse_xy_percent(row_get(f, "vivo_label_xy_pct", ""),      (0.16, 0.48)),
            "oxy":  parse_xy_percent(row_get(f, "operadora_label_xy_pct", ""), (0.83, 0.44)),
            "lvxy": parse_xy_percent(row_get(f, "link_vivo_label_xy_pct", ""), (0.49, 0.53)),
            "loxy": parse_xy_percent(row_get(f, "link_op_label_xy_pct", ""),   (0.49, 0.35)),
            "fsz":  _fl("diagram_font_size_pct",       0.040),
            "lfsz": _fl("link_labels_font_size_pct",   0.030),
            "fp":   row_get(f, "diagram_font_path", None),
            "r1":   parse_xy_percent(row_get(f, "roteador_op1_label_xy_pct", ""), (0.64, 0.58)),
            "r2":   parse_xy_percent(row_get(f, "roteador_op2_label_xy_pct", ""), (0.64, 0.39)),
            "rfp":  _fl("roteador_op_font_size_pct", 0.010),
            "rsp":  rs,
            "bw":   _it("diagram_max_w",            1900),
            "bh":   _it("diagram_max_h",            650),
            "aro":  _it("diagram_anchor_row_offset", -3),
            "aco":  _it("diagram_anchor_col_offset",  1),
        }

    # =========================================================================
    # HELPERS DE CÁLCULO DE IP (SipRouter)
    # =========================================================================
    @staticmethod
    def _parse_cidr(block: str) -> tuple[list[str], str] | tuple[None, None]:
        """'10.11.130.16/28' → (['10','11','130','16'], '28')"""
        block = (block or "").strip()
        if "/" not in block:
            return None, None
        ip, mask = block.split("/", 1)
        parts = ip.strip().split(".")
        if len(parts) != 4:
            return None, None
        return parts, mask.strip()

    @staticmethod
    def _add_last_octet(block: str, offset: int) -> str:
        """'10.11.130.16/28' + 4 → '10.11.130.20/28'"""
        parts, mask = PTIWorkbookBuilder._parse_cidr(block)
        if parts is None:
            return block
        try:
            new_last = int(parts[3]) + offset
        except ValueError:
            return block
        return f"{'.'.join(parts[:3])}.{new_last}/{mask}"

    @staticmethod
    def _cidr_to_netmask(mask_str: str) -> str:
        """'28' → '/28'"""
        try:
            return f"/{int(mask_str)}"
        except (ValueError, TypeError):
            return "/28"

    @staticmethod
    def _ip_only(block: str) -> str:
        """'10.11.130.20/28' → '10.11.130.20'"""
        return block.split("/")[0].strip() if block and "/" in block else block

    @staticmethod
    def _extract_pl_jg(bloco_ip: str) -> tuple[str, str]:
        """
        'PL:10.11.130.16/28 / JG:10.11.131.16/28' → ('10.11.130.16/28', '10.11.131.16/28')
        Aceita também formato sem prefixo (retorna mesmo valor para ambos).
        """
        pl = jg = ""
        if not bloco_ip:
            return pl, jg
        for part in bloco_ip.split(" / "):
            part = part.strip()
            if part.upper().startswith("PL:"):
                pl = part[3:].strip()
            elif part.upper().startswith("JG:"):
                jg = part[3:].strip()
        # Fallback: sem prefixo PL/JG → usar o mesmo bloco para ambos
        if not pl and not jg and bloco_ip.strip():
            pl = jg = bloco_ip.strip()
        return pl, jg

    @staticmethod
    def _sbc_vivo(vivo_row: dict) -> str:
        """SBC VIVO da interconexão (Ponta A).

        O SBC é escolhido pela capacidade no formulário da Engenharia. PTIs
        antigos, gravados quando o bloco vinha do SipRouter SP (bloco_ip
        preenchido e sem SBC escolhido), continuam com CDSIP_SPO_PL/JG.
        """
        sbc = (vivo_row.get("sbc_vivo", "") or "").strip()
        if sbc:
            return sbc
        if (vivo_row.get("bloco_ip", "") or "").strip():
            return SIPROUTER_SBC_LEGADO
        return ""

    # =========================================================================
    # MATCHING DE LINHAS
    # =========================================================================
    def _find_matching_op_row(self, cn: str, index: int) -> Optional[dict]:
        if cn:
            for op in self.op_rows:
                if op.get("cn", "") == cn:
                    return op
        return self.op_rows[index] if index < len(self.op_rows) else None

    def _find_matching_vivo_row(self, cn: str, index: int) -> dict:
        if cn:
            for vr in self.vivo_rows:
                if vr.get("cn", "") == cn:
                    return vr
        return self.vivo_rows[index] if index < len(self.vivo_rows) else {}

    # =========================================================================
    # MAPEAMENTO DE TRÁFEGO
    # =========================================================================
    def _resolve_traffic_mapping(
        self, traf_type: str, operadora: str, cn_vivo: str, cn_op: str
    ) -> dict:
        """Informações padrão de um tipo de tráfego (traffic_types.py).

        Tipos oficiais vêm da tabela; qualquer outro valor (PTIs antigos,
        texto livre) cai na heurística legada logo abaixo.
        """
        official = TRAFFIC_BY_NAME.get(traf_type.strip())
        if official:
            return {
                "formato_no_a": official["formato_no_a"],
                "enc_ab":       official["enc_ab"],
                "enc_ba":       official["enc_ba"],
                "codec":        official["codec"],
                "obs":          official["obs"],
            }
        legacy = self._resolve_traffic_mapping_legacy(
            traf_type, operadora, cn_vivo, cn_op
        )
        # Garante as chaves novas mesmo nos tipos antigos
        return {"formato_no_a": "", "obs": "SIP-I", **legacy}

    # -------------------------------------------------------------------------
    # Heurística antiga — usada só para tipos de tráfego fora da tabela oficial
    # (PTIs criados antes da atualização). Mantida para não perder o conteúdo
    # dos formulários já salvos.
    # -------------------------------------------------------------------------
    def _resolve_traffic_mapping_legacy(
        self, traf_type: str, operadora: str, cn_vivo: str, cn_op: str
    ) -> dict:
        t       = traf_type.strip()
        t_upper = t.upper()
        op = operadora or "Operadora"
        cv = cn_vivo or "CN"
        co = cn_op   or "CN"

        csp_match  = re.search(r'CSP\s*(\d{2})',    t, re.IGNORECASE)
        rota_match = re.search(r'Rota\s+LD\s*(\d{2})', t, re.IGNORECASE)
        xy = (
            csp_match.group(1) if csp_match
            else (rota_match.group(1) if rota_match else "XY")
        )
        sem_csp = bool(re.search(r's/\s*CSP', t, re.IGNORECASE))

        # LC+TR / LC com AL
        if "LC+TR" in t_upper or ("LC" in t_upper and "TR" in t_upper and "AL" in t_upper):
            return {"enc_ab": f"LC: (9090) PREF-MCDU {op} do CN {cv}",
                    "enc_ba": f"LC: (9090) PREF-MCDU VIVO do CN {cv}", "codec": "G711 / G729"}

        # LC simples
        if t_upper.strip() in ("LC",) or (
            "LC" in t_upper and "LD" not in t_upper and "TR" not in t_upper
            and "CSP" not in t_upper and "CNG" not in t_upper
        ):
            return {"enc_ab": f"LC: (9090) PREF-MCDU {op} do CN {cv}",
                    "enc_ba": f"LC: (9090) PREF-MCDU VIVO do CN {cv}", "codec": "G711 / G729"}

        # LD15 + CNG (VIVO)
        if "CSP" in t_upper and xy == "15" and "CNG" in t_upper and "VIVO" in t_upper:
            return {"enc_ab": f"LD: (9)01511 PREF-MCDU - NUM B do CN {cv}",
                    "enc_ba": (f"LD: (9) 015 {co} PREF-MCDU / LDI: 0015 (***)\n"
                               f"CNG VIVO: 08XX 03XX 05XX 09XX (10-11 dig) (***)\n"
                               f"SE: 10315/10615"), "codec": "G711 / G729"}

        # LD15 + CNG genérico
        if ("LD15" in t_upper or ("LD" in t_upper and "15" in t_upper)) and "CNG" in t_upper:
            return {"enc_ab": f"LD: (9)01511 PREF-MCDU - NUM B do CN {cv}",
                    "enc_ba": (f"LD: (9) 015 {co} PREF-MCDU / LDI: 0015 (***)\n"
                               f"CNG VIVO: 08XX 03XX 05XX 09XX (10-11 dig)"),
                    "codec": "G711 / G729"}

        # LDS/CSP + CNG
        if "CSP" in t_upper and "CNG" in t_upper and "TRANSP" not in t_upper and not sem_csp:
            return {"enc_ab": (f"LD: (9)0 {xy} {co} PREF-MCDU / LDI: 00 {xy} (***)\n"
                               f"CNG VIVO: 08XX 03XX 05XX 09XX (10-11 dig) (***)\n"
                               f"SE: 103{xy} / 106{xy}"),
                    "enc_ba": f"LD: (9)0+{xy}+{co} PREF-MCDU - NUM B do CN {cv}",
                    "codec": "G711 / G729"}

        # LD s/ CSP
        if "LD" in t_upper and "S/" in t_upper and "CSP" in t_upper and "TRANSP" not in t_upper:
            return {"enc_ab": "CNG: 08XX 03XX 05XX 09XX (10-11 dig)",
                    "enc_ba": f"LD: (9)0+{co} PREF-MCDU - NUM B do CN {cv}", "codec": "G711 / G729"}

        # Transporte + CSP + CNG
        if "TRANSP" in t_upper and "CSP" in t_upper and "CNG" in t_upper:
            return {"enc_ab": "CNG: 08XX 03XX 05XX 09XX (10-11 dig) (**)",
                    "enc_ba": f"LD: (9)0+{xy}+{co} PREF-MCDU - No. B diferente do CN {cv}",
                    "codec": "G711 / G729"}

        if "TRANSP" in t_upper and "LD" in t_upper and "S/" in t_upper:
            return {"enc_ab": "CNG: 08XX 03XX 05XX 09XX (10-11 dig) (**)",
                    "enc_ba": f"LD: (9)0+{co} PREF-MCDU No. B diferente do CN {cv}",
                    "codec": "G711 / G729"}

        # Transporte simples
        if "TRANSP" in t_upper:
            return {"enc_ab": f"Transporte de {op} — qualquer CN de origem",
                    "enc_ba": f"Transporte VIVO — qualquer CN de destino",
                    "codec": "G711 / G729"}

        # Concentração
        if "CONCENTR" in t_upper:
            return {"enc_ab": f"CONC: {cv} — encaminhamento concentrado para CN {cv}",
                    "enc_ba": f"CONC: {co} — encaminhamento concentrado para CN {co}",
                    "codec": "G711 / G729"}

        # VC1 com Rota LD
        if "VC1" in t_upper and "ROTA" in t_upper and "LD" in t_upper:
            return {"enc_ab": f"LD: (9) 0 {xy} {cv} PREF-MCDU do CN {cv} No de B VIVO",
                    "enc_ba": f"LD: (9) 0 {xy} {co} PREF-MCDU - CN{cv}  No de B VIVO",
                    "codec": "G-711 / AMR"}

        # VC1 simples
        if "VC1" in t_upper:
            return {"enc_ab": f"LC: (9) 0{cv} PREF-MCDU - CN {cv}",
                    "enc_ba": f"LC: (9) 0{co} PREF-MCDU - CN{cv} - SE: 1058",
                    "codec": "G-711 / AMR"}

        return {"enc_ab": "", "enc_ba": "", "codec": ""}

    # =========================================================================
    # SBC RESOLVER (via dados XLSX)
    # =========================================================================
    def _resolve_sbc_for_uf(self, uf: str, cidade: str = "") -> Optional[dict]:
        if not self.sbc_analyzer or not uf:
            return None
        uf = uf.strip().upper()
        try:
            measurements, _ = self.sbc_analyzer.processor.get_data()
        except Exception:
            return None
        if not measurements:
            return None

        uf_m = [m for m in measurements if m.uf == uf]
        fallback_label = ""

        if not uf_m:
            for n_uf in UF_NEIGHBORS.get(uf, []):
                uf_m = [m for m in measurements if m.uf == n_uf]
                if uf_m:
                    fallback_label = f"(vizinho: {n_uf})"
                    break

        if not uf_m:
            regional = UF_TO_REGIONAL.get(uf, "")
            if regional:
                regional_ufs = [u for u, r in UF_TO_REGIONAL.items() if r == regional]
                uf_m = [m for m in measurements if m.uf in regional_ufs]
                if uf_m:
                    fallback_label = f"(regional: {regional})"

        if not uf_m:
            return None

        groups: dict[str, list] = defaultdict(list)
        for m in uf_m:
            groups[m.sbc].append(m)

        cidade_norm = cidade.strip().upper() if cidade else ""
        if cidade_norm:
            for sbc_name, ms in groups.items():
                for m in ms:
                    if m.cidade.strip().upper() == cidade_norm:
                        caps = [x.caps for x in ms]
                        return {
                            "nome": sbc_name, "cidade": ms[0].cidade, "uf": ms[0].uf,
                            "caps_avg": round(sum(caps) / len(caps), 1),
                            "status": ms[0].status, "modelo": ms[0].modelo,
                            "total_medicoes": len(ms), "fallback": fallback_label,
                        }

        best = max(groups, key=lambda k: len(groups[k]))
        ms = groups[best]
        caps = [x.caps for x in ms]
        return {
            "nome": best, "cidade": ms[0].cidade, "uf": ms[0].uf,
            "caps_avg": round(sum(caps) / len(caps), 1),
            "status": ms[0].status, "modelo": ms[0].modelo,
            "total_medicoes": len(ms), "fallback": fallback_label,
        }

    # =========================================================================
    # ABA: ÍNDICE
    # =========================================================================
    def _build_index_sheet(self) -> None:
        s  = self.s
        ws = self.wb.active
        ws.title = "Índice"

        for c in range(1, 10):
            ws.column_dimensions[get_column_letter(c)].width = 12.0
        ws.column_dimensions["C"].width = 8.0
        ws.column_dimensions["D"].width = 45.0

        self._write_brand_header(ws, 2, 3, 8)
        ws.row_dimensions[2].height = 28.0

        ws.merge_cells(start_row=4, start_column=3, end_row=4, end_column=8)
        c = ws.cell(row=4, column=3, value="ANEXO 3: PROJETO TÉCNICO")
        c.font = s.font_subtitle; c.alignment = s.align_center; c.fill = s.fill_secondary

        ws.merge_cells(start_row=5, start_column=3, end_row=5, end_column=8)
        c = ws.cell(row=5, column=3,
                    value="PROJETO DE INTERLIGAÇÃO PARA ENCAMINHAMENTO DA TERMINAÇÃO DE CHAMADAS DE VOZ")
        c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c.alignment = s.align_center; c.fill = s.fill_secondary
        ws.row_dimensions[5].height = 30.0
        # Bloco de título (linhas 4–5): contorno externo grosso
        self._thin_box(ws, 4, 3, 5, 8)
        self._thick_outline(ws, 4, 3, 5, 8)

        items = [
            "1. Versões", "2. Projeto de Interligação",
            "2.2. Diagrama de Interligação",
            "2.3. Características do Projeto de Interligação e do Plano de Encaminhamento",
            "2.4. Plano de Contingência", "2.5. Concentração",
            "2.6. Plan NUM_Operadora", "2.7. Dados de MTL",
            "2.8. SE REG III (Vivo STFC Concessionária)", "2.9. Parâmetros de Programação",
        ]
        for i, t in enumerate(items):
            r = 7 + i
            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
            c = ws.cell(row=r, column=3, value=t)
            c.font = Font(name="Calibri", size=11)
            c.alignment = s.align_left; c.fill = s.alt_fill(i)
            ws.row_dimensions[r].height = 22.0
            self._thin_box(ws, r, 3, r, 8)
        # Tabela de conteúdo (itens): contorno externo grosso
        self._thick_outline(ws, 7, 3, 6 + len(items), 8)

        ws.freeze_panes = "C7"
        self._setup_page(ws)

    # =========================================================================
    # ABA: VERSÕES
    # =========================================================================
    def _build_versions_sheet(self) -> None:
        s  = self.s
        ws = self.wb.create_sheet(title="Versões")
        self._set_column_widths(ws, {1:2,2:2,3:8,4:12,5:28,6:28,7:28,8:10,9:18,10:10,11:2})
        self._write_brand_header(ws, 2, 3, 10)
        self._write_section_title(ws, 4, 3, 10, "CONTROLE DE VERSÕES DO PTI")

        for i, t in enumerate([
            "Versão", "Data", "Responsável Eng de ITX",
            "Responsável Gestão de ITX", "Escopo", "CN", "ÁREAS LOCAIS", "ATA",
        ]):
            c = ws.cell(row=6, column=3 + i, value=t)
            c.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
            c.alignment = s.align_center; c.fill = s.fill_light; c.border = s.box_border
        ws.row_dimensions[6].height = 25.0
        self._thick_outline(ws, 6, 3, 6, 10)   # cabeçalho: contorno grosso

        fd = {
            4: self.created_str, 5: self.resp_eng, 6: self.resp_atk,
            7: self.escopo_completo,
            8: ", ".join(self.cns_unicos)   if self.cns_unicos   else "",
            9: ", ".join(self.areas_locais) if self.areas_locais else "",
        }

        for v_idx, v_form in enumerate(self.all_versions):
            r = 7 + v_idx; ws.row_dimensions[r].height = 22.0; f = s.alt_fill(v_idx)
            ver_num = int(v_form["version"] or 1) if v_form["version"] else (v_idx + 1)
            v_cdt = parse_db_datetime(row_get(v_form, "created_at", ""))
            v_date = v_cdt.strftime("%d/%m/%Y") if v_cdt else ""
            v_escopo = (row_get(v_form, "escopo_text", "") or "").strip()
            try:
                flags = json.loads(row_get(v_form, "escopo_flags_json", "[]") or "[]")
                if flags:
                    v_escopo = (v_escopo + " [" + " | ".join(flags) + "]") if v_escopo else " | ".join(flags)
            except (json.JSONDecodeError, TypeError):
                pass
            v_resp_eng = (row_get(v_form, "responsavel_engenharia", "") or "").strip()
            v_resp_atk = (row_get(v_form, "responsavel_atacado", "") or row_get(v_form, "responsavel_vivo", "") or "").strip()
            # CNs desta versão
            try:
                vr = json.loads(row_get(v_form, "dados_vivo_json", "[]") or "[]")
                or_ = json.loads(row_get(v_form, "dados_operadora_json", "[]") or "[]")
                v_cns = ", ".join({r.get("cn","") for r in (vr + or_) if r.get("cn","")})
            except (json.JSONDecodeError, TypeError, AttributeError):
                v_cns = ""
            vd = {3: ver_num, 4: v_date, 5: v_resp_eng, 6: v_resp_atk, 7: v_escopo, 8: v_cns, 9: "", 10: ""}
            for col, val in vd.items():
                cell = ws.cell(row=r, column=col, value=val)
                cell.font = s.font_body; cell.fill = f; cell.border = s.box_border
                cell.alignment = (
                    Alignment(horizontal="center", vertical="center", wrap_text=True)
                    if col in (7, 9) else s.align_center
                )
        # Contorno externo grosso da tabela inteira (cabeçalho + versões)
        self._thick_outline(ws, 6, 3, 6 + len(self.all_versions), 10)
        ws.freeze_panes = "C7"
        self._setup_page(ws, ls=True)

    # =========================================================================
    # ABA: DIAGRAMA
    # =========================================================================
    def _build_diagram_sheet(self) -> None:
        s   = self.s
        ws  = self.wb.create_sheet(title="Diagrama de Interligação")
        cfg = self._diagram_cfg()
        n   = self.nome

        # Um bloco por interconexão — total = maior lado (VIVO ou Operadora).
        num_blocks     = max(1, len(self.vivo_rows), len(self.op_rows))
        LC, RC, MID    = 3, 14, 9

        col_widths = {1: 2.0, 2: 2.0, 15: 2.0}
        for c in range(LC, RC + 1):
            col_widths[c] = 12.0
        self._set_column_widths(ws, col_widths)

        # Cabeçalho global
        self._write_brand_header(ws, 2, LC, RC)
        for r_row, t, bold in [
            (4,  f"Diagrama de Interligação entre a VIVO e a {n}", True),
            (6,  "2.2.1 DIAGRAMAÇÃO DO PROJETO SIP", True),
            (7,  "2.2.1.1 Anúncio de Redes pelos 2 Links SPC", True),
            (8,  f"A {n} abordará os endereços da VIVO conforme abaixo.", False),
            (9,  f"A VIVO abordará os endereços da {n} conforme abaixo.", False),
            (11, "2.2.1.2 Parâmetros de Configuração do Link", True),
        ]:
            ws.merge_cells(start_row=r_row, start_column=LC, end_row=r_row, end_column=RC)
            c = ws.cell(row=r_row, column=LC, value=t)
            c.font = Font(name="Calibri", size=11 if bold else 10, bold=bold)
            c.alignment = s.align_left if bold else s.align_wrap
            ws.row_dimensions[r_row].height = 22.0 if bold else 25.0

        for o, label, val in [(0, "ASN VIVO", "10429 (Público)"), (1, f"ASN {n}", self.asn_op)]:
            rr = 12 + o
            ws.merge_cells(start_row=rr, start_column=LC, end_row=rr, end_column=LC+2)
            ws.cell(row=rr, column=LC, value=label).font = Font(name="Calibri", size=10, bold=True)
            ws.cell(row=rr, column=LC).alignment = s.align_center
            ws.merge_cells(start_row=rr, start_column=LC+3, end_row=rr, end_column=LC+5)
            ws.cell(row=rr, column=LC+3, value=val).font = s.font_body
            ws.cell(row=rr, column=LC+3).alignment = s.align_center

        traffic_items = [t for t in self.traffic if t]
        num_traffic   = max(len(traffic_items), 3)
        cursor        = 15

        for b_idx in range(num_blocks):
            # Pareia por índice; se um lado não tem esta linha, casa pelo CN.
            vivo_row = self.vivo_rows[b_idx] if b_idx < len(self.vivo_rows) else {}
            op_row   = self.op_rows[b_idx]   if b_idx < len(self.op_rows)   else {}
            if not vivo_row and op_row:
                vivo_row = self._find_matching_vivo_row(op_row.get("cn", ""), b_idx) or {}
            if not op_row and vivo_row:
                op_row = self._find_matching_op_row(vivo_row.get("cn", ""), b_idx) or {}

            cn_val        = vivo_row.get("cn", "") or (op_row.get("cn", "") if op_row else "")
            bloco_ip_vivo = vivo_row.get("bloco_ip", "") or ""
            pl_block, jg_block = self._extract_pl_jg(bloco_ip_vivo)
            # SBC escolhido pela capacidade (vazio em PTIs da época do SipRouter)
            sbc_vivo      = (vivo_row.get("sbc_vivo", "") or "").strip()
            faixa_ip_op   = (op_row.get("faixa_ip", "") if op_row else "") or ""

            def _gap(row, start, end):
                for gc in range(start, end):
                    ws.cell(row=row, column=gc).fill   = s.fill_white
                    ws.cell(row=row, column=gc).border = s.no_border

            # Linha CN + VRF
            r = cursor
            cn_label = f"CN {cn_val}"  # sem cidade
            ws.merge_cells(start_row=r, start_column=LC, end_row=r, end_column=LC+4)
            c = ws.cell(row=r, column=LC, value=cn_label if cn_val else "CN ___")
            c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            c.alignment = s.align_center; c.fill = s.fill_light
            ws.merge_cells(start_row=r, start_column=MID, end_row=r, end_column=RC)
            c2 = ws.cell(row=r, column=MID, value="VRF: __________________________")
            c2.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            c2.alignment = s.align_center; c2.fill = s.fill_light
            ws.row_dimensions[r].height = 25.0
            _gap(r, LC+5, MID)
            cursor += 2

            # Ponta A / B
            r = cursor
            for sc, ec, label, valor in [
                (LC, LC+4, "Ponta A", "VIVO"), (MID, RC, "Ponta B", n)
            ]:
                ws.merge_cells(start_row=r, start_column=sc, end_row=r, end_column=ec)
                ws.cell(row=r, column=sc, value=label).font = Font(name="Calibri", size=10, bold=True)
                ws.cell(row=r, column=sc).alignment = s.align_center
                ws.cell(row=r, column=sc).fill = s.fill_neutral
                ws.merge_cells(start_row=r+1, start_column=sc, end_row=r+1, end_column=ec)
                ws.cell(row=r+1, column=sc, value=valor).font = Font(name="Calibri", size=10, bold=True)
                ws.cell(row=r+1, column=sc).alignment = s.align_center
            ws.row_dimensions[r].height = ws.row_dimensions[r+1].height = 22.0
            for rr in (r, r+1):
                _gap(rr, LC+5, MID)
            cursor += 3

            # ── Cabeçalhos de tráfego ──────────────────────────────────────
            r = cursor
            # Título tabela esquerda: SBC VIVO escolhido (legado: CDSIP_SPO_PL)
            if sbc_vivo:
                titulo_vivo = f"SBC VIVO: {sbc_vivo}"
            elif pl_block:
                titulo_vivo = "CDSIP_SPO_PL (SipRouter)"
            else:
                titulo_vivo = "SBC VIVO"
            ws.merge_cells(start_row=r, start_column=LC, end_row=r, end_column=LC+2)
            c = ws.cell(row=r, column=LC, value=titulo_vivo)
            c.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
            c.alignment = s.align_center; c.fill = s.fill_accent
            # Título tabela direita: Rede Reservada IPAM
            ws.merge_cells(start_row=r, start_column=MID, end_row=r, end_column=MID+2)
            c2 = ws.cell(row=r, column=MID, value=f"Rede Reservada IPAM ({n})")
            c2.font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
            c2.alignment = s.align_center; c2.fill = s.fill_secondary
            _gap(r, LC+3, MID)
            ws.row_dimensions[r].height = 20.0
            cursor += 1

            r = cursor
            for lc_start in (LC, MID):
                for i, h in enumerate(["Tráfego", "Endereço IP", "NETMASK"]):
                    c = ws.cell(row=r, column=lc_start+i, value=h)
                    c.font = s.font_subheader; c.alignment = s.align_center
                    c.fill = s.fill_light; c.border = s.box_border
            ws.row_dimensions[r].height = 25.0
            _gap(r, LC+3, MID)
            cursor += 1

            # ── Calcular IPs ───────────────────────────────────────────────
            # Esquerda (SipRouter PL): bloco + 4 base, depois +1 por tráfego
            # Direita  (Rede Reservada IPAM): rede reservada + 2 por tráfego
            #   Base: self.ipam_reservas[cn_val] se disponível, senão faixa_ip_op
            ipam_rede   = self.ipam_reservas.get(cn_val, "") if cn_val else ""
            ip_op_base  = ipam_rede or faixa_ip_op   # nunca usar endereco_link como IP
            _, mask_ipam = self._parse_cidr(ip_op_base)
            netmask_op  = self._cidr_to_netmask(mask_ipam) if mask_ipam else "/28"

            _, pl_mask  = self._parse_cidr(pl_block)
            _, jg_mask  = self._parse_cidr(jg_block)
            netmask_pl  = self._cidr_to_netmask(pl_mask) if pl_mask else "/28"
            netmask_jg  = self._cidr_to_netmask(jg_mask) if jg_mask else "/28"

            num_visible    = len(traffic_items)
            data_start_row = cursor

            for j in range(num_traffic):
                r = cursor + j
                ws.row_dimensions[r].height = 22.0
                f_fill = s.alt_fill(j)
                traf_val = traffic_items[j] if j < num_visible else ""
                if j >= num_visible:
                    ws.row_dimensions[r].hidden = True

                # Lado esquerdo — PL
                pl_ip = (self._ip_only(self._add_last_octet(pl_block, 4 + j))
                         if pl_block else "")
                ws.cell(row=r, column=LC, value=traf_val).font = s.font_small
                ws.cell(row=r, column=LC).fill = f_fill
                ws.cell(row=r, column=LC).alignment = s.align_center
                ws.cell(row=r, column=LC).border = s.box_border

                ip_c = ws.cell(row=r, column=LC+1, value="" if num_visible > 1 else pl_ip)
                ip_c.font = s.font_small; ip_c.fill = f_fill
                ip_c.alignment = s.align_center; ip_c.border = s.box_border

                mk_c = ws.cell(row=r, column=LC+2, value="" if num_visible > 1 else netmask_pl)
                mk_c.font = s.font_small; mk_c.fill = f_fill
                mk_c.alignment = s.align_center; mk_c.border = s.box_border

                # Lado direito — Rede Reservada IPAM: +2 no primeiro, depois +1 por tráfego
                # j=0 → base+2, j=1 → base+3, j=2 → base+4 ...
                op_ip_val = (self._ip_only(self._add_last_octet(ip_op_base, 2 + j))
                             if ip_op_base else "")
                ws.cell(row=r, column=MID, value=traf_val).font = s.font_small
                ws.cell(row=r, column=MID).fill = f_fill
                ws.cell(row=r, column=MID).alignment = s.align_center
                ws.cell(row=r, column=MID).border = s.box_border

                op_ip_c = ws.cell(row=r, column=MID+1,
                                  value="" if num_visible > 1 else op_ip_val)
                op_ip_c.font = s.font_small; op_ip_c.fill = f_fill
                op_ip_c.alignment = s.align_center; op_ip_c.border = s.box_border

                op_mk_c = ws.cell(row=r, column=MID+2,
                                  value="" if num_visible > 1 else netmask_op)
                op_mk_c.font = s.font_small; op_mk_c.fill = f_fill
                op_mk_c.alignment = s.align_center; op_mk_c.border = s.box_border

                _gap(r, LC+3, MID)

            # Mesclar IP/Mask quando há múltiplos tráfegos
            if num_visible > 1:
                last_vis = data_start_row + num_visible - 1
                for j in range(num_visible):
                    pl_ip = self._ip_only(self._add_last_octet(pl_block, 4 + j)) if pl_block else ""
                    r_row = data_start_row + j
                    ws.cell(row=r_row, column=LC+1, value=pl_ip).font = s.font_small
                    ws.cell(row=r_row, column=LC+1).alignment = s.align_center
                    ws.cell(row=r_row, column=LC+1).border = s.box_border
                # Máscara PL: mesclar todas as linhas de tráfego em uma célula só
                if num_visible > 1:
                    ws.merge_cells(start_row=data_start_row, start_column=LC+2,
                                   end_row=last_vis, end_column=LC+2)
                ws.cell(row=data_start_row, column=LC+2, value=netmask_pl).font = s.font_small
                ws.cell(row=data_start_row, column=LC+2).alignment = s.align_center
                ws.cell(row=data_start_row, column=LC+2).border = s.box_border

                # Rede Reservada IPAM: +2 no primeiro, +1 nos seguintes (2+j)
                for j in range(num_visible):
                    op_ip_val = (self._ip_only(self._add_last_octet(ip_op_base, 2 + j))
                                 if ip_op_base else "")
                    r_row = data_start_row + j
                    ws.cell(row=r_row, column=MID+1, value=op_ip_val).font = s.font_small
                    ws.cell(row=r_row, column=MID+1).alignment = s.align_center
                    ws.cell(row=r_row, column=MID+1).border = s.box_border
                # Máscara OP: mesclar todas as linhas de tráfego em uma célula só
                if num_visible > 1:
                    ws.merge_cells(start_row=data_start_row, start_column=MID+2,
                                   end_row=last_vis, end_column=MID+2)
                ws.cell(row=data_start_row, column=MID+2, value=netmask_op).font = s.font_small
                ws.cell(row=data_start_row, column=MID+2).alignment = s.align_center
                ws.cell(row=data_start_row, column=MID+2).border = s.box_border

            cursor += num_traffic

            # ── Segunda tabela: CDSIP_SPO_JG ──────────────────────────────
            if jg_block:
                cursor += 1

                # Cabeçalho JG
                r = cursor
                for i, h in enumerate(["Tráfego", "Endereço IP (JG)", "NETMASK"]):
                    c = ws.cell(row=r, column=LC+i, value=h)
                    c.font = s.font_subheader; c.alignment = s.align_center
                    c.fill = s.fill_light; c.border = s.box_border
                ws.row_dimensions[r].height = 22.0
                _gap(r, LC+3, MID)
                cursor += 1

                for j in range(num_visible if num_visible else 1):
                    r = cursor + j
                    ws.row_dimensions[r].height = 20.0
                    f_fill = s.alt_fill(j)
                    traf_val = traffic_items[j] if j < num_visible else ""
                    jg_ip = (self._ip_only(self._add_last_octet(jg_block, 4 + j))
                             if jg_block else "")
                    for col, val in [(LC, traf_val), (LC+1, jg_ip)]:
                        c = ws.cell(row=r, column=col, value=val)
                        c.font = s.font_small; c.fill = f_fill
                        c.alignment = s.align_center; c.border = s.box_border
                    _gap(r, LC+3, MID)

                # Máscara JG: célula mesclada cobrindo todos os tráfegos
                jg_data_start = cursor
                jg_data_end   = cursor + max(num_visible, 1) - 1
                if num_visible > 1:
                    ws.merge_cells(start_row=jg_data_start, start_column=LC+2,
                                   end_row=jg_data_end, end_column=LC+2)
                c = ws.cell(row=jg_data_start, column=LC+2, value=netmask_jg)
                c.font = s.font_small; c.fill = s.alt_fill(0)
                c.alignment = s.align_center; c.border = s.box_border

                cursor += max(num_visible, 1)

            # Imagem
            img_path = cfg["img"]
            logger.info("Diagrama | verificando imagem: %s | existe: %s",
                        img_path, os.path.exists(img_path))
            if os.path.exists(img_path):
                cursor += 4
                img_anchor_row = cursor
                sbc_op_img = (op_row.get("sbc", "") if op_row else "") or ""
                row_labels = [
                    {"text": f"ROTEADOR {n}", "xy_pct": (0.65, 0.36), "font_size_pct": cfg["rfp"], "stroke_width_pct": cfg["rsp"]},
                    {"text": f"ROTEADOR {n}", "xy_pct": (0.65, 0.55), "font_size_pct": cfg["rfp"], "stroke_width_pct": cfg["rsp"]},
                    {"text": "LINK RESP VIVO", "xy_pct": (0.50, 0.70), "font_size_pct": cfg["lfsz"], "stroke_width_pct": 0.006},
                    {"text": f"LINK RESP {n}", "xy_pct": (0.50, 0.18), "font_size_pct": cfg["lfsz"], "stroke_width_pct": 0.006},
                    {"text": "ROTEADOR VIVO", "xy_pct": (0.34, 0.36), "font_size_pct": 0.010, "stroke_width_pct": 0.006},
                    {"text": "ROTEADOR VIVO", "xy_pct": (0.34, 0.55), "font_size_pct": 0.010, "stroke_width_pct": 0.006},
                    # Títulos acima das nuvens (não mais dentro delas)
                    {"text": "VIVO", "xy_pct": (0.14, 0.15), "font_size_pct": 0.035, "stroke_width_pct": 0.008},
                    {"text": n,      "xy_pct": (0.87, 0.15), "font_size_pct": 0.035, "stroke_width_pct": 0.008},
                ]
                # Lado VIVO: SBC escolhido pela capacidade. PTIs da época do
                # SipRouter (sem SBC escolhido) mantêm CDSIP_SPO_PL / CDSIP_SPO_JG.
                if sbc_vivo:
                    row_labels.append({"text": sbc_vivo, "xy_pct": (0.14, 0.41), "font_size_pct": 0.022, "stroke_width_pct": 0.006})
                elif pl_block or jg_block:
                    row_labels.append({"text": "CDSIP_SPO_PL", "xy_pct": (0.14, 0.41), "font_size_pct": 0.022, "stroke_width_pct": 0.006})
                    row_labels.append({"text": "CDSIP_SPO_JG", "xy_pct": (0.14, 0.58), "font_size_pct": 0.022, "stroke_width_pct": 0.006})
                if pl_block:
                    row_labels.append({"text": pl_block, "xy_pct": (0.14, 0.45), "font_size_pct": 0.020, "stroke_width_pct": 0.004})
                if jg_block:
                    row_labels.append({"text": jg_block, "xy_pct": (0.14, 0.62), "font_size_pct": 0.020, "stroke_width_pct": 0.004})
                # IP da rede reservada IPAM
                if ip_op_base:
                    row_labels.append({"text": ip_op_base, "xy_pct": (0.87, 0.45), "font_size_pct": 0.020, "stroke_width_pct": 0.005})
                # SBC da operadora (preenchido no formulário Atacado → Dados Operadora)
                if sbc_op_img:
                    row_labels.append({"text": sbc_op_img, "xy_pct": (0.87, 0.41), "font_size_pct": 0.022, "stroke_width_pct": 0.006})
                if cn_val:
                    cn_img_text = f"CN {cn_val}"  # sem cidade
                    row_labels.append({"text": cn_img_text, "xy_pct": (0.50, 0.08), "font_size_pct": 0.028, "stroke_width_pct": 0.007})
                try:
                    ann = render_labels_on_image(
                        image_path=img_path, vivo_text="", operator_text="",
                        vivo_xy_pct=cfg["vxy"], operator_xy_pct=cfg["oxy"],
                        font_path=cfg["fp"], font_size_pct=cfg["fsz"],
                        stroke_width_pct=0.009, extra_labels=row_labels,
                    )
                    xi = XLImage(ann)
                    try:
                        with PILImage.open(ann) as src:
                            ow, oh = src.size
                    except Exception:
                        ow, oh = 800, 300
                    fw, fh = fit_size_keep_aspect(ow, oh, cfg["bw"], cfg["bh"])
                    xi.width, xi.height = fw, fh
                    anchor_col = get_column_letter(LC + cfg["aco"])
                    anchor_row = max(1, img_anchor_row + cfg["aro"])
                    ws.add_image(xi, f"{anchor_col}{anchor_row}")
                    cursor = img_anchor_row + max(22, int(fh / 15) + 2) + 2
                except Exception as exc:
                    logger.error("Erro ao gerar imagem (bloco %d): %s | path: %s",
                                 b_idx + 1, exc, img_path)
                    cursor += 3
            else:
                logger.error("IMAGEM NÃO ENCONTRADA: '%s' — verifique DIAGRAM_IMAGE_PATH "
                             "ou coloque o arquivo em static/diagram.png", img_path)
                cursor += 3
        self._setup_page(ws, ls=True)

    # =========================================================================
    # ABA: ENCAMINHAMENTO
    # =========================================================================
    def _build_routing_sheet(self) -> None:
        """Aba Encaminhamento.

        Layout: um único cabeçalho no topo; cada interconexão (bloco) é
        empilhada como um grupo de linhas logo abaixo, sem repetir cabeçalho.
        A coluna ATIVAÇÃO foi ocultada a pedido (código preservado em
        comentários para reativação futura).
        """
        s  = self.s
        ws = self.wb.create_sheet(title="Encaminhamento")
        n  = self.nome

        # Colunas: 16 = "Formato do Nº de A" (entre CAPS e ENCAMINHAMENTO,
        # como na planilha oficial). A antiga "ATIVAÇÃO" segue removida.
        col_widths = {
            1:2,2:16,3:10,4:16,5:10,6:18,7:16,8:18,9:14,
            10:8,11:8,12:8,13:8,14:8,15:8,
            16:22,                       # Formato do Nº de A
            17:18,18:12,19:12,20:16,
            21:14,22:12,23:14,24:12,25:2,
        }
        max_widths = dict(col_widths)

        def _tw(col, value):
            if value:
                cw = len(str(value)) * 1.15 + 2
                if cw > max_widths.get(col, 0):
                    max_widths[col] = min(cw, 45)

        self._write_brand_header(ws, 2, 2, 24)
        self._write_section_title(ws, 4, 2, 24,
                 "2.3. CARACTERÍSTICAS DO PROJETO DE INTERLIGAÇÃO E DO PLANO DE ENCAMINHAMENTO")

        # Um bloco por interconexão. O total é o maior lado (VIVO ou Operadora),
        # para nunca perder linhas quando os lados têm contagens diferentes.
        num_blocks     = max(1, len(self.vivo_rows), len(self.op_rows))
        traffic_items  = [t for t in self.traffic if t]
        align_merged   = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # ------------------------------------------------------------------
        # CABEÇALHO ÚNICO (h1/h2/h3)
        # ------------------------------------------------------------------
        h1 = 6; ws.row_dimensions[h1].height = 30.0
        for sc, ec, t, vert in [
            (2,2,"ÁREA LOCAL",True),(3,6,"LOCALIZAÇÃO",False),(7,9,"DADOS DA ROTA",False),
            (10,11,"BANDA",False),(12,13,"CANAIS",False),(14,15,"CAPS",False),
            # (16,16,"ATIVAÇÃO",True),  # coluna ATIVAÇÃO ocultada
            (16,16,"FORMATO DO",False),
            (17,18,"ENCAMINHAMENTO",False),
            (19,20,"SINALIZAÇÃO",False),(21,24,"ENDEREÇO IP",False),
        ]:
            if vert:
                ws.merge_cells(start_row=h1, start_column=sc, end_row=h1+2, end_column=ec)
            else:
                ws.merge_cells(start_row=h1, start_column=sc, end_row=h1, end_column=ec)
            c = ws.cell(row=h1, column=sc, value=t)
            c.font = s.font_subheader; c.alignment = s.align_center_wrap
            c.fill = s.fill_secondary if vert else s.fill_light
            for col in range(sc, ec + 1):
                ws.cell(row=h1, column=col).border = s.box_border

        h2 = h1 + 1; ws.row_dimensions[h2].height = 25.0
        for col, t in [
            (3,"CN"),(4,"POI/PPI VIVO"),(5,"CN"),(6,f"POI/PPI {n.upper()}"),
            (7,"PONTA A VIVO"),(8,f"PONTA B {n.upper()}"),(9,"TIPO DE TRÁFEGO"),
            (10,"EXIST."),(11,"PLAN."),(12,"EXIST."),(13,"PLAN."),(14,"EXIST."),(15,"PLAN."),
            (16,"Nº DE A"),
            (17,"DE A > B\n(FORMATO DE ENTREGA)"),(18,"DE B > A"),(19,"CODEC"),(20,"OBSERVAÇÃO"),
        ]:
            ws.merge_cells(start_row=h2, start_column=col, end_row=h2+1, end_column=col)
            c = ws.cell(row=h2, column=col, value=t)
            c.font = Font(name="Calibri", size=7 if col==17 else 8, bold=True, color="FFFFFF")
            c.alignment = s.align_center_wrap; c.fill = s.fill_accent
            c.border = s.box_border
            ws.cell(row=h2+1, column=col).border = s.box_border
        for sc, ec, t in [(21,22,"VIVO"),(23,24,n.upper())]:
            ws.merge_cells(start_row=h2, start_column=sc, end_row=h2, end_column=ec)
            c = ws.cell(row=h2, column=sc, value=t)
            c.font = s.font_micro; c.alignment = s.align_center; c.fill = s.fill_accent
            for col in range(sc, ec + 1):
                ws.cell(row=h2, column=col).border = s.box_border

        h3 = h2 + 1; ws.row_dimensions[h3].height = 25.0
        for col, t in [(21,"IP ADDRESS"),(22,"NETMASK"),(23,"IP ADDRESS"),(24,"NETMASK")]:
            c = ws.cell(row=h3, column=col, value=t)
            c.font = Font(name="Calibri", size=7, bold=True, color="FFFFFF")
            c.alignment = s.align_center; c.fill = s.fill_light; c.border = s.box_border

        # Contorno externo grosso do bloco de cabeçalho (h1–h3)
        self._thick_outline(ws, h1, 2, h3, 24)

        # ------------------------------------------------------------------
        # BLOCOS DE INTERCONEXÃO — empilhados sob o cabeçalho único
        # ------------------------------------------------------------------
        cursor = h3 + 1

        for b_idx in range(num_blocks):
            # Pareia VIVO×Operadora pelo índice; se um lado não tem esta linha,
            # tenta casar pelo CN do outro (mantém o par correto e não descarta).
            vivo_row = self.vivo_rows[b_idx] if b_idx < len(self.vivo_rows) else {}
            op_row   = self.op_rows[b_idx]   if b_idx < len(self.op_rows)   else {}
            if not vivo_row and op_row:
                vivo_row = self._find_matching_vivo_row(op_row.get("cn", ""), b_idx) or {}
            if not op_row and vivo_row:
                op_row = self._find_matching_op_row(vivo_row.get("cn", ""), b_idx) or {}

            cn_vivo        = vivo_row.get("cn", "") or ""
            cn_op          = op_row.get("cn", "")   or cn_vivo
            localidade_vivo = vivo_row.get("localidade", "") or ""
            localidade_op  = op_row.get("localidade", "")   or ""
            bloco_ip_vivo  = vivo_row.get("bloco_ip", "") or ""
            pl_blk, jg_blk = self._extract_pl_jg(bloco_ip_vivo)
            # Ponta A VIVO: SBC escolhido pela capacidade (legado: SipRouter SP)
            # ponta_a_vivo = ... or "CDSIP_SPO_PL/CDSIP_SPO_JG"  (padrão antigo)
            ponta_a_vivo   = self._sbc_vivo(vivo_row)
            ponta_b_op     = op_row.get("sbc", "") or op_row.get("faixa_ip", "") or ""
            mask_vivo      = vivo_row.get("mask", "")   or ""
            faixa_ip_op    = op_row.get("faixa_ip", "") or ""
            cidade_vivo    = vivo_row.get("cidade", "") or ""
            cidade_op      = op_row.get("cidade", "")   or ""
            area_local     = cidade_vivo or cidade_op
            uf_val         = vivo_row.get("uf", "") or op_row.get("uf", "")

            ipam_rede_enc = self.ipam_reservas.get(cn_vivo, "") if cn_vivo else ""
            ip_op_base    = ipam_rede_enc or faixa_ip_op   # nunca usar endereco_link
            _, mask_cidr  = self._parse_cidr(ip_op_base)
            mask_op       = self._cidr_to_netmask(mask_cidr) if mask_cidr else mask_vivo
            _, pl_mask    = self._parse_cidr(pl_blk)
            _, jg_mask    = self._parse_cidr(jg_blk)
            netmask_vivo  = self._cidr_to_netmask(pl_mask) if pl_mask else mask_vivo

            ds              = cursor
            # ATIVAÇÃO ocultada — para reativar, reintroduzir a coluna 16:
            # ativacao_str = datetime.now().strftime("%d/%m/%Y") + " + 90 dias"
            num_traffic     = len(traffic_items) if traffic_items else 1
            DATA_ROWS       = max(num_traffic, 1)
            row_height      = max(18.0, round(110.0 / DATA_ROWS, 1))
            traffic_mappings = [
                self._resolve_traffic_mapping(ti, n, cn_vivo, cn_op)
                for ti in (traffic_items if traffic_items else [""])
            ]
            merged_cols = {
                2:(area_local,       Font(name="Calibri", size=10, bold=True)),
                3:(cn_vivo,          Font(name="Calibri", size=10, bold=True)),
                4:(localidade_vivo,  s.font_small),
                5:(cn_op,            Font(name="Calibri", size=10, bold=True)),
                6:(localidade_op,    s.font_small),
                7:(ponta_a_vivo,     s.font_small),
                8:(ponta_b_op,       s.font_small),
                # 16:(ativacao_str,  s.font_small),  # ATIVAÇÃO ocultada
            }
            # Netmasks (cols 22 e 24) mescladas para todos os tráfegos
            if DATA_ROWS > 1:
                for netmask_col, netmask_val in [(22, netmask_vivo), (24, mask_op)]:
                    ws.merge_cells(start_row=ds, start_column=netmask_col,
                                   end_row=ds+DATA_ROWS-1, end_column=netmask_col)
                    c = ws.cell(row=ds, column=netmask_col, value=netmask_val)
                    c.font = s.font_small; c.alignment = align_merged
                    c.fill = s.alt_fill(0); c.border = s.box_border
                merged_cols[22] = (netmask_vivo, s.font_small)
                merged_cols[24] = (mask_op,      s.font_small)
            for col_num, (val, fnt) in merged_cols.items():
                ws.merge_cells(start_row=ds, start_column=col_num,
                               end_row=ds+DATA_ROWS-1, end_column=col_num)
                c = ws.cell(row=ds, column=col_num, value=val)
                c.font = fnt; c.alignment = align_merged
                c.fill = s.alt_fill(0); c.border = s.box_border
                _tw(col_num, val)

            merged_set = set(merged_cols.keys())
            for i in range(DATA_ROWS):
                r = ds + i; ws.row_dimensions[r].height = row_height
                f = s.alt_fill(i)
                traf_val = traffic_items[i] if i < len(traffic_items) else ""
                traf_map = traffic_mappings[i] if i < len(traffic_mappings) else {}

                # IPs VIVO: PL_ip / JG_ip por tráfego (mesmo offset +4+i do diagrama)
                pl_ip_enc = self._ip_only(self._add_last_octet(pl_blk, 4 + i)) if pl_blk else ""
                jg_ip_enc = self._ip_only(self._add_last_octet(jg_blk, 4 + i)) if jg_blk else ""
                vivo_ip_str = f"{pl_ip_enc} / {jg_ip_enc}" if pl_ip_enc and jg_ip_enc else (pl_ip_enc or jg_ip_enc)

                # IP Operadora: IPAM base + 2+i (mesmo offset do diagrama)
                op_ip_enc = self._ip_only(self._add_last_octet(ip_op_base, 2 + i)) if ip_op_base else ""

                for col in range(2, 25):
                    if col in merged_set:
                        ws.cell(row=r, column=col).border = s.box_border
                        ws.cell(row=r, column=col).fill   = f
                        continue
                    if col == 21:
                        val = vivo_ip_str
                    elif col == 22:
                        val = netmask_vivo if DATA_ROWS == 1 else None
                    elif col == 23:
                        val = op_ip_enc
                    elif col == 24:
                        val = mask_op if DATA_ROWS == 1 else None
                    elif 10 <= col <= 15:
                        # BANDA / CANAIS / CAPS — Exist. e Plan. por tráfego
                        met = self.trafego_metricas.get(traf_val, {})
                        campo = {10:"banda_exist", 11:"banda_plan",
                                 12:"canais_exist", 13:"canais_plan",
                                 14:"caps_exist",  15:"caps_plan"}[col]
                        val = met.get(campo, "")
                    else:
                        val = (traf_val if col==9
                               else traf_map.get("formato_no_a","") if col==16
                               else traf_map.get("enc_ab","") if col==17
                               else traf_map.get("enc_ba","") if col==18
                               else traf_map.get("codec","")  if col==19
                               else traf_map.get("obs","SIP-I") if col==20 else "")
                    if val is None:
                        ws.cell(row=r, column=col).border = s.box_border
                        ws.cell(row=r, column=col).fill   = f
                        continue
                    c = ws.cell(row=r, column=col, value=val)
                    c.font = s.font_small; c.alignment = align_merged
                    c.border = s.box_border; c.fill = f
                    if i == 0: _tw(col, val)

            missing = []
            if not cn_vivo:       missing.append("CN VIVO")
            if not area_local:    missing.append("ÁREA LOCAL")
            if not ip_op_base:    missing.append("IP DO IPAM")
            if not uf_val:        missing.append("UF")

            if missing:
                note_row = ds + DATA_ROWS
                ws.merge_cells(start_row=note_row, start_column=2, end_row=note_row, end_column=24)
                c = ws.cell(row=note_row, column=2,
                            value=f"⚠ Campos não preenchidos: {', '.join(missing)}")
                c.font = Font(name="Calibri", size=8, italic=True, color="FF6600")
                c.alignment = s.align_left; ws.row_dimensions[note_row].height = 18.0
                cursor = note_row + 1
            else:
                cursor = ds + DATA_ROWS

        # Contorno externo grosso da tabela inteira (cabeçalho + dados)
        self._thick_outline(ws, h1, 2, cursor - 1, 24)

        for col_num, width in max_widths.items():
            ws.column_dimensions[get_column_letter(col_num)].width = max(
                col_widths.get(col_num, 8), width
            )
        for col in (1, 25):
            for r in range(1, cursor + 5):
                ws.cell(row=r, column=col).fill   = s.fill_white
                ws.cell(row=r, column=col).border = s.no_border
        ws.freeze_panes = "C9"
        self._setup_page(ws, ls=True)

    # =========================================================================
    # ABA: PLAN NUM_OPER
    # =========================================================================
    def _build_prefixes_sheet(self) -> None:
        s  = self.s
        ws = self.wb.create_sheet(title="Plan Num_Oper")
        self._set_column_widths(ws, {1:2,2:4,3:12,4:12,5:12,6:6,7:10,8:10,9:8,10:15})
        self._write_brand_header(ws, 2, 2, 10)
        self._write_section_title(ws, 4, 2, 10, f"2.6. Tabela de Prefixos da {self.nome.upper()}")
        ws.merge_cells(start_row=5, start_column=2, end_row=5, end_column=10)
        c = ws.cell(row=5, column=2,
                    value='OBS: Caso necessite incluir mais faixas, inserir linhas adicionais (coluna B - "Ref").')
        c.font = Font(name="Calibri", size=9, italic=True)
        c.alignment = s.align_wrap; c.fill = s.fill_background

        h1 = 7; ws.row_dimensions[h1].height = 25.0
        for sc, ec, t, vert in [
            (2,2,"Ref",True),(3,3,"Município",True),(4,4,"ÁREA LOCAL",True),
            (5,5,"CÓDIGO CNL",True),(6,6,"CN",True),(7,8,"SERIE AUTORIZADA",False),
            (9,9,"EOT LOCAL",True),(10,10,"SNOA STFC / SMP",True),
        ]:
            if vert:
                ws.merge_cells(start_row=h1, start_column=sc, end_row=h1+1, end_column=ec)
            else:
                ws.merge_cells(start_row=h1, start_column=sc, end_row=h1, end_column=ec)
            c = ws.cell(row=h1, column=sc, value=t)
            c.font = s.font_subheader; c.alignment = s.align_center
            c.fill = s.fill_secondary; c.border = s.box_border

        h2 = h1 + 1; ws.row_dimensions[h2].height = 30.0
        for col, t in [(7,"INCLUÍNOS\nN8 N7 N6 N5"),(8,"FINAL\nN4 N3 N2 N1")]:
            c = ws.cell(row=h2, column=col, value=t)
            c.font = s.font_micro; c.alignment = s.align_center_wrap
            c.fill = s.fill_accent; c.border = s.box_border

        ds = h2 + 1
        for i in range(5):
            r = ds + i; ws.row_dimensions[r].height = 20.0; f = s.alt_fill(i)
            for col in range(2, 11):
                c = ws.cell(row=r, column=col, value="")
                c.font = s.font_small; c.alignment = s.align_center; c.border = s.box_border; c.fill = f
            if i > 0:
                ws.row_dimensions[r].hidden = True
        # Cabeçalho (h1–h2) e tabela de prefixos inteira: contorno externo grosso
        self._thick_outline(ws, h1, 2, h2, 10)
        self._thick_outline(ws, h1, 2, ds + 4, 10)

        ps = ds + 6
        title_font  = Font(name="Calibri", size=9, bold=True)
        title_align = Alignment(horizontal="left", vertical="center")
        resp_align  = Alignment(horizontal="center", vertical="center", wrap_text=True)
        font_ok  = Font(name="Calibri", size=14, bold=True, color="008000")
        font_nok = Font(name="Calibri", size=14, bold=True, color="FF0000")
        font_tv  = Font(name="Calibri", size=10, bold=True)
        TITLE_H  = 22.0

        indicators = [
            ("RN1",             "text",   self.rn1),
            ("EOT Local",       "text",   self.eot_local),
            ("EOT LDN/LDI",    "text",   self.eot_ld),
            ("CSP",             "binary", self.csp),
            ("Código Especial:", "binary", self.servicos_especiais),
            ("CNG:",            "binary", self.cng),
            ("RN2",             "binary", False),
        ]
        for i, (label, tipo, valor) in enumerate(indicators):
            r = ps + i; ws.row_dimensions[r].height = TITLE_H; f = s.alt_fill(i)
            ct = ws.cell(row=r, column=2, value=label)
            ct.font = title_font; ct.alignment = title_align; ct.border = s.box_border; ct.fill = f
            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=10)
            cr = ws.cell(row=r, column=3)
            cr.alignment = resp_align; cr.border = s.box_border; cr.fill = f
            if tipo == "binary":
                cr.value = "✔" if valor else "✘"
                cr.font  = font_ok if valor else font_nok
            else:
                if valor and str(valor).strip():
                    cr.value = str(valor).strip(); cr.font = font_tv
                else:
                    cr.value = "✘"; cr.font = font_nok
        # Tabela de indicadores (SE REG III): contorno externo grosso
        self._thick_outline(ws, ps, 2, ps + len(indicators) - 1, 10)

        desc_start = ps + len(indicators)
        for o, t in [
            (0, "Operadora de transporte nas localidades onde não existir rota direta"),
            (1, "Transbordo nas localidades onde existe rota direta:"),
        ]:
            ws.merge_cells(start_row=desc_start+o, start_column=2, end_row=desc_start+o, end_column=10)
            ws.cell(row=desc_start+o, column=2, value=t).font = Font(name="Calibri", size=9, bold=True)
            ws.cell(row=desc_start+o, column=2).alignment = s.align_wrap
            ws.cell(row=desc_start+o, column=2).fill = s.fill_background

        tr = desc_start + 2
        for sc, ec, t in [(2,3,"UF"),(4,7,"SERIE AUTORIZADA")]:
            ws.merge_cells(start_row=tr, start_column=sc, end_row=tr, end_column=ec)
            c = ws.cell(row=tr, column=sc, value=t)
            c.font = s.font_subheader; c.alignment = s.align_center
            c.fill = s.fill_secondary; c.border = s.box_border
        sr = tr + 1
        for sc, ec in [(2,3),(4,5),(6,7)]:
            ws.merge_cells(start_row=sr, start_column=sc, end_row=sr, end_column=ec)
            ws.cell(row=sr, column=sc, value="").border = s.box_border
            ws.cell(row=sr, column=sc).fill = s.fill_background
        # Tabela UF / Série Autorizada: contorno grosso (cabeçalho + linha)
        self._thick_outline(ws, tr, 2, tr, 7)
        self._thick_outline(ws, tr, 2, sr, 7)
        cr = sr + 2
        for o, t in [
            (0,"A programação de CNG ocorre pelo processo de portabilidade intrínseca - via ABR"),
            (1,"O encaminhamento do CNG será programado nas localidades onde o ponto de entrega for informado"),
        ]:
            ws.merge_cells(start_row=cr+o, start_column=2, end_row=cr+o, end_column=10)
            ws.cell(row=cr+o, column=2, value=t).font = s.font_small
        sm = cr + 3
        ws.merge_cells(start_row=sm, start_column=2, end_row=sm, end_column=10)
        ws.cell(row=sm, column=2, value="2.12 Tabela de Prefixos Vivo SMP").font = Font(name="Calibri", size=10, bold=True)
        ws.merge_cells(start_row=sm+1, start_column=2, end_row=sm+1, end_column=10)
        ws.cell(row=sm+1, column=2,
                value="Os prefixos da VIVO SMP, poderão ser obtidos acessando o site www.telefonica.net.br/sp/transfer/").font = s.font_small
        ws.freeze_panes = f"B{h1}"
        self._setup_page(ws, ls=True)

    # =========================================================================
    # ABA: DADOS MTL
    # =========================================================================
    def _build_mtl_sheet(self) -> None:
        s  = self.s
        ws = self.wb.create_sheet(title="Dados MTL")
        n  = self.nome
        # Coluna "ID" removida a pedido — as colunas seguintes deslocaram
        # uma posição à esquerda. Para reativar, reinsira "ID" na coluna 4
        # e volte as larguras/índices comentados abaixo.
        # col_w = {1:2,2:6,3:8,4:10,5:12,6:18,7:18,8:20,9:15,10:18,11:18,12:20}
        col_w = {1:2,2:6,3:8,4:12,5:18,6:18,7:20,8:15,9:18,10:18,11:20}
        self._set_column_widths(ws, col_w)
        self._write_brand_header(ws, 2, 2, 11)
        self._write_section_title(ws, 4, 2, 11, "2.7. Dados de MTL")

        for row, text, italic in [
            (5,"Tabela de dados MTL para interligação entre VIVO e operadora.", False),
            (6,"Preencher com os dados técnicos dos circuitos de interligação.", True),
        ]:
            ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=11)
            c = ws.cell(row=row, column=2, value=text)
            c.font = Font(name="Calibri", size=9 if italic else 10, italic=italic)
            c.fill = s.fill_background
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        hr = 8; ws.row_dimensions[hr].height = 25.0
        for col, txt in [
            (2,"Ref"),(3,"Cn"),
            # (4,"ID"),  # coluna ID removida a pedido
            (4,"MTL"),(5,"PONTA VIVO"),
            (6,f"PONTA {n.upper()}"),(7,"DESIGNADOR DO CIRCUITO"),(8,"PROVEDOR"),
            (9,"BGP VIVO"),(10,f"BGP {n.upper()}"),(11,"OBSERVAÇÕES"),
        ]:
            c = ws.cell(row=hr, column=col, value=txt)
            c.font = s.font_subheader
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.fill = s.fill_secondary; c.border = s.box_border

        ac = Alignment(horizontal="center", vertical="center", wrap_text=True)
        al = Alignment(horizontal="left",   vertical="center", wrap_text=True)
        num_refs = max(len(self.vivo_rows), len(self.op_rows), 1)
        ds = hr + 1; row_idx = 0

        for ref_i in range(num_refs):
            vr  = self.vivo_rows[ref_i] if ref_i < len(self.vivo_rows) else {}
            opr = self.op_rows[ref_i]   if ref_i < len(self.op_rows)   else {}
            cv      = vr.get("cn", "").strip()
            co      = opr.get("cn", "").strip()
            ev      = vr.get("endereco_link", "").strip()
            eo      = opr.get("endereco_link", "").strip()
            id_vivo = vr.get("id_vivo", "").strip()
            id_op   = opr.get("id_op", "").strip()

            # Sem a coluna ID: MTL passou de 5→4, PONTA VIVO 6→5, etc.
            # Versão anterior (com ID em 4):
            #   {3:cv, 4:id_vivo, 5:n.upper(), 6:ev, 7:eo, 8:"", ...}
            for row_data, id_val in [
                ({3:cv,       4:n.upper(), 5:ev, 6:eo, 7:"", 8:"", 9:"", 10:"", 11:""}, id_vivo),
                ({3:co or cv, 4:"VIVO",    5:ev, 6:eo, 7:"", 8:"", 9:"", 10:"", 11:""}, id_op),
            ]:
                r = ds + row_idx; ws.row_dimensions[r].height = 22.0; f = s.alt_fill(row_idx)
                for col_num, val in row_data.items():
                    c = ws.cell(row=r, column=col_num, value=val)
                    c.font = s.font_small
                    c.alignment = al if col_num in (5,6,11) else ac
                    c.border = s.box_border; c.fill = f
                row_idx += 1

            r1 = ds + row_idx - 2; r2 = ds + row_idx - 1
            ws.merge_cells(start_row=r1, start_column=2, end_row=r2, end_column=2)
            c = ws.cell(row=r1, column=2, value=ref_i+1)
            c.font = s.font_small; c.alignment = ac; c.border = s.box_border; c.fill = s.alt_fill(row_idx - 2)

        # Cabeçalho e tabela MTL inteira: contorno externo grosso
        self._thick_outline(ws, hr, 2, hr, 11)
        self._thick_outline(ws, hr, 2, ds + row_idx - 1, 11)

        ws.freeze_panes = f"B{ds}"
        self._setup_page(ws, ls=True)

    # =========================================================================
    # ABA: PARÂMETROS DE PROGRAMAÇÃO
    # =========================================================================
    def _build_params_sheet(self) -> None:
        """Aba 2.9 — Parâmetros de Programação.

        Gerada a partir de params_programacao.py. Quem responde é a
        OPERADORA: ela preenche a coluna com o próprio valor (D) e marca
        "Cumpre?" (F). O retorno é reimportado pela Engenharia.
        """
        s  = self.s
        ws = self.wb.create_sheet(title="Parâmetros de Programação")
        n  = self.nome
        # B: item | C: valor requerido | D: resposta da operadora
        # E: grau | F: Cumpre?         | G: observação
        self._set_column_widths(ws, {1:2, 2:46, 3:32, 4:26, 5:14, 6:12, 7:32, 8:2})
        self._write_brand_header(ws, 2, 2, 7)
        self._write_section_title(ws, 4, 2, 7, "2.9 - Parâmetros de Programação")

        ws.merge_cells(start_row=5, start_column=2, end_row=5, end_column=7)
        c = ws.cell(row=5, column=2, value=(
            f"Relação dos parâmetros de configuração de rotas SIP entre a "
            f"TELEFÔNICA-VIVO e a Operadora {n.upper()}. "
            f"Preencher a coluna \"{n.upper()}\" com os valores praticados e "
            f"marcar \"Cumpre?\" com ✔ ou ❌."
        ))
        c.font = s.font_body; c.fill = s.fill_background
        c.alignment = s.align_wrap
        ws.row_dimensions[5].height = 30.0

        cur = 7
        resposta_cells: list[str] = []      # células "Cumpre?" com lista ✔/❌

        for bloco in PARAMS_SECOES:
            # ── Título da seção ──
            ws.merge_cells(start_row=cur, start_column=2, end_row=cur, end_column=7)
            c = ws.cell(row=cur, column=2, value=bloco["secao"])
            c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            c.alignment = s.align_center; c.fill = s.fill_light
            c.border = s.box_border
            self._thick_outline(ws, cur, 2, cur, 7)
            ws.row_dimensions[cur].height = 20.0
            cur += 1

            # ── Cabeçalho das colunas ──
            cabecalhos = [
                bloco["col_item"], bloco["col_valor"], n.upper(),
                "Grau", "Cumpre?", "Observação",
            ]
            for ci, t in enumerate(cabecalhos, 2):
                c = ws.cell(row=cur, column=ci, value=t)
                c.font = s.font_subheader; c.alignment = s.align_center_wrap
                c.fill = s.fill_secondary; c.border = s.box_border
            self._thick_outline(ws, cur, 2, cur, 7)
            ws.row_dimensions[cur].height = 25.0
            cur += 1

            # ── Linhas de parâmetros ──
            inicio_dados = cur
            for i, (item, valor, grau, obs, responder) in enumerate(bloco["itens"]):
                f = s.alt_fill(i)
                valores = {
                    2: item,
                    3: valor,
                    4: "",                                   # operadora preenche
                    5: grau,
                    6: NAO_RESPONDIDO if responder else "",  # operadora responde
                    7: obs,
                }
                for ci, v in valores.items():
                    c = ws.cell(row=cur, column=ci, value=v)
                    c.font = s.font_small
                    c.alignment = s.align_wrap if ci in (2, 7) else s.align_center
                    c.fill = f; c.border = s.box_border
                if responder:
                    resposta_cells.append(f"F{cur}")
                ws.row_dimensions[cur].height = 20.0
                cur += 1

            self._thick_outline(ws, inicio_dados - 2, 2, cur - 1, 7)
            cur += 1

        # A operadora marca ✔ ou ❌ por lista suspensa — evita texto livre e
        # torna a reimportação confiável.
        if resposta_cells:
            dv = DataValidation(type="list", formula1='"✔,❌"', allow_blank=True)
            dv.promptTitle = "Cumpre?"
            dv.prompt = "Selecione ✔ (cumpre) ou ❌ (não cumpre)"
            ws.add_data_validation(dv)
            for ref in resposta_cells:
                dv.add(ws[ref])

        for r in range(1, cur + 3):
            for col in (1, 8):
                ws.cell(row=r, column=col).fill = s.fill_white
                ws.cell(row=r, column=col).border = s.no_border

        ws.freeze_panes = "B7"
        self._setup_page(ws, ls=True)

    # =========================================================================
    # BUILD
    # =========================================================================
    def build(self) -> Workbook:
        logger.info("Gerando PTI Excel para '%s'...", self.nome)
        self.reserva_redes_ipam()
        self._build_index_sheet()
        self._build_versions_sheet()
        self._build_diagram_sheet()
        self._build_routing_sheet()
        self._build_prefixes_sheet()
        self._build_mtl_sheet()
        self._build_params_sheet()

        self.wb.properties.title   = f"PTI — Completo — {self.nome}"
        self.wb.properties.subject = "Projeto Técnico de Interligação"
        self.wb.properties.creator = "VIVOHUB"
        self.wb.active = 0

        logger.info("PTI Excel gerado com sucesso.")
        return self.wb

