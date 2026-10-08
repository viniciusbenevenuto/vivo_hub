"""excel/styles.py — Paleta centralizada de estilos para planilhas Excel."""
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


class ExcelStylePalette:
    """Instancia e agrupa todos os estilos usados nas abas do PTI."""

    def __init__(self) -> None:
        # Preenchimentos
        self.fill_primary    = PatternFill("solid", fgColor="4A148C")
        self.fill_secondary  = PatternFill("solid", fgColor="6A1B9A")
        self.fill_light      = PatternFill("solid", fgColor="7B1FA2")
        self.fill_accent     = PatternFill("solid", fgColor="8E24AA")
        self.fill_neutral    = PatternFill("solid", fgColor="E1BEE7")
        self.fill_background = PatternFill("solid", fgColor="F3E5F5")
        self.fill_white      = PatternFill("solid", fgColor="FFFFFF")

        # Fontes
        self.font_title     = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
        self.font_subtitle  = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
        self.font_brand     = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
        self.font_header    = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        self.font_subheader = Font(name="Calibri", size=9,  bold=True, color="FFFFFF")
        self.font_body      = Font(name="Calibri", size=10)
        self.font_small     = Font(name="Calibri", size=9)
        self.font_micro     = Font(name="Calibri", size=8,  bold=True, color="FFFFFF")
        self.font_check     = Font(name="Calibri", size=11, bold=True, color="008000")
        self.font_x         = Font(name="Calibri", size=11, bold=True, color="FF0000")

        # Bordas — padrão do PTI modelo: dados com linha fina preta,
        # cabeçalhos com linha média preta
        thin   = Side(style="thin",   color="000000")
        medium = Side(style="medium", color="000000")
        self.box_border    = Border(left=thin, right=thin, top=thin, bottom=thin)
        self.header_border = Border(left=medium, right=medium, top=medium, bottom=medium)
        self.no_border     = Border()

        # Alinhamentos
        self.align_center      = Alignment(horizontal="center", vertical="center")
        self.align_left        = Alignment(horizontal="left",   vertical="center")
        self.align_wrap        = Alignment(horizontal="left",   vertical="center", wrap_text=True)
        self.align_center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def alt_fill(self, idx: int):
        """Alternância zebrada de fundo."""
        return self.fill_background if idx % 2 == 0 else self.fill_white
