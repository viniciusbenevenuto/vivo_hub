"""excel/image_utils.py — Renderização de labels em imagens usando Pillow."""
import logging
import os
import tempfile
from typing import Optional

logger = logging.getLogger(__name__)

try:
    from PIL import Image as PILImage, ImageDraw, ImageFont
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


def parse_xy_percent(raw, default: tuple) -> tuple:
    """Converte string 'x,y' em tupla de floats entre 0 e 1."""
    if not raw:
        return default
    try:
        parts = [p.strip() for p in str(raw).split(",")]
        if len(parts) != 2:
            return default
        return (
            max(0.0, min(1.0, float(parts[0]))),
            max(0.0, min(1.0, float(parts[1]))),
        )
    except (ValueError, TypeError):
        return default


def fit_size_keep_aspect(ow: int, oh: int, mw: int, mh: int) -> tuple[int, int]:
    """Redimensiona mantendo proporção dentro de mw×mh."""
    if ow <= 0 or oh <= 0:
        return mw, mh
    s = min(mw / ow, mh / oh)
    return int(ow * s), int(oh * s)


def _find_system_font(size_px: int, preferred: Optional[str] = None):
    if not PIL_AVAILABLE:
        return None
    if preferred:
        try:
            return ImageFont.truetype(preferred, size_px)
        except (OSError, IOError):
            pass
    for path in [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans.ttf",
        "/app/static/fonts/DejaVuSans.ttf",
        r"C:\Windows\Fonts\arial.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    ]:
        try:
            return ImageFont.truetype(path, size_px)
        except (OSError, IOError):
            continue
    return ImageFont.load_default()


def render_labels_on_image(
    image_path: str,
    vivo_text: str,
    operator_text: str,
    *,
    vivo_xy_pct: tuple = (0.18, 0.55),
    operator_xy_pct: tuple = (0.80, 0.55),
    font_path: Optional[str] = None,
    font_size_pct: float = 0.06,
    fill: tuple = (255, 255, 255, 255),
    stroke_fill: tuple = (0, 0, 0, 255),
    stroke_width_pct: float = 0.012,
    extra_labels: Optional[list] = None,
) -> str:
    """
    Renderiza labels sobre uma imagem existente.
    Retorna o caminho de um arquivo PNG temporário.
    """
    if not PIL_AVAILABLE:
        raise RuntimeError("Pillow não instalado.")

    img = PILImage.open(image_path).convert("RGBA")
    W, H = img.size
    draw = ImageDraw.Draw(img)

    def _draw(
        text: str, xy: tuple, sz: float, *,
        lf=fill, ls=stroke_fill, lw: float = stroke_width_pct, anch: str = "mm",
    ) -> None:
        if not text:
            return
        x, y = int(W * xy[0]), int(H * xy[1])
        font = _find_system_font(max(12, int(H * sz)), font_path)
        sw = max(0, int(H * lw))
        try:
            draw.text(
                (x, y), text, font=font, fill=lf,
                stroke_width=sw, stroke_fill=ls, anchor=anch,
            )
        except TypeError:
            # Fonte bitmap (load_default) não suporta anchor/stroke — usar fallback
            try:
                bbox = draw.textbbox((0, 0), text, font=font)
                tw = bbox[2] - bbox[0]
                th = bbox[3] - bbox[1]
                draw.text((x - tw // 2, y - th // 2), text, font=font, fill=lf)
            except Exception:
                draw.text((x, y), text, font=font, fill=lf)

    _draw(vivo_text,     vivo_xy_pct,     font_size_pct)
    _draw(operator_text, operator_xy_pct, font_size_pct)

    if extra_labels:
        for lbl in extra_labels:
            _draw(
                lbl.get("text", ""),
                lbl.get("xy_pct", (0.5, 0.5)),
                float(lbl.get("font_size_pct", font_size_pct)),
                lf=lbl.get("fill", fill),
                ls=lbl.get("stroke_fill", stroke_fill),
                lw=float(lbl.get("stroke_width_pct", stroke_width_pct)),
                anch=lbl.get("anchor", "mm"),
            )

    fd, tmp = tempfile.mkstemp(prefix="vivohub_diagrama_", suffix=".png")
    os.close(fd)
    img.save(tmp, format="PNG")
    return tmp
