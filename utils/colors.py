"""
utils/colors.py
-----------------
Paleta cíclica de colores usada por la vista de "verificación cruzada":
el segmento N del texto original y el segmento N del texto traducido
reciben exactamente el mismo color (asignado por el Agente de Alineación),
lo que permite al usuario verificar visualmente la correspondencia.
"""

import html
from typing import List

from config import settings


def color_for_index(index: int) -> str:
    palette = settings.HIGHLIGHT_COLORS
    return palette[index % len(palette)]


def _contrast_text_color(hex_color: str) -> str:
    """Elige negro o blanco según la luminosidad del color de fondo."""
    hex_color = hex_color.lstrip("#")
    r, g, b = (int(hex_color[i:i + 2], 16) for i in (0, 2, 4))
    luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    return "#000000" if luminance > 0.6 else "#FFFFFF"


def render_highlighted_html(segments_text: List[str], colors: List[str]) -> str:
    """
    Genera HTML con cada segmento envuelto en un <span> coloreado.
    `segments_text` y `colors` deben tener la misma longitud (mismo orden
    que usa el par original/traducido para que los IDs coincidan).
    """
    blocks = []
    for text, color in zip(segments_text, colors):
        safe_text = html.escape(text).replace("\n", "<br>")
        fg = _contrast_text_color(color)
        blocks.append(
            f'<div style="background-color:{color}; color:{fg}; '
            f'padding:8px 10px; border-radius:6px; margin-bottom:8px; '
            f'font-size:0.92rem; line-height:1.4;">{safe_text}</div>'
        )
    return "\n".join(blocks)
