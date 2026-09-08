"""
agents/extractor_agent.py
---------------------------
AGENTE EXTRACTOR ACADÉMICO
---------------------------
Entrada : nombre de archivo + bytes crudos.
Salida  : lista de objetos `Segment` estructurados con metadatos de procedencia
          (sección, subsección, página, número de párrafo, tipo de elemento).
Rol     : aísla toda la complejidad de parseo de formato (.txt/.docx/.pdf) y
          segmentación académica del resto del pipeline, garantizando la trazabilidad
          precisa hacia el documento original.
"""

import logging
from typing import List

from utils.file_handlers import (
    extract_academic_segments,
    Segment,
    UnsupportedFormatError,
    CorruptFileError,
)

logger = logging.getLogger("translation_app.extractor")


class ExtractorAgent:
    name = "extractor"

    def __init__(self, *args, **kwargs):
        pass

    def run(self, context: dict) -> dict:
        filename = context["filename"]
        file_bytes = context["file_bytes"]

        logger.info("[Extractor] Extrayendo estructura académica de %s", filename)
        try:
            segments: List[Segment] = extract_academic_segments(filename, file_bytes)
        except (UnsupportedFormatError, CorruptFileError) as exc:
            logger.error("[Extractor] Error en %s: %s", filename, exc)
            context["error"] = str(exc)
            context["segments"] = []
            return context

        context["segments"] = segments
        context.setdefault("memory_log", []).append(
            f"[extractor] {len(segments)} párrafos/segmentos académicos extraídos de '{filename}' "
            f"con metadatos de página, sección y numeración."
        )
        logger.info("[Extractor] %d segmentos extraídos con éxito de %s", len(segments), filename)
        return context
