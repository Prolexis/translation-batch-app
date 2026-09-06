"""
agents/extractor_agent.py
---------------------------
AGENTE EXTRACTOR
-----------------
Entrada : nombre de archivo + bytes crudos.
Salida  : lista de objetos `Segment` (párrafos limpios, sin traducir).
Rol     : aísla toda la complejidad de parseo de formato (.txt/.docx/.pdf) del
          resto del pipeline, de modo que los agentes siguientes (traductor,
          validador, alineador) trabajen siempre sobre la misma estructura de
          datos homogénea, independientemente del formato de origen.
"""

import logging
from typing import List

from utils.file_handlers import extract_text_segments, Segment, UnsupportedFormatError, CorruptFileError

logger = logging.getLogger("translation_app.extractor")


class ExtractorAgent:
    name = "extractor"

    def run(self, context: dict) -> dict:
        filename = context["filename"]
        file_bytes = context["file_bytes"]

        logger.info("[Extractor] Procesando %s", filename)
        try:
            raw_paragraphs = extract_text_segments(filename, file_bytes)
        except (UnsupportedFormatError, CorruptFileError) as exc:
            logger.error("[Extractor] Error en %s: %s", filename, exc)
            context["error"] = str(exc)
            context["segments"] = []
            return context

        segments: List[Segment] = [
            Segment(id=i, original=paragraph)
            for i, paragraph in enumerate(raw_paragraphs)
        ]

        context["segments"] = segments
        context.setdefault("memory_log", []).append(
            f"[extractor] {len(segments)} segmentos extraídos de '{filename}'."
        )
        logger.info("[Extractor] %d segmentos extraídos de %s", len(segments), filename)
        return context
