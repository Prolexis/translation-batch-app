"""
agents/aligner_agent.py
--------------------------
AGENTE DE ALINEACIÓN
-----------------------
Entrada : `Segment`s ya traducidos y validados.
Salida  : a cada `Segment` se le asigna un `color` (de la paleta cíclica de
          `config.settings.HIGHLIGHT_COLORS`).
"""

import logging
from typing import List

import numpy as np

from config import settings
from utils.file_handlers import Segment
from utils.colors import color_for_index

logger = logging.getLogger("translation_app.aligner")


class AlignerAgent:
    name = "aligner"

    def __init__(self, api_key: str = None):
        self._embeddings = None
        self.api_key = api_key or settings.GEMINI_API_KEY

    def _get_embeddings_client(self):
        if self._embeddings is None:
            from langchain_google_genai import GoogleGenerativeAIEmbeddings
            self._embeddings = GoogleGenerativeAIEmbeddings(
                model=settings.GEMINI_EMBEDDING_MODEL,
                google_api_key=self.api_key,
            )
        return self._embeddings

    def _align_by_position(self, segments: List[Segment]) -> None:
        for seg in segments:
            seg.color = color_for_index(seg.id)

    def _align_by_semantics(self, segments: List[Segment]) -> None:
        client = self._get_embeddings_client()
        originals = [s.original for s in segments]
        translations = [s.translated or s.original for s in segments]

        try:
            emb_orig = np.array(client.embed_documents(originals))
            emb_trans = np.array(client.embed_documents(translations))
        except Exception as exc:  # noqa: BLE001
            logger.warning(
                "[Aligner] Falló el cálculo de embeddings (%s); "
                "se usa alineación por posición como respaldo.", exc,
            )
            self._align_by_position(segments)
            return

        def cosine(a, b):
            na, nb = np.linalg.norm(a), np.linalg.norm(b)
            if na == 0 or nb == 0:
                return 0.0
            return float(np.dot(a, b) / (na * nb))

        for i, seg in enumerate(segments):
            sim = cosine(emb_orig[i], emb_trans[i])
            seg.color = color_for_index(i)
            if sim < settings.SEMANTIC_SIMILARITY_THRESHOLD:
                seg.validation_notes.append(
                    f"Similitud semántica baja original/traducción ({sim:.2f}); "
                    "revisar alineación manualmente."
                )
                if seg.status == "ok":
                    seg.status = "sospechoso"

    def run(self, context: dict) -> dict:
        segments: List[Segment] = context.get("segments", [])
        mode = context.get("alignment_mode", settings.ALIGNMENT_MODE)

        if mode == "semantic":
            self._align_by_semantics(segments)
        else:
            self._align_by_position(segments)

        context.setdefault("memory_log", []).append(
            f"[aligner] {len(segments)} segmentos alineados (modo='{mode}')."
        )
        return context
