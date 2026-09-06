"""
agents/translator_agent.py
-----------------------------
AGENTE TRADUCTOR
------------------
Entrada : lista de `Segment` (con .original poblado) + idioma origen/destino.
Salida  : los mismos `Segment`, con .translated poblado.
Rol     : envuelve el LLM Gemini (vía LangChain `ChatGoogleGenerativeAI`) en
          una cadena LCEL (prompt | llm | parser) y aplica rate-limiting +
          backoff exponencial en cada llamada.
"""

import logging
from typing import List

from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import StrOutputParser
from langchain_google_genai import ChatGoogleGenerativeAI

from config import settings
from utils.file_handlers import Segment
from utils.rate_limiter import with_backoff

logger = logging.getLogger("translation_app.translator")

_SYSTEM_PROMPT = (
    "Eres un traductor profesional. Traduce el texto del usuario al idioma '{target_lang}'. "
    "Si el idioma origen es '{source_lang}' o 'auto', detecta e interpreta el contenido del original. "
    "Devuelve ÚNICAMENTE la traducción al idioma '{target_lang}', sin notas explicativas, "
    "sin comillas adicionales, preservando la estructura, saltos de línea y el formato original."
)

_RETRY_SUFFIX = (
    "\n\nIMPORTANTE: un intento previo de traducción fue rechazado por un "
    "verificador automático por el siguiente motivo: {retry_reason}. "
    "Corrige ese problema específico en esta nueva traducción."
)


class TranslatorAgent:
    name = "translator"

    def __init__(self, model_name: str = None, api_key: str = None, temperature: float = 0.2):
        self.model_name = model_name or settings.GEMINI_MODEL
        self.api_key = api_key or settings.GEMINI_API_KEY
        self._llm = ChatGoogleGenerativeAI(
            model=self.model_name,
            google_api_key=self.api_key,
            temperature=temperature,
        )
        self._parser = StrOutputParser()

    def _build_chain(self, retry_reason: str = None):
        system = _SYSTEM_PROMPT
        if retry_reason:
            system = system + _RETRY_SUFFIX.format(retry_reason=retry_reason)
        prompt = ChatPromptTemplate.from_messages([
            ("system", system),
            ("human", "{text}"),
        ])
        return prompt | self._llm | self._parser

    @with_backoff()
    def _translate_one(self, text: str, source_lang: str, target_lang: str, retry_reason: str = None) -> str:
        chain = self._build_chain(retry_reason=retry_reason)
        result = chain.invoke({
            "text": text,
            "source_lang": source_lang,
            "target_lang": target_lang,
        })
        return str(result).strip()

    def run(self, context: dict, on_progress: Optional[Callable[[str, float], None]] = None) -> dict:
        segments: List[Segment] = context.get("segments", [])
        source_lang = context.get("source_lang", settings.DEFAULT_SOURCE_LANG)
        target_lang = context.get("target_lang", settings.DEFAULT_TARGET_LANG)
        total_segs = len(segments) or 1

        for i, seg in enumerate(segments):
            if on_progress:
                frac = 0.25 + (0.45 * (i / total_segs))
                on_progress(f"Traduciendo segmento {i + 1}/{total_segs}", frac)
            if seg.status == "ok":
                continue
            try:
                seg.translated = self._translate_one(seg.original, source_lang, target_lang)
                seg.status = "traducido"
            except Exception as exc:  # noqa: BLE001
                logger.error("[Translator] Fallo traduciendo segmento %d: %s", seg.id, exc)
                seg.status = "error"
                seg.validation_notes.append(f"Error de traducción: {exc}")

        context.setdefault("memory_log", []).append(
            f"[translator] {len(segments)} segmentos procesados "
            f"({source_lang} -> {target_lang})."
        )
        return context

    def retranslate_segment(self, seg: Segment, source_lang: str, target_lang: str, reason: str) -> None:
        """Reintenta un único segmento marcado como sospechoso por el validador."""
        try:
            seg.translated = self._translate_one(
                seg.original, source_lang, target_lang, retry_reason=reason
            )
        except Exception as exc:  # noqa: BLE001
            logger.error("[Translator] Fallo en reintento del segmento %d: %s", seg.id, exc)
            seg.status = "error"
            seg.validation_notes.append(f"Error en reintento: {exc}")
