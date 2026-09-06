"""
agents/validator_agent.py
----------------------------
AGENTE VALIDADOR
------------------
Entrada : `Segment` ya traducido (.original + .translated).
Salida  : `Segment.status` en {"ok", "sospechoso", "error"} + notas de
          validación, y disparo de reintentos hacia el Agente Traductor
          (hasta `settings.MAX_RETRIES` veces por segmento).
"""

import logging
import re
from typing import List, TYPE_CHECKING

from config import settings
from utils.file_handlers import Segment

if TYPE_CHECKING:
    from agents.translator_agent import TranslatorAgent

logger = logging.getLogger("translation_app.validator")

_STOPWORD_WHITELIST = {
    "ok", "internet", "email", "online", "web", "app", "gps", "usb", "pdf",
}

_WORD_RE = re.compile(r"[A-Za-zÀ-ÿ]{%d,}" % settings.MIN_UNTRANSLATED_WORD_LEN)


class ValidatorAgent:
    name = "validator"

    def _length_ratio_suspicious(self, original: str, translated: str) -> bool:
        len_o, len_t = len(original.strip()), len(translated.strip())
        if len_o == 0:
            return False
        diff_ratio = abs(len_t - len_o) / len_o
        return diff_ratio > settings.LENGTH_DIFF_THRESHOLD

    def _find_untranslated_words(self, original: str, translated: str) -> List[str]:
        orig_words = {w.lower() for w in _WORD_RE.findall(original)}
        trans_words = {w.lower() for w in _WORD_RE.findall(translated)}
        overlap = orig_words & trans_words
        return sorted(w for w in overlap if w not in _STOPWORD_WHITELIST)

    def validate_segment(self, seg: Segment) -> bool:
        """Devuelve True si el segmento pasa la validación (status='ok')."""
        seg.validation_notes = []
        problems = []

        if self._length_ratio_suspicious(seg.original, seg.translated):
            len_o, len_t = len(seg.original.strip()), len(seg.translated.strip())
            ratio = abs(len_t - len_o) / max(len_o, 1)
            problems.append(
                f"Diferencia de longitud del {ratio:.0%} supera el umbral "
                f"({settings.LENGTH_DIFF_THRESHOLD:.0%})."
            )

        untranslated = self._find_untranslated_words(seg.original, seg.translated)
        if untranslated:
            problems.append(
                "Posibles palabras sin traducir: " + ", ".join(untranslated[:8])
            )

        if problems:
            seg.status = "sospechoso"
            seg.validation_notes = problems
            return False

        seg.status = "ok"
        return True

    def run(self, context: dict, translator: "TranslatorAgent" = None) -> dict:
        segments: List[Segment] = context.get("segments", [])
        source_lang = context.get("source_lang", settings.DEFAULT_SOURCE_LANG)
        target_lang = context.get("target_lang", settings.DEFAULT_TARGET_LANG)
        translator = translator or context.get("translator_agent")

        error_count = 0
        for seg in segments:
            if seg.status == "error":
                error_count += 1
                continue

            passed = self.validate_segment(seg)
            while not passed and seg.retries < settings.MAX_RETRIES and translator is not None:
                seg.retries += 1
                reason = "; ".join(seg.validation_notes)
                logger.info(
                    "[Validator] Segmento %d sospechoso (intento %d/%d): %s",
                    seg.id, seg.retries, settings.MAX_RETRIES, reason,
                )
                translator.retranslate_segment(seg, source_lang, target_lang, reason)
                if seg.status == "error":
                    break
                passed = self.validate_segment(seg)

            if seg.status != "ok":
                error_count += 1

        total = len(segments) or 1
        error_rate = error_count / total
        context["error_rate"] = error_rate
        context.setdefault("memory_log", []).append(
            f"[validator] {total - error_count}/{total} segmentos OK "
            f"(tasa de error {error_rate:.1%})."
        )
        logger.info("[Validator] Tasa de error del archivo: %.1f%%", error_rate * 100)
        return context
