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

_TECH_WHITELIST = {
    "ok", "internet", "email", "online", "web", "app", "gps", "usb", "pdf",
    "arduino", "ide", "led", "iot", "nano", "protoboard", "bootloader", "driver",
    "software", "hardware", "pin", "monitor", "python", "linux", "windows",
    "http", "https", "url", "bluetooth", "wifi", "microcontroller", "chip",
}


class ValidatorAgent:
    name = "validator"

    def _length_ratio_suspicious(self, original: str, translated: str) -> bool:
        len_o, len_t = len(original.strip()), len(translated.strip())
        if len_o < 30:  # títulos o líneas cortas varían naturalmente
            return False
        diff_ratio = abs(len_t - len_o) / len_o
        return diff_ratio > settings.LENGTH_DIFF_THRESHOLD

    def _find_untranslated_words(self, original: str, translated: str, source_lang: str, target_lang: str) -> List[str]:
        if source_lang.lower().strip() == target_lang.lower().strip():
            return []

        orig_clean = re.sub(r"https?://\S+|[^\w\s]", " ", original)
        trans_clean = re.sub(r"https?://\S+|[^\w\s]", " ", translated)

        orig_words = [w for w in orig_clean.split() if len(w) >= settings.MIN_UNTRANSLATED_WORD_LEN]
        if not orig_words:
            return []

        trans_words_set = {w.lower() for w in trans_clean.split()}
        
        suspicious = []
        for w in orig_words:
            wl = w.lower()
            if wl in _TECH_WHITELIST:
                continue
            if w[0].isupper():
                continue
            if wl in trans_words_set:
                suspicious.append(wl)

        unique_suspicious = sorted(set(suspicious))
        if len(unique_suspicious) >= 3 and (len(suspicious) / len(orig_words)) > 0.40:
            return unique_suspicious
        return []

    def validate_segment(self, seg: Segment, source_lang: str = "auto", target_lang: str = "es") -> bool:
        seg.validation_notes = []
        problems = []

        if not seg.translated or not seg.translated.strip():
            problems.append("La traducción está vacía.")

        if self._length_ratio_suspicious(seg.original, seg.translated):
            len_o, len_t = len(seg.original.strip()), len(seg.translated.strip())
            ratio = abs(len_t - len_o) / max(len_o, 1)
            problems.append(
                f"Diferencia de longitud ({ratio:.0%}) supera el umbral configurado."
            )

        untranslated = self._find_untranslated_words(seg.original, seg.translated, source_lang, target_lang)
        if untranslated:
            problems.append(
                "Posibles palabras sin traducir: " + ", ".join(untranslated[:6])
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

            passed = self.validate_segment(seg, source_lang, target_lang)
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
                passed = self.validate_segment(seg, source_lang, target_lang)

            if seg.status != "ok":
                error_count += 1

        total = len(segments) or 1
        error_rate = error_count / total
        context["error_rate"] = error_rate
        context.setdefault("memory_log", []).append(
            f"[validador] {total - error_count}/{total} segmentos OK "
            f"(tasa de error {error_rate:.1%})."
        )
        logger.info("[Validator] Tasa de error del archivo: %.1f%%", error_rate * 100)
        return context
