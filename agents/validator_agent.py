"""
agents/validator_agent.py
----------------------------
AGENTE VALIDADOR ACADÉMICO
----------------------------
Entrada : `Segment` ya traducido (.original + .translated).
Salida  : `Segment.status` en {"ok", "sospechoso", "error"} + notas de
          validación y verificación de integridad de citas académicas [1], (Smith et al., 2020),
          así como disparo de reintentos hacia el Agente Traductor.
"""

import logging
import re
from typing import List, TYPE_CHECKING

from config import settings
from utils.file_handlers import Segment

if TYPE_CHECKING:
    from agents.translator_agent import TranslatorAgent

logger = logging.getLogger("translation_app.validator")

# Palabras técnicas y académicas comunes en inglés aceptables en traducciones al español
_ACADEMIC_WHITELIST = {
    "ok", "internet", "email", "online", "web", "app", "gps", "usb", "pdf",
    "attention", "transformer", "bert", "gpt", "roberta", "t5", "lstm", "rnn", "cnn",
    "dataset", "datasets", "benchmark", "benchmarks", "score", "scores",
    "bleu", "rouge", "f1", "accuracy", "loss", "softmax", "embedding", "embeddings",
    "token", "tokens", "tokenizer", "tokenizers", "fine-tuning", "prompt", "prompts",
    "zero-shot", "few-shot", "self-attention", "multi-head", "layer", "layers",
    "dropout", "feed-forward", "encoder", "decoder", "cross-attention",
    "et", "al", "doi", "arxiv", "ieee", "acm", "acl", "emnlp", "neurips", "iclr",
}

# Regex para detectar citas entre corchetes [1], [1, 2], [3-5]
_BRACKET_CITATION_REGEX = re.compile(r"\[\s*\d+(?:[\s,\-–—]+\d+)*\s*\]")

# Regex para detectar citas autor-año (Vaswani et al., 2017) o (Devlin, 2018)
_AUTHOR_YEAR_REGEX = re.compile(r"\(([A-Z][a-zA-Z\s]+(?:et\s+al\.?)?,?\s*(?:19|20)\d{2}[a-z]?)\)")


class ValidatorAgent:
    name = "validator"

    def _check_citation_integrity(self, original: str, translated: str) -> List[str]:
        """Verifica que las citas numéricas y de autores del original sigan existiendo en la traducción."""
        issues = []

        # 1. Citas entre corchetes [1], [2], etc.
        orig_brackets = set(_BRACKET_CITATION_REGEX.findall(original))
        if orig_brackets:
            trans_brackets = set(_BRACKET_CITATION_REGEX.findall(translated))
            missing = orig_brackets - trans_brackets
            if missing:
                issues.append(f"Citas numéricas omitidas o alteradas: {', '.join(sorted(missing)[:3])}")

        # 2. Citas autor-año
        orig_authors = _AUTHOR_YEAR_REGEX.findall(original)
        for auth_cite in orig_authors:
            # Buscar si el año o el autor está en la traducción
            year_match = re.search(r"(?:19|20)\d{2}", auth_cite)
            if year_match:
                year = year_match.group(0)
                if year not in translated:
                    issues.append(f"Cita bibliográfica omitida ({auth_cite})")

        return issues

    def _length_ratio_suspicious(self, original: str, translated: str) -> bool:
        len_o, len_t = len(original.strip()), len(translated.strip())
        if len_o < 40:  # Títulos o fórmulas cortas varían naturalmente
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
            if wl in _ACADEMIC_WHITELIST:
                continue
            if w[0].isupper():  # Nombres propios o siglas
                continue
            if wl in trans_words_set:
                suspicious.append(wl)

        unique_suspicious = sorted(set(suspicious))
        if len(unique_suspicious) >= 4 and (len(suspicious) / len(orig_words)) > 0.45:
            return unique_suspicious
        return []

    def validate_segment(self, seg: Segment, source_lang: str = "auto", target_lang: str = "es") -> bool:
        # Segmentos de referencia o fórmulas se asumen válidos directamente
        if seg.element_type in ["reference", "formula"]:
            seg.status = "ok"
            seg.validation_notes = []
            return True

        seg.validation_notes = []
        problems = []

        if not seg.translated or not seg.translated.strip():
            problems.append("La traducción académica está vacía.")

        # Verificar citas académicas
        citation_issues = self._check_citation_integrity(seg.original, seg.translated)
        if citation_issues:
            problems.extend(citation_issues)

        # Verificar proporción de longitud
        if self._length_ratio_suspicious(seg.original, seg.translated):
            len_o, len_t = len(seg.original.strip()), len(seg.translated.strip())
            ratio = abs(len_t - len_o) / max(len_o, 1)
            problems.append(
                f"Diferencia de longitud ({ratio:.0%}) excede el umbral configurado."
            )

        # Palabras en inglés sin traducir
        untranslated = self._find_untranslated_words(seg.original, seg.translated, source_lang, target_lang)
        if untranslated:
            problems.append(
                "Términos no traducidos detectados: " + ", ".join(untranslated[:5])
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
        repaired_citations = 0

        for seg in segments:
            if seg.status == "error":
                error_count += 1
                continue

            passed = self.validate_segment(seg, source_lang, target_lang)
            # Reintentar únicamente si es un fallo crítico (traducción vacía o cita perdida)
            is_critical = any("vacía" in n.lower() or "citas numéricas" in n.lower() for n in seg.validation_notes)
            while not passed and is_critical and seg.retries < settings.MAX_RETRIES and translator is not None:
                seg.retries += 1
                reason = "; ".join(seg.validation_notes)
                logger.info(
                    "[Validator] Segmento %d requiere reintento crítico (%d/%d): %s",
                    seg.id, seg.retries, settings.MAX_RETRIES, reason,
                )
                translator.retranslate_segment(seg, source_lang, target_lang, reason)
                if seg.status == "error":
                    break
                passed = self.validate_segment(seg, source_lang, target_lang)
                if passed:
                    repaired_citations += 1

            if seg.status != "ok":
                error_count += 1

        total = len(segments) or 1
        error_rate = error_count / total
        context["error_rate"] = error_rate
        context.setdefault("memory_log", []).append(
            f"[validador] {total - error_count}/{total} párrafos aprobados con éxito "
            f"(tasa de advertencia {error_rate:.1%}, {repaired_citations} reintentos autocorregidos)."
        )
        logger.info("[Validator] Validación finalizada: %.1f%% advertencias", error_rate * 100)
        return context
