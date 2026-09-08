"""
agents/translator_agent.py
-----------------------------
AGENTE TRADUCTOR ACADÉMICO DE ALTO RENDIMIENTO
-------------------------------------------------
Entrada : lista de `Segment` con .original y metadatos académicos.
Salida  : los mismos `Segment`, con .translated poblado respetando convenciones académicas.
Rol     : envuelve el LLM Gemini (vía LangChain `ChatGoogleGenerativeAI`) con:
          - Agrupación por lotes (batching) de múltiples párrafos por llamada (reducción 8x de llamadas HTTP).
          - Concurrencia multihilo (ThreadPoolExecutor con 4-5 workers en paralelo).
          - Preservación estricta de citas in-text [1], (Author, Year) y fórmulas.
          - Passthrough instantáneo (0 llamadas API) para bibliografía y fórmulas.
          - Rate-limiting + backoff exponencial ante 429.
"""

import logging
import re
from typing import List, Optional, Callable
from concurrent.futures import ThreadPoolExecutor, as_completed

from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import StrOutputParser
from langchain_google_genai import ChatGoogleGenerativeAI

from config import settings
from utils.file_handlers import Segment
from utils.rate_limiter import with_backoff

logger = logging.getLogger("translation_app.translator")

LANGUAGE_NAMES = {
    "auto": "Detección automática (auto-detect original language)",
    "es": "Español formal y académico (Spanish)",
    "en": "Inglés académico (English)",
    "fr": "Francés (French)",
    "de": "Alemán (German)",
    "it": "Italiano (Italian)",
    "pt": "Portugués (Portuguese)",
    "zh": "Chino (Chinese)",
    "ja": "Japonés (Japanese)",
    "ru": "Ruso (Russian)",
    "ar": "Árabe (Arabic)",
}

# Tamaño de lote óptimo de párrafos por prompt (16 párrafos por llamada para 5x-8x aceleración sin pausas de rate-limit)
BATCH_SIZE = getattr(settings, "DEFAULT_BATCH_SIZE", 16)
# Número de trabajadores concurrentes
MAX_CONCURRENT_WORKERS = getattr(settings, "DEFAULT_MAX_WORKERS", 5)


class TranslatorAgent:
    name = "translator"

    def __init__(self, model_name: str = None, api_key: str = None, temperature: float = 0.15,
                 batch_size: int = None, max_workers: int = None,
                 *args, **kwargs):
        self.model_name = model_name or settings.GEMINI_MODEL
        self.api_key = api_key or settings.GEMINI_API_KEY
        self.batch_size = int(batch_size) if batch_size is not None else BATCH_SIZE
        self.max_workers = int(max_workers) if max_workers is not None else MAX_CONCURRENT_WORKERS
        self._llm = ChatGoogleGenerativeAI(
            model=self.model_name,
            google_api_key=self.api_key,
            temperature=temperature,
        )
        self._parser = StrOutputParser()

    def _build_batch_chain(self, source_name: str, target_name: str):
        system = (
            f"You are an elite academic translator specializing in scientific papers and research articles.\n"
            f"Translate each numbered academic paragraph provided below from {source_name} into formal scientific {target_name}.\n\n"
            "STRICT ACADEMIC RULES:\n"
            "1. TERMINOLOGY: Use formal academic vocabulary in Spanish (e.g., 'mecanismo de atención', 'aprendizaje profundo', 'conjunto de datos', 'red neuronal').\n"
            "2. CITATIONS: Absolutely PRESERVE all in-text citations exactly as written, including bracket citations [1], [2, 3] and author-year citations (Smith et al., 2020). NEVER alter or remove citation markers.\n"
            "3. CROSS-REFERENCES: Preserve cross-reference labels accurately ('Figure 1' -> 'Figura 1', 'Table 2' -> 'Tabla 2', 'Eq. (3)' -> 'Ec. (3)').\n"
            "4. DELIMITERS: CRITICAL REQUIREMENT: You MUST prepend each translated paragraph with its exact marker [P_{{id}}], followed by the translated text on a new line.\n"
            "Do not omit any [P_{{id}}] marker. Return ONLY the marked translated paragraphs without any extra conversational filler.\n\n"
            "Example format:\n"
            "[P_0]\n"
            "Texto traducido del primer párrafo...\n\n"
            "[P_1]\n"
            "Texto traducido del segundo párrafo..."
        )
        prompt = ChatPromptTemplate.from_messages([
            ("system", system),
            ("human", "{batch_text}"),
        ])
        return prompt | self._llm | self._parser

    def _build_single_chain(self, retry_reason: str = None):
        system = (
            "You are an expert academic translator specializing in scholarly papers.\n"
            "Translate the source text accurately into {target_name}.\n"
            "Preserve in-text citations [1], (Author, Year), equations, and formal terminology.\n"
            "Return ONLY the direct translation."
        )
        if retry_reason:
            system += f"\n\nCRITICAL FIX: A previous translation was rejected due to: {retry_reason}. Rectify this defect."

        prompt = ChatPromptTemplate.from_messages([
            ("system", system),
            ("human", "Translate this academic text from {source_name} into {target_name}:\n\n{text}"),
        ])
        return prompt | self._llm | self._parser

    @with_backoff()
    def _translate_batch_call(self, batch_text: str, source_name: str, target_name: str) -> str:
        chain = self._build_batch_chain(source_name, target_name)
        result = chain.invoke({"batch_text": batch_text})
        return str(result).strip()

    @with_backoff()
    def _translate_one(self, text: str, source_lang: str, target_lang: str, retry_reason: str = None) -> str:
        source_name = LANGUAGE_NAMES.get(source_lang.lower().strip(), source_lang)
        target_name = LANGUAGE_NAMES.get(target_lang.lower().strip(), target_lang)
        chain = self._build_single_chain(retry_reason=retry_reason)
        result = chain.invoke({
            "text": text,
            "source_name": source_name,
            "target_name": target_name,
        })
        return str(result).strip()

    def _process_single_batch(self, batch: List[Segment], source_name: str, target_name: str) -> List[Segment]:
        """Procesa un lote de 6-8 párrafos en una única llamada LLM."""
        batch_input = "\n\n".join(f"[P_{s.id}]\n{s.original}" for s in batch)
        try:
            raw_response = self._translate_batch_call(batch_input, source_name, target_name)
            # Parsear bloques delimitados por [P_{id}]
            blocks = re.split(r'\[P_(\d+)\]', raw_response)
            parsed_dict = {}
            for i in range(1, len(blocks), 2):
                seg_id = blocks[i].strip()
                seg_text = blocks[i + 1].strip()
                parsed_dict[seg_id] = seg_text

            for seg in batch:
                if str(seg.id) in parsed_dict and parsed_dict[str(seg.id)]:
                    seg.translated = parsed_dict[str(seg.id)]
                    seg.status = "traducido"
                else:
                    # Fallback individual para este segmento específico si el tag faltó
                    logger.warning("[Translator] Segmento %d no encontrado en respuesta de lote, traduciendo individualmente", seg.id)
                    seg.translated = self._translate_one(seg.original, source_name, target_name)
                    seg.status = "traducido"

        except Exception as exc:  # noqa: BLE001
            logger.error("[Translator] Fallo en lote de %d segmentos: %s. Aplicando fallback individual...", len(batch), exc)
            for seg in batch:
                try:
                    seg.translated = self._translate_one(seg.original, source_name, target_name)
                    seg.status = "traducido"
                except Exception as inner_exc:  # noqa: BLE001
                    seg.status = "error"
                    seg.validation_notes.append(f"Error de traducción: {inner_exc}")

        return batch

    def run(self, context: dict, on_progress: Optional[Callable[[str, float], None]] = None) -> dict:
        segments: List[Segment] = context.get("segments", [])
        source_lang = context.get("source_lang", settings.DEFAULT_SOURCE_LANG)
        target_lang = context.get("target_lang", settings.DEFAULT_TARGET_LANG)
        source_name = LANGUAGE_NAMES.get(source_lang.lower().strip(), source_lang)
        target_name = LANGUAGE_NAMES.get(target_lang.lower().strip(), target_lang)

        total_segs = len(segments) or 1
        skipped_refs = 0
        skipped_formulas = 0
        segs_to_translate: List[Segment] = []

        # 1. Passthrough instantáneo para referencias y fórmulas (0 llamadas API)
        for seg in segments:
            if seg.element_type == "reference":
                seg.translated = seg.original
                seg.status = "ok"
                skipped_refs += 1
            elif seg.element_type == "formula":
                seg.translated = seg.original
                seg.status = "ok"
                skipped_formulas += 1
            elif seg.status == "ok" and seg.translated:
                pass
            else:
                segs_to_translate.append(seg)

        num_to_translate = len(segs_to_translate)
        logger.info(
            "[Translator] Total segmentos: %d | A traducir: %d | Omitidos (refs/fórmulas): %d",
            total_segs, num_to_translate, skipped_refs + skipped_formulas,
        )

        if not segs_to_translate:
            context.setdefault("memory_log", []).append(
                f"[translator] Todos los {total_segs} segmentos eran referencias o ya estaban traducidos."
            )
            return context

        # 2. Agrupar en lotes de tamaño batch_size
        eff_batch_size = context.get("batch_size", self.batch_size)
        eff_max_workers = context.get("max_workers", self.max_workers)

        batches = [
            segs_to_translate[i:i + eff_batch_size]
            for i in range(0, num_to_translate, eff_batch_size)
        ]
        total_batches = len(batches)
        completed_batches = 0

        if on_progress:
            on_progress(f"Iniciando traducción ultra rápida de {num_to_translate} párrafos ({total_batches} lotes en paralelo)...", 0.22)

        # 3. Ejecución concurrente multihilo con ThreadPoolExecutor
        workers = min(eff_max_workers, total_batches)
        with ThreadPoolExecutor(max_workers=workers) as executor:
            future_to_batch = {
                executor.submit(self._process_single_batch, b, source_name, target_name): b
                for b in batches
            }
            for future in as_completed(future_to_batch):
                completed_batches += 1
                frac = 0.22 + (0.50 * (completed_batches / total_batches))
                if on_progress:
                    done_count = min(completed_batches * eff_batch_size, num_to_translate)
                    on_progress(
                        f"Traduciendo en paralelo: lote {completed_batches}/{total_batches} ({done_count}/{num_to_translate} párrafos)",
                        frac,
                    )
                try:
                    future.result()
                except Exception as exc:  # noqa: BLE001
                    logger.error("[Translator] Error inesperado en hilo de traducción: %s", exc)

        msg = (
            f"[translator] {total_segs} segmentos procesados ({num_to_translate} traducidos en {total_batches} lotes paralelos con {workers} hilos; "
            f"{skipped_refs} referencias y {skipped_formulas} fórmulas preservadas intactas)."
        )
        context.setdefault("memory_log", []).append(msg)
        logger.info(msg)
        return context

    def retranslate_segment(self, seg: Segment, source_lang: str, target_lang: str, reason: str) -> None:
        """Reintenta un único segmento marcado como sospechoso por el validador."""
        try:
            seg.translated = self._translate_one(
                seg.original, source_lang, target_lang, retry_reason=reason
            )
            seg.status = "traducido"
        except Exception as exc:  # noqa: BLE001
            logger.error("[Translator] Fallo en reintento del segmento %d: %s", seg.id, exc)
            seg.status = "error"
            seg.validation_notes.append(f"Error en reintento: {exc}")
