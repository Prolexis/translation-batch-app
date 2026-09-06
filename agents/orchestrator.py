"""
agents/orchestrator.py
--------------------------
ORQUESTADOR DE AGENTES
--------------------------
Coordina la ejecución secuencial de:
    ExtractorAgent -> TranslatorAgent -> ValidatorAgent -> AlignerAgent
"""

import logging
import time
from typing import Callable, Optional

from config import settings
from agents.extractor_agent import ExtractorAgent
from agents.translator_agent import TranslatorAgent
from agents.validator_agent import ValidatorAgent
from agents.aligner_agent import AlignerAgent

logger = logging.getLogger("translation_app.orchestrator")


class ConversationBufferMemory:
    """Buffer de memoria para registrar la trazabilidad y eventos del pipeline."""
    def __init__(self, return_messages: bool = True):
        self.return_messages = return_messages
        self.messages: list[str] = []

    def save_context(self, inputs: dict, outputs: dict):
        inp = inputs.get("input", "")
        out = outputs.get("output", "")
        self.messages.append(f"{inp}: {out}")

    def get_transcript(self) -> str:
        return "\n".join(self.messages)


ProgressCallback = Optional[Callable[[str, float], None]]


class TranslationOrchestrator:
    def __init__(self, source_lang: str = None, target_lang: str = None,
                 alignment_mode: str = None, api_key: str = None):
        self.source_lang = source_lang or settings.DEFAULT_SOURCE_LANG
        self.target_lang = target_lang or settings.DEFAULT_TARGET_LANG
        self.alignment_mode = alignment_mode or settings.ALIGNMENT_MODE
        self.api_key = api_key or settings.GEMINI_API_KEY

        # Memoria compartida entre agentes
        self.memory = ConversationBufferMemory(return_messages=True)

        self.extractor = ExtractorAgent()
        self.translator = TranslatorAgent(api_key=self.api_key)
        self.validator = ValidatorAgent()
        self.aligner = AlignerAgent(api_key=self.api_key)

    def _log_to_memory(self, context: dict, filename: str):
        for line in context.get("memory_log", []):
            self.memory.save_context(
                {"input": f"[{filename}] evento"},
                {"output": line},
            )

    def process_file(self, filename: str, file_bytes: bytes,
                      on_progress: ProgressCallback = None) -> dict:
        context = {
            "filename": filename,
            "file_bytes": file_bytes,
            "source_lang": self.source_lang,
            "target_lang": self.target_lang,
            "alignment_mode": self.alignment_mode,
            "translator_agent": self.translator,
            "memory_log": [],
            "started_at": time.time(),
        }

        stages = [
            ("Extrayendo texto", self.extractor.run, 0.25),
            ("Traduciendo con Gemini", self.translator.run, 0.60),
            ("Validando traducción", lambda ctx: self.validator.run(ctx, self.translator), 0.85),
            ("Alineando segmentos", self.aligner.run, 1.0),
        ]

        for stage_name, fn, progress_fraction in stages:
            if context.get("error") and stage_name != "Extrayendo texto":
                break
            try:
                context = fn(context)
            except Exception as exc:  # noqa: BLE001
                logger.exception("Fallo inesperado en etapa '%s' para %s", stage_name, filename)
                context["error"] = f"Error en etapa '{stage_name}': {exc}"
                break
            if on_progress:
                on_progress(stage_name, progress_fraction)

        context["finished_at"] = time.time()
        context["duration_seconds"] = context["finished_at"] - context["started_at"]
        self._log_to_memory(context, filename)

        if context.get("error"):
            context["file_status"] = "error"
        elif context.get("error_rate", 0) > 0.10:
            context["file_status"] = "traducido_con_advertencias"
        else:
            context["file_status"] = "traducido"

        return context

    def get_memory_transcript(self) -> str:
        return self.memory.get_transcript()
