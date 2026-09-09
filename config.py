"""
config.py
---------
Configuración centralizada de la aplicación: credenciales, nombres de modelo,
umbrales de validación y parámetros de rate-limiting.

Todos los valores pueden sobreescribirse mediante variables de entorno
(ver .env.example), lo que facilita su despliegue en Docker sin tocar código.
"""

import os
from dataclasses import dataclass, field

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass


def _get_conf(key: str, default: str = "") -> str:
    """Obtiene configuración desde os.environ o st.secrets de Streamlit Cloud."""
    val = os.getenv(key)
    if val:
        return val
    try:
        import streamlit as st
        if hasattr(st, "secrets") and key in st.secrets:
            return str(st.secrets[key])
    except Exception:
        pass
    return default


@dataclass(frozen=True)
class Settings:
    # --- Credenciales / modelo ---
    GEMINI_API_KEY: str = _get_conf("GEMINI_API_KEY", "")
    GEMINI_MODEL: str = _get_conf("GEMINI_MODEL", "gemini-3.5-flash-lite")
    GEMINI_EMBEDDING_MODEL: str = _get_conf("GEMINI_EMBEDDING_MODEL", "models/text-embedding-004")

    # --- Idiomas (por defecto Paper en Inglés -> Español) ---
    DEFAULT_SOURCE_LANG: str = os.getenv("DEFAULT_SOURCE_LANG", "en")
    DEFAULT_TARGET_LANG: str = os.getenv("DEFAULT_TARGET_LANG", "es")

    # --- Validación (Agente Validador) ---
    LENGTH_DIFF_THRESHOLD: float = float(os.getenv("LENGTH_DIFF_THRESHOLD", "0.85"))  # ±85% (español es naturalmente más extenso)
    MAX_RETRIES: int = int(os.getenv("MAX_RETRIES", "1"))
    MIN_UNTRANSLATED_WORD_LEN: int = int(os.getenv("MIN_UNTRANSLATED_WORD_LEN", "4"))

    # --- Alineación ---
    ALIGNMENT_MODE: str = os.getenv("ALIGNMENT_MODE", "position")  # "position" | "semantic"
    SEMANTIC_SIMILARITY_THRESHOLD: float = float(os.getenv("SEMANTIC_SIMILARITY_THRESHOLD", "0.55"))

    # --- Rate limiting / backoff exponencial (14 req/min para ajustarse al Free Tier de 15 RPM sin 429) ---
    MAX_CALLS_PER_MINUTE: int = int(os.getenv("MAX_CALLS_PER_MINUTE", "14"))
    BACKOFF_BASE_SECONDS: float = float(os.getenv("BACKOFF_BASE_SECONDS", "2.0"))
    BACKOFF_MAX_RETRIES: int = int(os.getenv("BACKOFF_MAX_RETRIES", "5"))
    BACKOFF_MAX_SECONDS: float = float(os.getenv("BACKOFF_MAX_SECONDS", "60.0"))

    # --- Rendimiento y Concurrencia de Traducción ---
    DEFAULT_BATCH_SIZE: int = int(os.getenv("DEFAULT_BATCH_SIZE", "16"))  # 16 párrafos por llamada (aceleración 5x-8x)
    DEFAULT_MAX_WORKERS: int = int(os.getenv("DEFAULT_MAX_WORKERS", "3"))   # 3 hilos concurrentes

    # --- Segmentación de texto ---
    MAX_CHARS_PER_SEGMENT: int = int(os.getenv("MAX_CHARS_PER_SEGMENT", "1200"))

    # --- Colores para verificación cruzada (paleta cíclica) ---
    HIGHLIGHT_COLORS: tuple = field(default_factory=lambda: (
        "#FFD54F",  # amarillo
        "#81C784",  # verde
        "#64B5F6",  # azul
        "#F06292",  # rosa
        "#BA68C8",  # violeta
        "#FF8A65",  # naranja
        "#4DB6AC",  # turquesa
        "#A1887F",  # marrón
    ))

    # --- Archivos soportados ---
    SUPPORTED_EXTENSIONS: tuple = (".txt", ".docx", ".pdf")


settings = Settings()
