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

# Carga opcional de .env en entornos locales (en Docker se inyectan via env_file)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass


@dataclass(frozen=True)
class Settings:
    # --- Credenciales / modelo ---
    GEMINI_API_KEY: str = os.getenv("GEMINI_API_KEY", "")
    GEMINI_MODEL: str = os.getenv("GEMINI_MODEL", "gemini-3.1-flash-lite")
    GEMINI_EMBEDDING_MODEL: str = os.getenv("GEMINI_EMBEDDING_MODEL", "models/text-embedding-004")

    # --- Idiomas ---
    DEFAULT_SOURCE_LANG: str = os.getenv("DEFAULT_SOURCE_LANG", "auto")
    DEFAULT_TARGET_LANG: str = os.getenv("DEFAULT_TARGET_LANG", "es")

    # --- Validación (Agente Validador) ---
    LENGTH_DIFF_THRESHOLD: float = float(os.getenv("LENGTH_DIFF_THRESHOLD", "0.40"))  # ±40%
    MAX_RETRIES: int = int(os.getenv("MAX_RETRIES", "2"))
    MIN_UNTRANSLATED_WORD_LEN: int = int(os.getenv("MIN_UNTRANSLATED_WORD_LEN", "4"))

    # --- Alineación ---
    ALIGNMENT_MODE: str = os.getenv("ALIGNMENT_MODE", "position")  # "position" | "semantic"
    SEMANTIC_SIMILARITY_THRESHOLD: float = float(os.getenv("SEMANTIC_SIMILARITY_THRESHOLD", "0.55"))

    # --- Rate limiting / backoff exponencial ante 429 ---
    MAX_CALLS_PER_MINUTE: int = int(os.getenv("MAX_CALLS_PER_MINUTE", "30"))
    BACKOFF_BASE_SECONDS: float = float(os.getenv("BACKOFF_BASE_SECONDS", "2.0"))
    BACKOFF_MAX_RETRIES: int = int(os.getenv("BACKOFF_MAX_RETRIES", "5"))
    BACKOFF_MAX_SECONDS: float = float(os.getenv("BACKOFF_MAX_SECONDS", "60.0"))

    # --- Segmentación de texto ---
    MAX_CHARS_PER_SEGMENT: int = int(os.getenv("MAX_CHARS_PER_SEGMENT", "1200"))

    # --- Colores para verificación cruzada (paleta cíclica, mínimo 6) ---
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
