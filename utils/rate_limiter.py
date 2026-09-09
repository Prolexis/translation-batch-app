"""
utils/rate_limiter.py
----------------------
Control de límites de la API de Gemini:

1. `RateLimiter`: limita el número de llamadas por minuto (token bucket simple),
   evitando disparar 429 por exceso de throughput.
2. `with_backoff`: decorador que reintenta una función con backoff exponencial
   + jitter cuando la API responde 429 (RESOURCE_EXHAUSTED) o errores
   transitorios de red/timeout.

Ambos mecanismos son agnósticos del cliente HTTP: solo inspeccionan el texto
de la excepción para detectar "429" / "rate limit" / "quota", por lo que
funcionan tanto con la librería oficial de Google como con LangChain.
"""

import random
import time
import threading
import functools
import logging

from config import settings

logger = logging.getLogger("translation_app.rate_limiter")


class RateLimiter:
    """Limita las llamadas salientes a N por minuto (thread-safe) con espaciado suave anti-ráfagas."""

    def __init__(self, max_calls_per_minute: int = None):
        self.max_calls = max_calls_per_minute or settings.MAX_CALLS_PER_MINUTE
        self._lock = threading.Lock()
        self._timestamps: list[float] = []
        self._last_call_time: float = 0.0

    def acquire(self):
        with self._lock:
            now = time.time()
            # 1. Descartar timestamps fuera de la ventana de 60s
            self._timestamps = [t for t in self._timestamps if now - t < 60.0]
            if len(self._timestamps) >= self.max_calls:
                sleep_for = 60.0 - (now - self._timestamps[0]) + 0.2
                logger.info("Rate limit local alcanzado (%d llamadas en 60s), esperando %.2fs", len(self._timestamps), sleep_for)
                time.sleep(max(sleep_for, 0.1))
                now = time.time()
                self._timestamps = [t for t in self._timestamps if now - t < 60.0]

            # 2. Espaciado uniforme mínimo entre llamadas sucesivas (evita ráfagas concurrentes que activan 429)
            # Para 14 RPM max en Free Tier, espaciado de ~4.3s
            min_interval = 60.0 / max(self.max_calls, 1)
            if self._last_call_time > 0:
                elapsed = now - self._last_call_time
                if elapsed < min_interval:
                    wait_spacing = min_interval - elapsed
                    time.sleep(wait_spacing)
                    now = time.time()

            self._last_call_time = now
            self._timestamps.append(now)


# Instancia compartida por toda la app (un único cliente Gemini)
global_rate_limiter = RateLimiter()


def _is_retryable_error(exc: Exception) -> bool:
    msg = str(exc).lower()
    keywords = [
        "429", "resource_exhausted", "rate limit", "quota",
        "deadline exceeded", "timeout", "unavailable", "503",
    ]
    return any(k in msg for k in keywords)


def with_backoff(max_retries: int = None, base_seconds: float = None, max_seconds: float = None):
    """
    Decorador de reintento con backoff exponencial + jitter.
    Se activa únicamente ante errores considerados transitorios/cuota (ver
    `_is_retryable_error`); cualquier otro error se propaga de inmediato.
    """
    max_retries = max_retries if max_retries is not None else settings.BACKOFF_MAX_RETRIES
    base_seconds = base_seconds if base_seconds is not None else settings.BACKOFF_BASE_SECONDS
    max_seconds = max_seconds if max_seconds is not None else settings.BACKOFF_MAX_SECONDS

    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            attempt = 0
            while True:
                try:
                    global_rate_limiter.acquire()
                    return func(*args, **kwargs)
                except Exception as exc:  # noqa: BLE001
                    attempt += 1
                    if attempt > max_retries or not _is_retryable_error(exc):
                        logger.error("Fallo definitivo tras %d intentos: %s", attempt, exc)
                        raise
                    delay = min(base_seconds * (2 ** (attempt - 1)), max_seconds)
                    delay += random.uniform(0, delay * 0.25)  # jitter
                    logger.warning(
                        "Error transitorio (%s). Reintento %d/%d en %.1fs",
                        exc, attempt, max_retries, delay,
                    )
                    time.sleep(delay)
        return wrapper
    return decorator
