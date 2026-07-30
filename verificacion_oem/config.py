"""
Configuración centralizada para Verificación OEM.
"""
from dataclasses import dataclass, field
from typing import List


@dataclass
class AppConfig:
    """Configuración general de la aplicación."""
    DEFAULT_BATCH: int = 500
    DEFAULT_DELAY: float = 0.15
    DEFAULT_AUTOSAVE_INTERVAL: int = 50

    # Nombres de columnas esperadas en Excel
    COL_PIEZA: List[str] = field(default_factory=lambda: ["Pieza"])
    COL_REF_ORIGINAL: List[str] = field(default_factory=lambda: ["Ref. Original (Concesionarios)"])
    COL_REF_PARALELO: List[str] = field(default_factory=lambda: ["Ref. Paralelo (Recambistas)"])
    COL_VALIDACION_ORIGINAL: str = "Validacion Original"
    COL_VALIDACION_PARALELO: str = "Validacion Paralelo"


@dataclass
class CounterConfig:
    """Configuración para el scraper de Ecooparts."""
    headless: bool = True
    timeout_ms: int = 30000
    proxy: str | None = None
    max_pages: int = 20
    # Límite superior de conteo que nos interesa (para optimizar trabajo)
    max_count: int = 30
    per_page: int = 30
    scroll_rounds: int = 10
    scroll_wait_ms: int = 800
    block_images: bool = True
