"""
Sistema de logging y métricas para el proceso de validación.
"""
import json
import os
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Tuple


@dataclass
class RunStats:
    """Estadísticas de ejecución del proceso."""
    started_utc: str
    ended_utc: str = ""
    rows_total: int = 0
    rows_processed: int = 0

    queries_total: int = 0
    queries_skipped_forbidden: int = 0
    queries_empty: int = 0

    cache_hits: int = 0
    cache_misses: int = 0

    timeouts: int = 0
    errors: int = 0

    total_seconds: float = 0.0
    avg_seconds_per_row: float = 0.0
    rows_per_minute: float = 0.0

    def to_dict(self) -> dict:
        """Convierte stats a diccionario."""
        return self.__dict__


def utc_now_iso() -> str:
    """Retorna timestamp UTC en formato ISO."""
    return datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


def append_jsonl(path: Path, obj: dict) -> None:
    """
    Agrega una línea JSON al archivo log.

    Args:
        path: Ruta al archivo JSONL
        obj: Diccionario a escribir
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "a", encoding="utf-8") as f:
        f.write(json.dumps(obj, ensure_ascii=False) + "\n")


def safe_int(x) -> int:
    """Convierte a int de forma segura, retorna 0 si falla."""
    try:
        return int(x)
    except Exception:
        return 0


def load_resume_from_jsonl(run_log_path: Path) -> Tuple[int, Dict[int, Tuple[int, int]]]:
    """
    Carga el estado de un run previo desde JSONL para resume.

    Args:
        run_log_path: Ruta al archivo run_log.jsonl

    Returns:
        Tupla con:
        - last_row_index: Último índice procesado (-1 si no hay)
        - values_by_row: Dict {row_index: (count_orig, count_para)}
    """
    if not run_log_path.exists():
        return -1, {}

    last_idx = -1
    values: Dict[int, Tuple[int, int]] = {}

    with open(run_log_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue

            try:
                obj = json.loads(line)
            except Exception:
                continue

            ri = obj.get("row_index")
            if ri is None:
                continue

            try:
                ri = int(ri)
            except Exception:
                continue

            co = safe_int(obj.get("count_original", 0))
            cp = safe_int(obj.get("count_paralelo", 0))
            values[ri] = (co, cp)

            if ri > last_idx:
                last_idx = ri

    return last_idx, values


def save_summary(path: Path, stats: RunStats) -> None:
    """
    Guarda el resumen de ejecución en JSON.

    Args:
        path: Ruta donde guardar el summary.json
        stats: Estadísticas a guardar
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(stats.to_dict(), f, ensure_ascii=False, indent=2)
