"""
Procesador principal que orquesta la validación de OEM.
"""
import time
import random
from pathlib import Path
from typing import Optional

import pandas as pd

from .config import AppConfig, CounterConfig
from .excel_handler import get_col, coerce_int_series
from .query_builder import should_search, build_query_from_sheets
from .scraper import EcoopartsCounter, PlaywrightTimeoutError
from .logger import (
    RunStats,
    utc_now_iso,
    append_jsonl,
    safe_int,
    load_resume_from_jsonl,
    save_summary
)
from tqdm import tqdm


def process(
    df: pd.DataFrame,
    original_path: Path,
    *,
    batch: int = 500,
    start: int = 0,
    delay: float = 0.15,
    autosave_interval: int = 50,
    counter_cfg: Optional[CounterConfig] = None,
    verbose: bool = False,
    out_dir: Optional[Path] = None,
    resume: bool = False,
) -> pd.DataFrame:
    """
    Procesa el DataFrame validando códigos OEM en Ecooparts.

    Args:
        df: DataFrame con los datos a procesar
        original_path: Ruta del archivo Excel original
        batch: Número de filas a procesar
        start: Fila inicial (0-indexed)
        delay: Delay base entre filas en segundos
        autosave_interval: Guardar parcial cada N filas (0 desactiva)
        counter_cfg: Configuración del scraper
        verbose: Mostrar logs detallados
        out_dir: Directorio para artifacts (por defecto junto al Excel)
        resume: Continuar desde última ejecución

    Returns:
        DataFrame con columnas Validacion Original y Validacion Paralelo
    """
    cfg = AppConfig()

    # Extraer columnas
    pieza = get_col(df, cfg.COL_PIEZA, 1)
    ref_original = get_col(df, cfg.COL_REF_ORIGINAL, 4)
    ref_paralelo = get_col(df, cfg.COL_REF_PARALELO, 5)

    n = len(df)

    # Inicializar columnas de salida
    if cfg.COL_VALIDACION_ORIGINAL in df.columns:
        out_orig = coerce_int_series(df[cfg.COL_VALIDACION_ORIGINAL])
    else:
        out_orig = pd.Series([0] * n)

    if cfg.COL_VALIDACION_PARALELO in df.columns:
        out_para = coerce_int_series(df[cfg.COL_VALIDACION_PARALELO])
    else:
        out_para = pd.Series([0] * n)

    # Configurar paths de salida
    base_name = original_path.stem
    folder = original_path.parent

    if out_dir is None:
        out_dir = folder / "run_artifacts"

    run_log_path = out_dir / "run_log.jsonl"
    summary_path = out_dir / "summary.json"

    # Inicializar estadísticas
    stats = RunStats(started_utc=utc_now_iso(), rows_total=n)
    t0_run = time.perf_counter()

    # Resume automático desde JSONL
    if resume:
        last_idx, values_by_row = load_resume_from_jsonl(run_log_path)
        if last_idx >= 0:
            # Restaurar valores previos
            for ri, (co, cp) in values_by_row.items():
                if 0 <= ri < n:
                    out_orig.iloc[ri] = int(co)
                    out_para.iloc[ri] = int(cp)

            new_start = last_idx + 1
            if new_start > start:
                print(f"✓ Resume: continuando desde fila {new_start+1}/{n}")
                start = new_start
            else:
                print(f"✓ Resume: último procesado={last_idx+1}, usando --start={start+1}")
        else:
            print("✓ Resume: sin progreso previo, iniciando desde cero")

    # Validar rango
    if start >= n:
        stats.ended_utc = utc_now_iso()
        stats.total_seconds = time.perf_counter() - t0_run
        save_summary(summary_path, stats)
        print(f"⚠ No hay filas para procesar (start={start} >= total={n})")
        df_out = df.copy()
        df_out[cfg.COL_VALIDACION_ORIGINAL] = out_orig
        df_out[cfg.COL_VALIDACION_PARALELO] = out_para
        return df_out

    end = min(n, start + batch)

    if counter_cfg is None:
        counter_cfg = CounterConfig()

    # Procesar filas (con barra de progreso)
    with EcoopartsCounter(counter_cfg) as counter:
        for i in tqdm(
            range(start, end),
            desc="Procesando",
            unit="fila",
            bar_format="{desc}: {percentage:3.0f}%|{bar}| {n_fmt}/{total_fmt}",
        ):
            row_t0 = time.perf_counter()

            p = str(pieza.iloc[i]).strip()
            e = str(ref_original.iloc[i]).strip()
            f = str(ref_paralelo.iloc[i]).strip()

            valid_orig = 0
            valid_para = 0

            status_orig = "SKIPPED"
            status_para = "SKIPPED"

            err_type = ""
            err_msg = ""

            try:
                # Procesar ORIGINAL
                if not e:
                    stats.queries_empty += 1
                    status_orig = "EMPTY"
                elif not should_search(e):
                    stats.queries_skipped_forbidden += 1
                    status_orig = "FORBIDDEN_CHARS"
                else:
                    query_o = build_query_from_sheets(p, e)
                    stats.queries_total += 1
                    status_orig = "OK"
                    valid_orig = counter.count(query_o, verbose=verbose)

                # Procesar PARALELO
                if not f:
                    stats.queries_empty += 1
                    status_para = "EMPTY"
                elif not should_search(f):
                    stats.queries_skipped_forbidden += 1
                    status_para = "FORBIDDEN_CHARS"
                else:
                    query_p = build_query_from_sheets(p, f)
                    stats.queries_total += 1
                    status_para = "OK"
                    valid_para = counter.count(query_p, verbose=verbose)

            except PlaywrightTimeoutError as ex:
                stats.timeouts += 1
                err_type = "TIMEOUT"
                err_msg = str(ex)[:400]
            except Exception as ex:
                stats.errors += 1
                err_type = "ERROR"
                err_msg = str(ex)[:400]

            # Normalizar / capear resultados a 0..max_count
            maxc = int(getattr(counter.cfg, "max_count", 0) or 0)
            if maxc <= 0:
                maxc = 30

            def _normalize(v: int) -> int:
                try:
                    vv = int(v)
                except Exception:
                    vv = 0
                if vv <= 0:
                    return 0
                return min(vv, maxc)

            out_val_orig = _normalize(valid_orig)
            out_val_para = _normalize(valid_para)

            # Guardar resultados
            out_orig.iloc[i] = int(out_val_orig)
            out_para.iloc[i] = int(out_val_para)

            row_seconds = time.perf_counter() - row_t0

            # Actualizar stats
            stats.rows_processed += 1
            stats.cache_hits = counter.cache_hits
            stats.cache_misses = counter.cache_misses

            # Log JSONL
            append_jsonl(run_log_path, {
                "ts_utc": utc_now_iso(),
                "row_index": i,
                "row_number": i + 1,
                "pieza": p,
                "oem_original": e,
                "oem_paralelo": f,
                "status_original": status_orig,
                "status_paralelo": status_para,
                # Guardar los valores normalizados (0..max_count)
                "count_original": safe_int(out_val_orig),
                "count_paralelo": safe_int(out_val_para),
                "row_seconds": round(row_seconds, 3),
                "cache_hits": counter.cache_hits,
                "cache_misses": counter.cache_misses,
                "error_type": err_type,
                "error_msg": err_msg,
            })

            print(
                f"Fila {i+1}/{n} → orig={out_val_orig} para={out_val_para} | "
                f"pieza='{p}' | E='{e}' | F='{f}'"
            )

            # Autosave
            processed = i - start + 1
            if autosave_interval and (processed % autosave_interval == 0 or i == end - 1):
                df_out = df.copy()
                df_out[cfg.COL_VALIDACION_ORIGINAL] = out_orig
                df_out[cfg.COL_VALIDACION_PARALELO] = out_para

                partial_name = folder / f"{base_name}_validated_partial_{start}_{i+1}.xlsx"
                try:
                    df_out.to_excel(partial_name, index=False)
                    print(f"💾 Autosave: {partial_name.name}")
                except Exception as ex:
                    print(f"⚠ Error en autosave: {ex}")

            # Delay con jitter
            if delay and delay > 0:
                time.sleep(max(0.0, float(delay)) + random.uniform(0.0, 0.15))

    # Guardar summary final
    stats.total_seconds = time.perf_counter() - t0_run
    stats.ended_utc = utc_now_iso()

    if stats.rows_processed > 0:
        stats.avg_seconds_per_row = stats.total_seconds / stats.rows_processed
        stats.rows_per_minute = (
            (stats.rows_processed / stats.total_seconds) * 60.0
            if stats.total_seconds > 0 else 0.0
        )

    save_summary(summary_path, stats)

    print(f"\n📊 Logs: {run_log_path}")
    print(f"📊 Summary: {summary_path}")

    # Retornar DataFrame final
    df_out = df.copy()
    df_out[cfg.COL_VALIDACION_ORIGINAL] = out_orig
    df_out[cfg.COL_VALIDACION_PARALELO] = out_para
    return df_out
