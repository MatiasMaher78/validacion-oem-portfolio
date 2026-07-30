"""
Módulo para lectura y escritura de archivos Excel.
"""
import os
from pathlib import Path
from typing import List

import pandas as pd


class ExcelValidationError(Exception):
    """Error en la validación del archivo Excel."""
    pass


def locate_unique_excel(folder: Path) -> Path:
    """
    Busca el único Excel de entrada en la carpeta especificada.

    Args:
        folder: Ruta a la carpeta donde buscar

    Returns:
        Path al archivo Excel encontrado

    Raises:
        FileNotFoundError: Si no se encuentra ningún Excel
        ExcelValidationError: Si hay múltiples archivos o solo temporales
    """
    if not folder.exists():
        raise FileNotFoundError(f"La carpeta no existe: {folder}")

    candidates = [
        f for f in folder.iterdir()
        if f.suffix.lower() in ('.xls', '.xlsx')
        and not f.name.startswith('~$')
        and not f.name.startswith('.')
    ]

    # Filtrar archivos de salida del script
    preferred = [f for f in candidates if "_validated" not in f.name.lower()]
    if preferred:
        candidates = preferred

    if not candidates:
        raise FileNotFoundError(
            f"No se encontró ningún archivo Excel de entrada en {folder}"
        )

    if len(candidates) > 1:
        raise ExcelValidationError(
            f"Se encontraron varios archivos Excel en {folder}: "
            f"{[f.name for f in candidates]}. "
            "Deja solo 1 archivo de entrada."
        )

    return candidates[0]


def read_workbook(path: Path) -> pd.DataFrame:
    """
    Lee un archivo Excel y retorna un DataFrame.

    Args:
        path: Ruta al archivo Excel

    Returns:
        DataFrame con los datos del Excel
    """
    return pd.read_excel(path, sheet_name=0, dtype=str).fillna("")


def get_col(df: pd.DataFrame, names: List[str], fallback_idx: int) -> pd.Series:
    """
    Obtiene una columna del DataFrame por nombre o índice de respaldo.

    Args:
        df: DataFrame de pandas
        names: Lista de nombres posibles para la columna
        fallback_idx: Índice de columna a usar si no se encuentra por nombre

    Returns:
        Serie con los datos de la columna
    """
    cols = {c.lower(): c for c in df.columns}

    for n in names:
        if n.lower() in cols:
            return df[cols[n.lower()]].astype(str).fillna("")

    if df.shape[1] > fallback_idx:
        return df.iloc[:, fallback_idx].astype(str).fillna("")

    return pd.Series([""] * len(df))


def validate_dataframe(df: pd.DataFrame) -> List[str]:
    """
    Valida que el DataFrame contenga las columnas mínimas requeridas.

    Args:
        df: DataFrame a validar

    Returns:
        Lista de mensajes de error (vacía si todo OK)
    """
    errors = []

    if df.empty:
        errors.append("El archivo Excel está vacío")
        return errors

    # Verificar que tenga al menos 6 columnas (índices 0-5)
    if df.shape[1] < 6:
        errors.append(
            f"El Excel debe tener al menos 6 columnas, "
            f"se encontraron {df.shape[1]}"
        )

    return errors


def save_results(original_path: Path, df_out: pd.DataFrame) -> Path:
    """
    Guarda los resultados validados en un nuevo archivo Excel.

    Args:
        original_path: Ruta del archivo Excel original
        df_out: DataFrame con resultados

    Returns:
        Path del archivo guardado
    """
    folder = original_path.parent
    base = original_path.stem
    out_path = folder / f"{base}_validated.xlsx"

    df_out.to_excel(out_path, index=False)
    return out_path


def coerce_int_series(s: pd.Series) -> pd.Series:
    """
    Convierte una serie a enteros, usando 0 para valores inválidos.

    Args:
        s: Serie de pandas

    Returns:
        Serie con valores enteros
    """
    out = []
    for v in s.astype(str).fillna("").tolist():
        v = v.strip()
        if v == "":
            out.append(0)
        else:
            try:
                out.append(int(float(v)))
            except Exception:
                out.append(0)
    return pd.Series(out)
