import os
import time
import argparse
import random
import string
import re
import base64
from dataclasses import dataclass
from typing import Dict, Optional, Set

import pandas as pd

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
except Exception:
    sync_playwright = None  # type: ignore
    PlaywrightTimeoutError = Exception  # type: ignore


DEFAULT_BATCH = 500


# ----------------------------
# Excel I/O
# ----------------------------
def locate_unique_excel(folder: str) -> str:
    """
    Busca el Excel "de entrada" dentro de la carpeta.
    - Ignora temporales (~$)
    - Ignora salidas generadas por el script (*_validated*.xlsx)
    """
    candidates = [f for f in os.listdir(folder) if f.lower().endswith((".xls", ".xlsx"))]
    files = [f for f in candidates if not f.startswith("~$") and not f.startswith(".")]

    # prefer input-like files (exclude validated outputs)
    preferred = [f for f in files if "_validated" not in f.lower()]
    if preferred:
        files = preferred

    if not files and candidates:
        raise FileExistsError(
            f"Se detectaron sólo archivos temporales o salidas del script en {folder}: {candidates}. "
            "Cierra el archivo en Excel o mueve/renombra los _validated*.xlsx."
        )

    if not files:
        raise FileNotFoundError(f"No se encontró ningún .xls/.xlsx de entrada en {folder}")

    if len(files) > 1:
        raise FileExistsError(
            f"Se encontraron varios Excel posibles en {folder}: {files}. "
            "Deja sólo 1 archivo de entrada (o mueve los demás)."
        )

    return os.path.join(folder, files[0])


def read_workbook(path: str) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name=0, dtype=str).fillna("")


def get_col(df: pd.DataFrame, names, fallback_idx: int) -> pd.Series:
    cols = {c.lower(): c for c in df.columns}
    for n in names:
        if n.lower() in cols:
            return df[cols[n.lower()]].astype(str).fillna("")
    if df.shape[1] > fallback_idx:
        return df.iloc[:, fallback_idx].astype(str).fillna("")
    return pd.Series([""] * len(df))


# ----------------------------
# OEM "forbidden chars" rule (USER RULE)
# ----------------------------
_FORBIDDEN_CHARS_RE = re.compile(r"[ \t\r\n/\-\,\.]")


def should_search(code: str) -> bool:
    """
    Regla pedida (se aplica SOLO al OEM, no al texto de búsqueda completo):
    - Si el OEM contiene espacio, '/', '-', ',' o '.' => NO buscar (resultado = 0).
    """
    if not code:
        return False
    s = str(code)
    if s.strip() == "":
        return False
    return _FORBIDDEN_CHARS_RE.search(s) is None


# ----------------------------
# Google Sheets formula (pieza + oem) -> query text
# ----------------------------
def build_query_from_sheets(pieza: str, oem: str) -> str:
    """
    Replica 1:1 la fórmula de Google Sheets que pasaste.

    B = pieza (Export_Ref_CATe!$B2)
    N = oem   (Export_Ref_CATe!$N2)

    Devuelve el texto final a buscar en Ecooparts.
    """
    B = (pieza or "").strip().upper()
    N = (oem or "").strip()

    if not B and not N:
        return ""

    def clean_full(text: str) -> str:
        s = text
        s = s.replace("/RELE", "")
        s = s.replace(" DEL", "")
        s = s.replace(" TRA", "")
        s = s.replace(" DCHO", "")
        s = s.replace(" DCHA", "")
        s = s.replace(" IZDO", "")
        s = s.replace(" IZDA", "")
        s = s.replace(".", "")
        s = re.sub(r"\s+", " ", s).strip()
        return s

    # SI(O(IZQUIERDA(B;5)="MOTOR";IZQUIERDA(B;4)="CAJA"); ... )
    if B.startswith("MOTOR") or B.startswith("CAJA"):
        # SI(IZQUIERDA(B;12)="CAJA CAMBIOS";"CAJA CAMBIOS "&N; clean_full(B&" "&N))
        if B.startswith("CAJA CAMBIOS"):
            return f"CAJA CAMBIOS {N}".strip()
        return clean_full(f"{B} {N}".strip())

    # ELSE:
    special_prefixes = (
        B.startswith("MOTOR ARRA")
        or B.startswith("MOTOR CALE")
        or B.startswith("MOTOR LIMPIAPARAB")
        or B.startswith("CAJA CAMBIOS")
        or B.startswith("CAJA DIRECCION")
        or B.startswith("CAJA FUSIBLES")
        or B.startswith("CAJA RELES")
    )

    if special_prefixes:
        return clean_full(f"{B} {N}".strip())

    if " " not in B:
        # Solo quita /RELE en este caso (igual que Sheets)
        return f"{(B + ' ' + N).replace('/RELE', '')}".strip()

    first_word = B.split(" ", 1)[0]
    return clean_full(f"{first_word} {N}".strip())


# ----------------------------
# Ecooparts URL builder
# ----------------------------
def _b64(s: str) -> str:
    return base64.b64encode(s.encode("utf-8")).decode("ascii")


def build_ecooparts_search_url(query_text: str, *, page: int = 1, per_page: int = 30) -> str:
    """
    Construye una URL como las reales (incluye token 'toen').

    IMPORTANTE: 'query_text' puede tener espacios (pieza + OEM),
    y está OK porque se codifica dentro de base64.
    """
    c = str(query_text).strip()
    token = "".join(random.choices(string.ascii_lowercase + string.digits, k=22))

    params = {
        "pag": "pro",
        "busval": _b64(f"|{c}|ninguno|producto|-1|0|0|0|0||0|0|0|0"),
        "filval": "",
        "panu": _b64(str(page)),
        "tebu": _b64(c),
        "ord": _b64("ninguno"),
        "valo": _b64("-1"),
        "ubic": "",
        "toen": _b64(token),
        "veid": _b64("0"),
        "qregx": _b64(str(per_page)),
        "tmin": _b64("1"),
        "ttseu": "",
        "txbu": _b64(c),
        "ivevh": "",
        "ivevhmat": "",
        "ivevhsel": "",
        "ivevhcsver": "",
        "ivevhse": "",
        "oem": "",
        "vin": "",
    }

    ordered_keys = [
        "pag", "busval", "filval", "panu", "tebu", "ord", "valo", "ubic", "toen", "veid",
        "qregx", "tmin", "ttseu", "txbu", "ivevh", "ivevhmat", "ivevhsel", "ivevhcsver",
        "ivevhse", "oem", "vin"
    ]
    query = "&".join(f"{k}={params[k]}" for k in ordered_keys)
    return f"https://ecooparts.com/recambios-automovil-segunda-mano/?{query}"
