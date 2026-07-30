"""
Módulo para construcción de queries de búsqueda según reglas de negocio.
"""
import re
from typing import List


# Caracteres prohibidos en códigos OEM
_FORBIDDEN_CHARS_RE = re.compile(r"[ \t\r\n/\-\,\.]")

# Palabras a remover en limpieza de queries
WORDS_TO_REMOVE = ["/RELE", " DEL", " TRA", " DCHO", " DCHA", " IZDO", " IZDA"]


def should_search(code: str) -> bool:
    """
    Determina si un código OEM debe ser buscado.

    Regla de negocio: Si el OEM contiene espacio, '/', '-', ',' o '.'
    => NO buscar (resultado = 0).

    Args:
        code: Código OEM a validar

    Returns:
        True si el código es válido para búsqueda, False en caso contrario

    Examples:
        >>> should_search("ABC123")
        True
        >>> should_search("ABC/123")
        False
        >>> should_search("ABC-123")
        False
        >>> should_search("")
        False
    """
    if not code:
        return False

    s = str(code).strip()
    if s == "":
        return False

    return _FORBIDDEN_CHARS_RE.search(s) is None


def clean_query_text(text: str) -> str:
    """
    Limpia el texto de búsqueda removiendo palabras específicas.

    Args:
        text: Texto a limpiar

    Returns:
        Texto limpio
    """
    result = text
    for word in WORDS_TO_REMOVE:
        result = result.replace(word, "")
    result = result.replace(".", "")
    return re.sub(r"\s+", " ", result).strip()


def build_query_from_sheets(pieza: str, oem: str) -> str:
    """
    Construye el query de búsqueda basado en reglas de Google Sheets.

    Replica la lógica de la fórmula original:
    - Casos especiales para MOTOR y CAJA
    - Prefijos especiales (MOTOR ARRA, MOTOR CALE, etc.)
    - Limpieza de texto automática

    Args:
        pieza: Nombre de la pieza (columna Pieza)
        oem: Código OEM

    Returns:
        Query text limpio para búsqueda en Ecooparts

    Examples:
        >>> build_query_from_sheets("MOTOR ARRANQUE", "ABC123")
        'MOTOR ARRANQUE ABC123'
        >>> build_query_from_sheets("CAJA CAMBIOS", "XYZ789")
        'CAJA CAMBIOS XYZ789'
    """
    B = (pieza or "").strip().upper()
    N = (oem or "").strip()

    if not B and not N:
        return ""

    # Caso especial: MOTOR (excepto subcasos especiales)
    if B.startswith("MOTOR") and not B.startswith("MOTOR ARRA") and not B.startswith("MOTOR CALE") and not B.startswith("MOTOR LIMPIAPARAB"):
        return clean_query_text(f"{B} {N}".strip())

    # Caso especial: CAJA CAMBIOS
    if B.startswith("CAJA CAMBIOS"):
        return f"CAJA CAMBIOS {N}".strip()

    # Caso especial: Otros CAJA
    if B.startswith("CAJA"):
        return clean_query_text(f"{B} {N}".strip())

    # Prefijos especiales que mantienen todo el texto
    special_prefixes = (
        B.startswith("MOTOR ARRA")
        or B.startswith("MOTOR CALE")
        or B.startswith("MOTOR LIMPIAPARAB")
        or B.startswith("CAJA DIRECCION")
        or B.startswith("CAJA FUSIBLES")
        or B.startswith("CAJA RELES")
    )

    if special_prefixes:
        return clean_query_text(f"{B} {N}".strip())

    # Si no tiene espacios, concatenar directamente
    if " " not in B:
        return f"{(B + ' ' + N).replace('/RELE', '')}".strip()

    # Caso general: tomar primera palabra
    first_word = B.split(" ", 1)[0]
    return clean_query_text(f"{first_word} {N}".strip())
