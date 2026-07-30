"""
Tests unitarios para Verificación OEM.

Ejecutar con: pytest test_verificacion_oem.py -v
"""
import pytest
import pandas as pd
from pathlib import Path

from verificacion_oem.query_builder import should_search, build_query_from_sheets, clean_query_text
from verificacion_oem.excel_handler import get_col, coerce_int_series, validate_dataframe
from verificacion_oem.scraper import build_ecooparts_search_url


class TestShouldSearch:
    """Tests para la validación de códigos OEM."""

    def test_empty_code_returns_false(self):
        """Códigos vacíos no deben ser buscados."""
        assert should_search("") is False
        assert should_search("   ") is False
        assert should_search(None) is False

    def test_forbidden_chars_return_false(self):
        """Códigos con caracteres prohibidos no deben ser buscados."""
        assert should_search("ABC/123") is False
        assert should_search("ABC-123") is False
        assert should_search("ABC 123") is False
        assert should_search("ABC.123") is False
        assert should_search("ABC,123") is False

    def test_valid_code_returns_true(self):
        """Códigos válidos deben pasar la validación."""
        assert should_search("ABC123") is True
        assert should_search("XYZ789") is True
        assert should_search("12345") is True


class TestCleanQueryText:
    """Tests para limpieza de texto de queries."""

    def test_removes_special_words(self):
        """Debe remover palabras específicas."""
        assert clean_query_text("MOTOR/RELE ABC") == "MOTOR ABC"
        assert clean_query_text("FARO DEL DERECHO") == "FARO DERECHO"
        assert clean_query_text("LUZ TRA IZDA") == "LUZ"

    def test_removes_dots(self):
        """Debe remover puntos."""
        assert clean_query_text("A.B.C") == "ABC"

    def test_normalizes_whitespace(self):
        """Debe normalizar espacios múltiples."""
        assert clean_query_text("ABC    123") == "ABC 123"
        assert clean_query_text("  ABC  ") == "ABC"


class TestBuildQueryFromSheets:
    """Tests para construcción de queries según reglas de negocio."""

    def test_empty_inputs_return_empty(self):
        """Entradas vacías deben retornar string vacío."""
        assert build_query_from_sheets("", "") == ""
        assert build_query_from_sheets("  ", "  ") == ""

    def test_motor_arranque(self):
        """MOTOR ARRANQUE debe mantener el texto completo."""
        result = build_query_from_sheets("MOTOR ARRANQUE", "ABC123")
        assert "MOTOR ARRANQUE" in result
        assert "ABC123" in result

    def test_motor_calefaccion(self):
        """MOTOR CALEFACCION debe mantener el texto completo."""
        result = build_query_from_sheets("MOTOR CALEFACCION", "XYZ789")
        assert "MOTOR CALEFACCION" in result
        assert "XYZ789" in result

    def test_motor_generic(self):
        """MOTOR genérico debe procesar normalmente."""
        result = build_query_from_sheets("MOTOR", "123ABC")
        assert "MOTOR" in result
        assert "123ABC" in result

    def test_caja_cambios(self):
        """CAJA CAMBIOS debe tener formato especial."""
        result = build_query_from_sheets("CAJA CAMBIOS", "AAA111")
        assert result == "CAJA CAMBIOS AAA111"

    def test_single_word_pieza(self):
        """Pieza de una sola palabra concatena con OEM."""
        result = build_query_from_sheets("FARO", "ZZZ999")
        assert "FARO" in result
        assert "ZZZ999" in result

    def test_multi_word_pieza(self):
        """Pieza multi-palabra usa solo primera palabra."""
        result = build_query_from_sheets("FARO DELANTERO DERECHO", "BBB222")
        # Debe usar solo "FARO" + OEM
        words = result.split()
        assert "FARO" in words
        assert "BBB222" in words
        assert "DELANTERO" not in result


class TestGetCol:
    """Tests para extracción de columnas de DataFrames."""

    def test_finds_column_by_name(self):
        """Debe encontrar columna por nombre (case-insensitive)."""
        df = pd.DataFrame({
            "Pieza": ["A", "B"],
            "Ref": ["1", "2"]
        })
        result = get_col(df, ["Pieza"], 0)
        assert list(result) == ["A", "B"]

    def test_case_insensitive(self):
        """Búsqueda debe ser case-insensitive."""
        df = pd.DataFrame({"PIEZA": ["X", "Y"]})
        result = get_col(df, ["pieza"], 0)
        assert list(result) == ["X", "Y"]

    def test_uses_fallback_index(self):
        """Debe usar índice de respaldo si no encuentra por nombre."""
        df = pd.DataFrame({
            "Col0": ["A", "B"],
            "Col1": ["C", "D"]
        })
        result = get_col(df, ["NoExiste"], 1)
        assert list(result) == ["C", "D"]

    def test_returns_empty_if_invalid_fallback(self):
        """Debe retornar serie vacía si índice fallback es inválido."""
        df = pd.DataFrame({"Col0": ["A", "B"]})
        result = get_col(df, ["NoExiste"], 5)
        assert len(result) == 2
        assert all(result == "")


class TestCoerceIntSeries:
    """Tests para conversión de series a enteros."""

    def test_converts_valid_ints(self):
        """Debe convertir enteros válidos."""
        s = pd.Series(["1", "2", "3"])
        result = coerce_int_series(s)
        assert list(result) == [1, 2, 3]

    def test_converts_floats(self):
        """Debe convertir floats a enteros."""
        s = pd.Series(["1.5", "2.9", "3.1"])
        result = coerce_int_series(s)
        assert list(result) == [1, 2, 3]

    def test_handles_empty_strings(self):
        """Strings vacíos deben convertirse a 0."""
        s = pd.Series(["1", "", "3"])
        result = coerce_int_series(s)
        assert list(result) == [1, 0, 3]

    def test_handles_invalid_values(self):
        """Valores inválidos deben convertirse a 0."""
        s = pd.Series(["1", "abc", "3"])
        result = coerce_int_series(s)
        assert list(result) == [1, 0, 3]


class TestValidateDataframe:
    """Tests para validación de DataFrames."""

    def test_empty_dataframe_returns_error(self):
        """DataFrame vacío debe retornar error."""
        df = pd.DataFrame()
        errors = validate_dataframe(df)
        assert len(errors) > 0
        assert "vacío" in errors[0].lower()

    def test_insufficient_columns_returns_error(self):
        """DataFrame con pocas columnas debe retornar error."""
        df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
        errors = validate_dataframe(df)
        assert len(errors) > 0
        assert "columnas" in errors[0].lower()

    def test_valid_dataframe_returns_no_errors(self):
        """DataFrame válido no debe retornar errores."""
        df = pd.DataFrame({
            f"Col{i}": [1, 2] for i in range(7)
        })
        errors = validate_dataframe(df)
        assert len(errors) == 0


class TestBuildEcoopartsSearchUrl:
    """Tests para construcción de URLs de Ecooparts."""

    def test_builds_valid_url(self):
        """Debe construir URL válida."""
        url = build_ecooparts_search_url("MOTOR ABC123")
        assert url.startswith("https://ecooparts.com")
        assert "pag=pro" in url
        assert "tebu=" in url

    def test_includes_page_parameter(self):
        """Debe incluir parámetro de página."""
        url = build_ecooparts_search_url("TEST", page=5)
        assert "panu=" in url

    def test_includes_per_page_parameter(self):
        """Debe incluir parámetro de resultados por página."""
        url = build_ecooparts_search_url("TEST", per_page=50)
        assert "qregx=" in url

    def test_encodes_query_in_base64(self):
        """Query debe estar codificado en base64."""
        url = build_ecooparts_search_url("MOTOR")
        # Debe contener valores base64 (caracteres [A-Za-z0-9+/=])
        import re
        base64_pattern = r'[A-Za-z0-9+/=]+'
        assert re.search(base64_pattern, url)


# Fixture para datos de prueba
@pytest.fixture
def sample_dataframe():
    """DataFrame de ejemplo para tests."""
    return pd.DataFrame({
        "Pieza": ["MOTOR ARRANQUE", "FARO", "CAJA CAMBIOS"],
        "Marca": ["AUDI", "BMW", "VW"],
        "Modelo": ["A4", "X5", "GOLF"],
        "Version": ["2.0", "3.0", "1.6"],
        "Ref. Original (Concesionarios)": ["ABC123", "DEF456", "GHI789"],
        "Ref. Paralelo (Recambistas)": ["XXX111", "YYY222", "ZZZ333"],
    })


def test_sample_dataframe_structure(sample_dataframe):
    """Verifica que el DataFrame de prueba tenga la estructura correcta."""
    assert len(sample_dataframe) == 3
    assert "Pieza" in sample_dataframe.columns
    assert "Ref. Original (Concesionarios)" in sample_dataframe.columns
