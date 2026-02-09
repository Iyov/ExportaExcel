"""
Tests para el módulo de validadores.
"""
import pytest
import pandas as pd
from validators import (
    validate_dataframe_not_empty,
    validate_required_columns,
    validate_numeric_value,
    validate_string_not_empty,
    validate_date_range
)


def test_validate_dataframe_not_empty():
    """Test para validación de DataFrame no vacío."""
    # DataFrame con datos
    df = pd.DataFrame({'col1': [1, 2, 3]})
    assert validate_dataframe_not_empty(df) is True
    
    # DataFrame vacío
    df_empty = pd.DataFrame()
    assert validate_dataframe_not_empty(df_empty) is False
    
    # None
    assert validate_dataframe_not_empty(None) is False


def test_validate_required_columns():
    """Test para validación de columnas requeridas."""
    df = pd.DataFrame({'col1': [1], 'col2': [2], 'col3': [3]})
    
    # Todas las columnas presentes
    assert validate_required_columns(df, ['col1', 'col2']) is True
    
    # Columna faltante
    assert validate_required_columns(df, ['col1', 'col4']) is False
    
    # DataFrame vacío
    assert validate_required_columns(pd.DataFrame(), ['col1']) is False


def test_validate_numeric_value():
    """Test para validación de valores numéricos."""
    # Valor válido
    assert validate_numeric_value(10) is True
    assert validate_numeric_value(10.5) is True
    
    # Con rango
    assert validate_numeric_value(10, min_value=5, max_value=15) is True
    assert validate_numeric_value(10, min_value=15) is False
    assert validate_numeric_value(10, max_value=5) is False
    
    # None
    assert validate_numeric_value(None, allow_none=True) is True
    assert validate_numeric_value(None, allow_none=False) is False
    
    # No numérico
    assert validate_numeric_value("abc") is False


def test_validate_string_not_empty():
    """Test para validación de strings no vacíos."""
    # String válido
    assert validate_string_not_empty("test") is True
    
    # String vacío
    assert validate_string_not_empty("") is False
    assert validate_string_not_empty("   ") is False
    
    # None
    assert validate_string_not_empty(None) is False
    
    # No string
    assert validate_string_not_empty(123) is False


def test_validate_date_range():
    """Test para validación de rangos de fechas."""
    # Rango válido
    assert validate_date_range("20220101", "20221231") is True
    
    # Rango inválido (desde > hasta)
    assert validate_date_range("20221231", "20220101") is False
    
    # Formato inválido
    assert validate_date_range("2022-01-01", "2022-12-31") is False
