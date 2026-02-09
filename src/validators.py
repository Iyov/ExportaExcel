"""
Validadores para datos y configuraciones.
"""
import pandas
from pathlib import Path
from typing import Optional, List
from src.logger_config import setup_logger

logger = setup_logger('validators')


def validate_dataframe_not_empty(df: pandas.DataFrame, name: str = "DataFrame") -> bool:
    """
    Valida que un DataFrame no esté vacío.
    
    Args:
        df: DataFrame a validar
        name: Nombre del DataFrame para logging
        
    Returns:
        True si tiene datos, False si está vacío
    """
    if df is None:
        logger.warning(f"{name} es None")
        return False
    
    if df.empty:
        logger.warning(f"{name} está vacío")
        return False
    
    logger.debug(f"{name} tiene {len(df)} filas")
    return True


def validate_required_columns(
    df: pandas.DataFrame, 
    required_columns: List[str],
    name: str = "DataFrame"
) -> bool:
    """
    Valida que un DataFrame tenga las columnas requeridas.
    
    Args:
        df: DataFrame a validar
        required_columns: Lista de nombres de columnas requeridas
        name: Nombre del DataFrame para logging
        
    Returns:
        True si tiene todas las columnas, False en caso contrario
    """
    if df is None or df.empty:
        logger.error(f"{name} está vacío, no se pueden validar columnas")
        return False
    
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        logger.error(f"{name} no tiene las columnas: {missing_columns}")
        return False
    
    logger.debug(f"{name} tiene todas las columnas requeridas")
    return True


def validate_file_exists(file_path: str) -> bool:
    """
    Valida que un archivo exista.
    
    Args:
        file_path: Ruta del archivo
        
    Returns:
        True si existe, False en caso contrario
    """
    path = Path(file_path)
    
    if not path.exists():
        logger.error(f"Archivo no encontrado: {file_path}")
        return False
    
    if not path.is_file():
        logger.error(f"La ruta no es un archivo: {file_path}")
        return False
    
    logger.debug(f"Archivo validado: {file_path}")
    return True


def validate_date_range(desde: str, hasta: str) -> bool:
    """
    Valida que un rango de fechas sea válido.
    
    Args:
        desde: Fecha inicial (formato YYYYMMDD)
        hasta: Fecha final (formato YYYYMMDD)
        
    Returns:
        True si el rango es válido, False en caso contrario
    """
    try:
        fecha_desde = pandas.to_datetime(desde, format='%Y%m%d')
        fecha_hasta = pandas.to_datetime(hasta, format='%Y%m%d')
        
        if fecha_desde > fecha_hasta:
            logger.error(f"Fecha 'desde' ({desde}) es posterior a 'hasta' ({hasta})")
            return False
        
        logger.debug(f"Rango de fechas válido: {desde} - {hasta}")
        return True
        
    except Exception as e:
        logger.error(f"Error validando fechas: {e}")
        return False


def validate_numeric_value(
    value: any,
    name: str = "valor",
    allow_none: bool = False,
    min_value: Optional[float] = None,
    max_value: Optional[float] = None
) -> bool:
    """
    Valida que un valor sea numérico y esté en el rango esperado.
    
    Args:
        value: Valor a validar
        name: Nombre del valor para logging
        allow_none: Si True, permite valores None
        min_value: Valor mínimo permitido (opcional)
        max_value: Valor máximo permitido (opcional)
        
    Returns:
        True si es válido, False en caso contrario
    """
    if value is None:
        if allow_none:
            return True
        logger.error(f"{name} es None y no se permite")
        return False
    
    try:
        num_value = float(value)
        
        if min_value is not None and num_value < min_value:
            logger.error(f"{name} ({num_value}) es menor que el mínimo ({min_value})")
            return False
        
        if max_value is not None and num_value > max_value:
            logger.error(f"{name} ({num_value}) es mayor que el máximo ({max_value})")
            return False
        
        return True
        
    except (ValueError, TypeError) as e:
        logger.error(f"{name} no es un valor numérico válido: {value}")
        return False


def validate_string_not_empty(value: str, name: str = "valor") -> bool:
    """
    Valida que un string no esté vacío.
    
    Args:
        value: String a validar
        name: Nombre del valor para logging
        
    Returns:
        True si no está vacío, False en caso contrario
    """
    if value is None:
        logger.error(f"{name} es None")
        return False
    
    if not isinstance(value, str):
        logger.error(f"{name} no es un string: {type(value)}")
        return False
    
    if not value.strip():
        logger.error(f"{name} está vacío")
        return False
    
    return True
