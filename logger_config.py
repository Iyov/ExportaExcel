"""
Configuración centralizada de logging para el proyecto.
"""
import logging
import sys
from pathlib import Path
from datetime import datetime


def setup_logger(name: str = 'ExportaExcel', log_level: int = logging.INFO) -> logging.Logger:
    """
    Configura y retorna un logger con formato consistente.
    
    Args:
        name: Nombre del logger
        log_level: Nivel de logging (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        
    Returns:
        Logger configurado
    """
    logger = logging.getLogger(name)
    logger.setLevel(log_level)
    
    # Evitar duplicar handlers si ya existe
    if logger.handlers:
        return logger
    
    # Crear directorio de logs si no existe
    log_dir = Path('logs')
    log_dir.mkdir(exist_ok=True)
    
    # Formato del log
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Handler para consola
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(log_level)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # Handler para archivo
    log_file = log_dir / f'{name}_{datetime.now().strftime("%Y%m%d")}.log'
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(log_level)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    return logger


# Logger por defecto para el proyecto
logger = setup_logger()
