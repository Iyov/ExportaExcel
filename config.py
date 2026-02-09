"""
Módulo de configuración para conexión a base de datos.
Soporta configuración desde archivo INI y variables de entorno.
"""
import os
from configparser import ConfigParser
from typing import Dict


def config(filename: str = 'database.ini', section: str = 'postgresql') -> Dict[str, str]:
    """
    Lee la configuración de PostgreSQL desde archivo INI.
    
    Args:
        filename: Nombre del archivo de configuración
        section: Sección del archivo INI a leer
        
    Returns:
        Diccionario con parámetros de conexión
        
    Raises:
        Exception: Si la sección no existe en el archivo
    """
    parser = ConfigParser()
    parser.read(filename)

    db = {}
    if parser.has_section(section):
        params = parser.items(section)
        for param in params:
            db[param[0]] = param[1]
    else:
        raise Exception(f'Section {section} not found in the {filename} file')

    return db


def configSQLServer(filename: str = 'database.ini', section: str = 'sqlserver') -> Dict[str, str]:
    """
    Lee la configuración de SQL Server desde archivo INI o variables de entorno.
    Las variables de entorno tienen prioridad sobre el archivo.
    
    Variables de entorno soportadas:
        - DB_SERVER: Servidor de base de datos
        - DB_DATABASE: Nombre de la base de datos
        - DB_UID: Usuario
        - DB_PWD: Contraseña
    
    Args:
        filename: Nombre del archivo de configuración
        section: Sección del archivo INI a leer
        
    Returns:
        Diccionario con parámetros de conexión
        
    Raises:
        Exception: Si la sección no existe en el archivo y no hay variables de entorno
    """
    # Intentar leer desde variables de entorno primero
    db = {}
    if os.getenv('DB_SERVER'):
        db['server'] = os.getenv('DB_SERVER')
        db['database'] = os.getenv('DB_DATABASE', '')
        db['uid'] = os.getenv('DB_UID', '')
        db['pwd'] = os.getenv('DB_PWD', '')
        return db
    
    # Si no hay variables de entorno, leer desde archivo
    parser = ConfigParser()
    parser.read(filename)

    if parser.has_section(section):
        params = parser.items(section)
        for param in params:
            db[param[0]] = param[1]
    else:
        raise Exception(f'Section {section} not found in the {filename} file')

    return db
