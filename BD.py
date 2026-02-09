"""
Módulo para operaciones de base de datos relacionadas con CNE.
Refactorizado para usar parámetros preparados y mejor manejo de errores.
"""
import pyodbc
from typing import Optional
from config import configSQLServer
from logger_config import setup_logger
from db_utils import DatabaseConnection

logger = setup_logger('BD')
db = DatabaseConnection()


def InsertaBarra(
    fila: int,
    IdContrato: int,
    NomBarra: str,
    Energia: float,
    PrecioEnergia: float,
    Potencia: float,
    PrecioPotencia: float,
    debug: bool = False
) -> bool:
    """
    Inserta un registro de barra en la base de datos.
    
    Args:
        fila: Número de fila (para logging)
        IdContrato: ID del contrato
        NomBarra: Nombre de la barra
        Energia: Energía en kWh
        PrecioEnergia: Precio de energía
        Potencia: Potencia en kW
        PrecioPotencia: Precio de potencia
        debug: Si True, solo imprime la query sin ejecutarla
        
    Returns:
        True si la inserción fue exitosa, False en caso contrario
    """
    try:
        TotalEnergia = Energia * PrecioEnergia
        TotalPotencia = Potencia * PrecioPotencia
        
        query = """
            INSERT INTO CNE_Barra (
                IdContrato, NomBarra, Energia, PrecioEnergia, TotalEnergia,
                Potencia, PrecioPotencia, TotalPotencia
            ) 
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """
        
        params = (
            IdContrato, NomBarra, Energia, PrecioEnergia, TotalEnergia,
            Potencia, PrecioPotencia, TotalPotencia
        )
        
        if debug:
            logger.debug(f"Query: {query}")
            logger.debug(f"Params: {params}")
            return True
        
        db.execute_query(query, params, fetch=False)
        logger.info(f"Barra insertada exitosamente en fila {fila}")
        return True
        
    except pyodbc.Error as e:
        logger.error(
            f"Error insertando Barra en fila={fila}: "
            f"IdContrato={IdContrato}, NomBarra={NomBarra}, "
            f"Energia={Energia}, PrecioEnergia={PrecioEnergia}, "
            f"Potencia={Potencia}, PrecioPotencia={PrecioPotencia}"
        )
        logger.error(f"Detalle del error: {e}")
        return False


def SeleccionaDatos(Tabla: str, debug: bool = False) -> Optional[list]:
    """
    Selecciona todos los datos de una tabla.
    
    Args:
        Tabla: Nombre de la tabla
        debug: Si True, solo imprime la query
        
    Returns:
        Lista de tuplas con los resultados o None si hay error
    """
    try:
        # Nota: Validar nombre de tabla para evitar SQL injection
        # En producción, usar una whitelist de tablas permitidas
        if not Tabla.replace('_', '').isalnum():
            logger.error(f"Nombre de tabla inválido: {Tabla}")
            return None
        
        query = f"SELECT * FROM {Tabla}"
        
        if debug:
            logger.debug(f"Query: {query}")
            return None
        
        results = db.execute_query(query)
        
        if results:
            for row in results:
                logger.info(row)
        
        return results
        
    except pyodbc.Error as e:
        logger.error(f"Error en la Conexión a la BD: {e}")
        return None


def InsertaContrato(
    Fecha: str,
    Anho: int,
    Mes: int,
    Codigo: str,
    DX: str,
    NomEmpresaDistribuidora: str,
    GX: str,
    NomSuministrador: str,
    CodigoContrato: str,
    PuntoRetiro: str,
    Contrato: str,
    Energia_kWh: float,
    Potencia_kW: float,
    Sistema: str,
    debug: bool = False
) -> Optional[int]:
    """
    Inserta un contrato en la base de datos.
    
    Args:
        Fecha: Fecha del contrato (formato YYYYMMDD)
        Anho: Año
        Mes: Mes
        Codigo: Código
        DX: Distribuidora
        NomEmpresaDistribuidora: Nombre de empresa distribuidora
        GX: Generadora
        NomSuministrador: Nombre del suministrador
        CodigoContrato: Código del contrato
        PuntoRetiro: Punto de retiro
        Contrato: Contrato
        Energia_kWh: Energía en kWh
        Potencia_kW: Potencia en kW
        Sistema: Sistema
        debug: Si True, solo imprime la query
        
    Returns:
        ID del contrato insertado o None si hay error
    """
    try:
        query = """
            INSERT INTO CNE_Contrato (
                Fecha, Anho, Mes, Codigo, DX, NomEmpresaDistribuidora,
                GX, NomSuministrador, CodigoContrato, PuntoRetiro, Contrato,
                Energia_kWh, Potencia_kW, Sistema
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
        
        params = (
            Fecha, Anho, Mes, Codigo, DX, NomEmpresaDistribuidora,
            GX, NomSuministrador, CodigoContrato, PuntoRetiro, Contrato,
            Energia_kWh, Potencia_kW, Sistema
        )
        
        if debug:
            logger.debug(f"Query: {query}")
            logger.debug(f"Params: {params}")
            return -2
        
        id_contrato = db.execute_insert(query, params)
        
        if id_contrato:
            logger.info(f"Contrato insertado exitosamente, ID: {id_contrato}")
        
        return id_contrato
        
    except pyodbc.Error as e:
        logger.error(f"Error insertando Contrato: {e}")
        logger.error(f"Datos: Fecha={Fecha}, GX={GX}, DX={DX}, Contrato={Contrato}")
        return None
