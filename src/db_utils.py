"""
Utilidades para manejo de base de datos con mejores prácticas de seguridad.
"""
import pyodbc
from typing import Optional, Any, Dict, List
from contextlib import contextmanager
from config import configSQLServer
from src.logger_config import setup_logger
from src.constants import SQL_SERVER_DRIVER

logger = setup_logger('db_utils')


class DatabaseConnection:
    """Clase para manejar conexiones a SQL Server de forma segura."""
    
    def __init__(self):
        """Inicializa la configuración de la base de datos."""
        try:
            bd = configSQLServer()
            self.server = bd['server']
            self.database = bd['database']
            self.username = bd['uid']
            self.password = bd['pwd']
            logger.info(f"Configuración de BD cargada: {self.server}/{self.database}")
        except Exception as e:
            logger.error(f"Error al cargar configuración de BD: {e}")
            raise
    
    def get_connection_string(self) -> str:
        """
        Retorna el string de conexión a SQL Server.
        
        Returns:
            String de conexión
        """
        return (
            f'DRIVER={{{SQL_SERVER_DRIVER}}};'
            f'SERVER={self.server};'
            f'DATABASE={self.database};'
            f'UID={self.username};'
            f'PWD={self.password}'
        )
    
    @contextmanager
    def get_connection(self):
        """
        Context manager para obtener una conexión a la base de datos.
        Asegura que la conexión se cierre correctamente.
        
        Yields:
            Conexión a la base de datos
            
        Example:
            with db.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM tabla")
        """
        conn = None
        try:
            conn = pyodbc.connect(self.get_connection_string())
            logger.debug("Conexión a BD establecida")
            yield conn
        except pyodbc.Error as e:
            logger.error(f"Error de conexión a BD: {e}")
            raise
        finally:
            if conn:
                conn.close()
                logger.debug("Conexión a BD cerrada")
    
    def execute_query(
        self, 
        query: str, 
        params: Optional[tuple] = None,
        fetch: bool = True
    ) -> Optional[List[tuple]]:
        """
        Ejecuta una query SQL de forma segura usando parámetros.
        
        Args:
            query: Query SQL con placeholders (?)
            params: Tupla de parámetros para la query
            fetch: Si True, retorna los resultados (SELECT), si False no retorna nada (INSERT/UPDATE)
            
        Returns:
            Lista de tuplas con resultados si fetch=True, None si fetch=False
            
        Example:
            results = db.execute_query(
                "SELECT * FROM tabla WHERE id = ?",
                (123,)
            )
        """
        with self.get_connection() as conn:
            cursor = conn.cursor()
            try:
                if params:
                    cursor.execute(query, params)
                else:
                    cursor.execute(query)
                
                if fetch:
                    results = cursor.fetchall()
                    logger.debug(f"Query ejecutada, {len(results)} filas retornadas")
                    return results
                else:
                    conn.commit()
                    logger.debug("Query ejecutada y commiteada")
                    return None
                    
            except pyodbc.Error as e:
                logger.error(f"Error ejecutando query: {e}")
                logger.error(f"Query: {query}")
                logger.error(f"Params: {params}")
                raise
            finally:
                cursor.close()
    
    def execute_insert(
        self, 
        query: str, 
        params: Optional[tuple] = None
    ) -> Optional[int]:
        """
        Ejecuta un INSERT y retorna el ID generado.
        
        Args:
            query: Query INSERT con placeholders (?)
            params: Tupla de parámetros para la query
            
        Returns:
            ID del registro insertado o None si hay error
            
        Example:
            id_contrato = db.execute_insert(
                "INSERT INTO tabla (col1, col2) VALUES (?, ?)",
                ('valor1', 'valor2')
            )
        """
        with self.get_connection() as conn:
            cursor = conn.cursor()
            try:
                if params:
                    cursor.execute(query, params)
                else:
                    cursor.execute(query)
                
                cursor.execute("SELECT SCOPE_IDENTITY()")
                row = cursor.fetchone()
                id_generado = int(row[0]) if row and row[0] else None
                
                conn.commit()
                logger.debug(f"INSERT ejecutado, ID generado: {id_generado}")
                return id_generado
                
            except pyodbc.Error as e:
                logger.error(f"Error ejecutando INSERT: {e}")
                logger.error(f"Query: {query}")
                logger.error(f"Params: {params}")
                conn.rollback()
                return None
            finally:
                cursor.close()


# Instancia global para uso en el proyecto
db_connection = DatabaseConnection()
