"""
Tests para utilidades de base de datos.
"""
import pytest
from unittest.mock import Mock, patch, MagicMock
from db_utils import DatabaseConnection


@pytest.fixture
def mock_config():
    """Fixture para configuración mock."""
    return {
        'server': 'test_server',
        'database': 'test_db',
        'uid': 'test_user',
        'pwd': 'test_pass'
    }


@patch('db_utils.configSQLServer')
def test_database_connection_init(mock_config_func, mock_config):
    """Test de inicialización de DatabaseConnection."""
    mock_config_func.return_value = mock_config
    
    db = DatabaseConnection()
    
    assert db.server == 'test_server'
    assert db.database == 'test_db'
    assert db.username == 'test_user'
    assert db.password == 'test_pass'


@patch('db_utils.configSQLServer')
def test_get_connection_string(mock_config_func, mock_config):
    """Test de generación de connection string."""
    mock_config_func.return_value = mock_config
    
    db = DatabaseConnection()
    conn_str = db.get_connection_string()
    
    assert 'test_server' in conn_str
    assert 'test_db' in conn_str
    assert 'test_user' in conn_str
    assert 'test_pass' in conn_str


@patch('db_utils.configSQLServer')
@patch('db_utils.pyodbc.connect')
def test_get_connection_context_manager(mock_connect, mock_config_func, mock_config):
    """Test del context manager de conexión."""
    mock_config_func.return_value = mock_config
    mock_conn = MagicMock()
    mock_connect.return_value = mock_conn
    
    db = DatabaseConnection()
    
    with db.get_connection() as conn:
        assert conn == mock_conn
    
    # Verificar que se cerró la conexión
    mock_conn.close.assert_called_once()
