"""
Script para verificar que el entorno esté correctamente configurado.
"""
import sys
from pathlib import Path
from src.logger_config import setup_logger

logger = setup_logger('check_setup')


def check_python_version():
    """Verifica la versión de Python."""
    version = sys.version_info
    logger.info(f"Python version: {version.major}.{version.minor}.{version.micro}")
    
    if version.major < 3 or (version.major == 3 and version.minor < 9):
        logger.error("❌ Python 3.9 o superior es requerido")
        return False
    
    logger.info("✅ Versión de Python correcta")
    return True


def check_dependencies():
    """Verifica que las dependencias estén instaladas."""
    required_packages = [
        'pandas',
        'sqlalchemy',
        'openpyxl',
        'xlwings',
        'pyodbc',
        'xlsxwriter'
    ]
    
    missing = []
    for package in required_packages:
        try:
            __import__(package)
            logger.info(f"✅ {package} instalado")
        except ImportError:
            logger.error(f"❌ {package} NO instalado")
            missing.append(package)
    
    if missing:
        logger.error(f"Paquetes faltantes: {', '.join(missing)}")
        logger.info("Ejecuta: pip install -r requirements.txt")
        return False
    
    return True


def check_database_config():
    """Verifica la configuración de base de datos."""
    config_file = Path('database.ini')
    example_file = Path('database.ini.example')
    
    if not example_file.exists():
        logger.error("❌ database.ini.example no encontrado")
        return False
    
    logger.info("✅ database.ini.example existe")
    
    if not config_file.exists():
        logger.warning("⚠️  database.ini no encontrado")
        logger.info("Copia database.ini.example a database.ini y configúralo")
        return False
    
    logger.info("✅ database.ini existe")
    
    # Verificar que no tenga valores por defecto
    try:
        from config import configSQLServer
        db_config = configSQLServer()
        
        if 'YOUR_' in db_config.get('server', '').upper():
            logger.warning("⚠️  database.ini parece tener valores por defecto")
            logger.info("Actualiza database.ini con tus credenciales reales")
            return False
        
        logger.info("✅ database.ini configurado")
        return True
        
    except Exception as e:
        logger.error(f"❌ Error leyendo database.ini: {e}")
        return False


def check_templates():
    """Verifica que los templates existan."""
    templates = ['Template_LAP.xlsx', 'Template_ACC.xlsx']
    
    all_exist = True
    for template in templates:
        template_path = Path(template)
        if template_path.exists():
            logger.info(f"✅ {template} encontrado")
        else:
            logger.warning(f"⚠️  {template} no encontrado")
            all_exist = False
    
    if not all_exist:
        logger.info("Algunos templates no están disponibles")
        logger.info("El script puede fallar si intenta usarlos")
    
    return all_exist


def check_directories():
    """Verifica y crea directorios necesarios."""
    directories = ['logs', 'tests']
    
    for directory in directories:
        dir_path = Path(directory)
        if not dir_path.exists():
            dir_path.mkdir(exist_ok=True)
            logger.info(f"✅ Directorio '{directory}' creado")
        else:
            logger.info(f"✅ Directorio '{directory}' existe")
    
    return True


def check_gitignore():
    """Verifica que .gitignore esté configurado correctamente."""
    gitignore = Path('.gitignore')
    
    if not gitignore.exists():
        logger.warning("⚠️  .gitignore no encontrado")
        return False
    
    content = gitignore.read_text()
    
    critical_entries = ['database.ini', '*.log', '__pycache__']
    missing_entries = []
    
    for entry in critical_entries:
        if entry not in content:
            missing_entries.append(entry)
    
    if missing_entries:
        logger.warning(f"⚠️  .gitignore no incluye: {', '.join(missing_entries)}")
        return False
    
    logger.info("✅ .gitignore configurado correctamente")
    return True


def main():
    """Ejecuta todas las verificaciones."""
    logger.info("=" * 60)
    logger.info("Verificando configuración del proyecto ExportaExcel")
    logger.info("=" * 60)
    
    checks = [
        ("Versión de Python", check_python_version),
        ("Dependencias", check_dependencies),
        ("Configuración de BD", check_database_config),
        ("Templates", check_templates),
        ("Directorios", check_directories),
        (".gitignore", check_gitignore),
    ]
    
    results = []
    for name, check_func in checks:
        logger.info(f"\n--- Verificando: {name} ---")
        try:
            result = check_func()
            results.append((name, result))
        except Exception as e:
            logger.error(f"Error en verificación de {name}: {e}")
            results.append((name, False))
    
    # Resumen
    logger.info("\n" + "=" * 60)
    logger.info("RESUMEN DE VERIFICACIONES")
    logger.info("=" * 60)
    
    all_passed = True
    for name, result in results:
        status = "✅ PASS" if result else "❌ FAIL"
        logger.info(f"{status} - {name}")
        if not result:
            all_passed = False
    
    logger.info("=" * 60)
    
    if all_passed:
        logger.info("🎉 ¡Todo está configurado correctamente!")
        logger.info("Puedes ejecutar: python ExportaExcel.py")
        return 0
    else:
        logger.warning("⚠️  Algunas verificaciones fallaron")
        logger.info("Revisa los mensajes anteriores para más detalles")
        return 1


if __name__ == "__main__":
    sys.exit(main())
