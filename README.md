# ExportaExcel

Sistema de exportación a Excel para generar reportes de reliquidación de contratos eléctricos.

## 📋 Descripción

Genera archivos Excel consolidando datos de EFACT, SIGGE y CEN para la reliquidación de contratos eléctricos.

Crea un archivo por cada combinación de: **Licitación + Empresa Generadora + Bloque + Distribuidora**

### Ejemplos de archivos generados
```
Lic2013-03_2 Caren BS1A 1-CEC.xlsx
Lic2013-03_2 San Juan BS2C 4-CGE_DISTRIBUCION.xlsx
Lic2013-03_2 Norvind BS4 28-SAESA.xlsx
```

## 🚀 Instalación

```bash
# Clonar repositorio
git clone https://github.com/tu-usuario/ExportaExcel.git
cd ExportaExcel

# Instalar dependencias
pip install -r requirements.txt
```

## ⚙️ Configuración

Crear archivo `database.ini` con las credenciales de SQL Server:

```ini
[sqlserver]
server=TU_SERVIDOR
database=TU_BASE_DE_DATOS
uid=TU_USUARIO
pwd=TU_CONTRASEÑA
```

**⚠️ Importante:** Este archivo NO se sube a Git por seguridad.

## 💻 Uso

```bash
python ExportaExcel.py
```

El script procesará todas las agrupaciones y generará los archivos Excel en el directorio actual.

## 📁 Estructura del Proyecto

```
ExportaExcel/
├── ExportaExcel.py              # ⭐ Script principal
├── BD.py                        # Funciones de BD
├── config.py                    # Configuración
├── GeneraExportacion.py         # Generación de reportes
├── database.ini                 # Credenciales (no en Git)
├── Template_LAP.xlsx            # Plantilla LAP
├── Template_ACC.xlsx            # Plantilla ACC
├── requirements.txt             # Dependencias
└── src/                         # 📦 Utilidades opcionales
    ├── constants.py             # Constantes reutilizables
    ├── db_utils.py              # Conexiones seguras
    ├── excel_utils.py           # Funciones de Excel
    ├── logger_config.py         # Sistema de logs
    ├── validators.py            # Validación de datos
    └── check_setup.py           # Verificación de setup
```

### Archivos principales (raíz)
- **ExportaExcel.py** - Script principal de ejecución
- **BD.py** - Funciones de consultas a base de datos
- **config.py** - Configuración de conexión
- **GeneraExportacion.py** - Generación de reportes

### Carpeta `src/` (utilidades opcionales)
Contiene código nuevo que agrega funcionalidades extras pero **no es obligatorio usar**

## 🔧 Requisitos

- Python 3.9 o superior
- SQL Server con acceso configurado
- Microsoft Excel (para xlwings)
- Dependencias en `requirements.txt`

## 📝 Notas

### ¿Para qué sirve la carpeta `src/`?

Los archivos en `src/` agregan funcionalidades extras:
- ✅ **constants.py** - Evita repetir valores en el código
- ✅ **db_utils.py** - Conexiones seguras (previene SQL injection)
- ✅ **excel_utils.py** - Funciones reutilizables de Excel
- ✅ **logger_config.py** - Logs para debugging
- ✅ **validators.py** - Validación de datos
- ✅ **check_setup.py** - Verifica instalación

**¿Los necesitas?** Solo si:
- Varias personas usan el código
- Necesitas debugging avanzado
- Te preocupa la seguridad

**Si solo tú lo usas internamente**, los archivos en la raíz (`ExportaExcel.py`, `BD.py`, `config.py`) son suficientes.

## 🤝 Contribuir

Las contribuciones son bienvenidas. Por favor:

1. Haz fork del proyecto
2. Crea una rama para tu feature (`git checkout -b feature/nueva-funcionalidad`)
3. Commit tus cambios (`git commit -m 'Agrega nueva funcionalidad'`)
4. Push a la rama (`git push origin feature/nueva-funcionalidad`)
5. Abre un Pull Request

Ver [CONTRIBUTING.md](CONTRIBUTING.md) para más detalles.

## 🔒 Seguridad

Para reportar vulnerabilidades de seguridad, ver [SECURITY.md](SECURITY.md).

## 📄 Licencia

Este proyecto está bajo la Licencia MIT. Ver [LICENSE](LICENSE) para más detalles.

## 📞 Soporte

- 🐛 Reportar bugs: [GitHub Issues](https://github.com/tu-usuario/ExportaExcel/issues)
- 📖 Documentación: Este README
- ✉️ Contacto: [tu-email@ejemplo.com]
