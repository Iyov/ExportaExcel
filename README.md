# ExportaExcel
Proyecto de Exportación a Excel para generar el detalle de la Reliquidación

## Librerías a instalar
```
pip install pandas
pip install sqlalchemy
python -m pip install --upgrade 'sqlalchemy<2.0'
pip install openpyxl
pip install xlwings
pip install pyodbc
pip install xlsxwriter
```

## Descripción
Se genera un archivo Excel por cada Licitación, Empresa Generadora, Bloque Suministro, Empresa Distribuidora

### Ejemplos
```
Lic2013-03_2 EmpresaGx1 BS1A 1-EmpresaDx.xlsx
Lic2013-03_2 EmpresaGx2 BS2C 7-EmpresaDx.xlsx
Lic2013-03_2 EmpresaGx3 BS3 12-EmpresaDx.xlsx
Lic2013-03_2 EmpresaGx4 BS4 26-EmpresaDx.xlsx
```

### Configuración de Conexión a Base de Datos
Se debe generar un archivo ```database.ini``` con la Conexión a Base de Datos, en el contenido se debe escribir lo siguiente:

```
[sqlserver]
server=SERVIDOR
database=BASE_DE_DATOS
uid=USER_NAME_SQL
pwd=PASSWORD_SQL
```

