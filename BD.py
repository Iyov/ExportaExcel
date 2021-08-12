import pyodbc
from config import configSQLServer

bd = configSQLServer()

Server = bd['server']    #tcp:myserver.database.windows.net
Database = bd['database']
Username = bd['uid']
Password = bd['pwd']
NumErrores = 0

def GeneraDataExcelEfact(Agrupacion, GxBloque, Tag, debug):
    IdContrato = -1
    try:
        sql = f"""
            
            VALUES ('{Agrupacion}', {GxBloque}, {Tag})
        """
        if( debug ):
            print(sql)
            IdContrato = -2
        else:
            conn = pyodbc.connect('DRIVER={SQL Server};SERVER='+Server+';DATABASE='+Database+';UID='+Username+';PWD='+ Password)
            cursor = conn.cursor()
            cursor.execute(sql)
            cursor.execute("SELECT SCOPE_IDENTITY()")
            row = cursor.fetchone()
            IdContrato = row[0]
            conn.commit()
            cursor.close()
            conn.close()

        return IdContrato

    except pyodbc.Error as e:
        print(f"Error Contrato en fila={fila}: ('{Anho}{Mes}01', {Anho}, {Mes}, '{Codigo}', '{DX}', '{NomEmpresaDistribuidora}', '{GX}', '{NomSuministrador}', '{CodigoContrato}', '{PuntoRetiro}', '{Contrato}', {Energia_kWh}, {Potencia_kW}, '{Sistema}')")
        NumErrores += 1
        return IdContrato            

def InsertaBarra(fila,IdContrato, NomBarra,Energia,PrecioEnergia,Potencia,PrecioPotencia,debug):
    try:
        
        sql = f"""
            INSERT INTO CNE_Barra (IdContrato,NomBarra,Energia,PrecioEnergia,TotalEnergia,Potencia,PrecioPotencia,TotalPotencia) 
            VALUES ({IdContrato}, '{NomBarra}', {Energia}, {PrecioEnergia}, {Energia*PrecioEnergia}, {Potencia}, {PrecioPotencia}, {Potencia*PrecioPotencia})
        """
        if( debug ):
            print(sql)
            IdBarra = -2
        else:
            conn = pyodbc.connect('DRIVER={SQL Server};SERVER='+Server+';DATABASE='+Database+';UID='+Username+';PWD='+ Password)
            cursor = conn.cursor()
            cursor.execute(sql)
            # cursor.execute("SELECT SCOPE_IDENTITY()")
            # row = cursor.fetchone()
            # IdBarra = row[0]
            conn.commit()
            cursor.close()
            conn.close()

    except pyodbc.Error as e:
        print(f"Error Barra en fila={fila}: ({IdContrato}, '{NomBarra}', {Energia}, {PrecioEnergia}, {Potencia}, {PrecioPotencia})")
        NumErrores += 1

def SeleccionaDatos(Tabla, debug):
    try:
        sql = f"""
            SELECT * FROM {Tabla}
        """
        if( debug ):
            print(sql)
        else:
            cursor.execute(sql)
            for row in cursor:
                print(row)

    except pyodbc.Error as e:
        print("Error en la Conexion a la BD", e)

def TerminarConexion():
    return NumErrores
