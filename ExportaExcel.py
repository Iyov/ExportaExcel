import pandas
import sqlalchemy
import openpyxl
import xlwings as xw
import os
import pyodbc
from config import configSQLServer
import urllib

bd = configSQLServer()

Server = bd['server']
Database = bd['database']
Username = bd['uid']
Password = bd['pwd']

# Crea la conexión con la BD de SQL Server
params = urllib.parse.quote_plus("DRIVER={SQL Server};"
                                 "SERVER="+Server+";"
                                 "DATABASE="+Database+";"
                                 "UID="+Username+";"
                                 "PWD="+Password)

engine = sqlalchemy.create_engine("mssql+pyodbc:///?odbc_connect={}".format(params))
    
# Lee desde el SQL Server y crea los DataFrames con Pandas.

#Datos del Cliente
IdCliente = 1
Cliente = pandas.read_sql(f"""SELECT * FROM dbo.Cliente WHERE IdCliente = {IdCliente}""", engine)

NomCliente = Cliente.loc[[0]]["NomCliente"].item()
AbrevCliente = Cliente.loc[[0]]["AbrevCliente"].item()
RangoDesde = Cliente.loc[[0]]["RangoDesde"].item()
RangoHasta = Cliente.loc[[0]]["RangoHasta"].item()
print( NomCliente, AbrevCliente, RangoDesde, RangoHasta )

#Son 30, es un SP => son 27 con datos, filtrar Ids 12, 19, 20
ListaDeAgrup = pandas.read_sql('SELECT IdAgrupacion, NomAgrupacion FROM dbo.Agrupacion WHERE IdAgrupacion IN (30)', engine) #NOT IN (12, 19, 20)

#Son 8: Caren BS1 A, B, C, BS3; Norvind BS4; San Juan BS2A, 2C y 3
ListaGxBloques = pandas.read_sql("SELECT * FROM dbo.CNE_GxBloque WHERE IdCliente = '" + AbrevCliente + "'", engine) # AND GX = ''

NomTemplate = "Template_" + AbrevCliente + ".xlsx"
TemplateRM = pandas.read_excel(NomTemplate, sheet_name="README")
TemplateRE = pandas.read_excel(NomTemplate, sheet_name="ReliquidacionEFACT")
if( IdCliente == 1 ):
        TemplateRC = pandas.read_excel(NomTemplate, sheet_name="ReliquidacionCEN")
TemplateRT = pandas.read_excel(NomTemplate, sheet_name="ResumenRetiros")
# TemplateD = pandas.read_excel(NomTemplate, sheet_name="MAPA_DATA")
TemplateS = pandas.read_excel(NomTemplate, sheet_name="SIN_DATA")
TemplateI = pandas.read_excel(NomTemplate, sheet_name="INTERESES")

for i1, Agrupacion in ListaDeAgrup.iterrows():
    IdAgrupacion = int(Agrupacion["IdAgrupacion"])
    NomAgrupacion = Agrupacion["NomAgrupacion"]
    print( IdAgrupacion, "|", NomAgrupacion )

    for i2, GxBloque in ListaGxBloques.iterrows():
        Licitacion = GxBloque["Licitacion"]
        Empresa = GxBloque["GX"]
        Bloque = GxBloque["Bloque"]
        GX_CNE = GxBloque["GX_CNE"]
        GX_Sigge = GxBloque["GX_Sigge"]
        GX_CEN = GxBloque["GX_CEN"]
        # print( Empresa, "|", Bloque, "|", GX_CNE, "|", GX_Sigge )

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        NomExcel = f"""{Licitacion} {Empresa} {Bloque} {IdAgrupacion}-{NomAgrupacion}.xlsx"""
        writer = pandas.ExcelWriter(NomExcel, engine='xlsxwriter')

        

        EfactCNE = pandas.read_sql(f"""
            SELECT	c.IdContrato,
                    CONVERT(VARCHAR(100), c.Fecha, 105) AS MesDevengado,
                    c.Anho,
                    c.Mes,
                    c.Codigo,
                    c.DX,
                    c.NomEmpresaDistribuidora,
                    c.GX,
                    c.NomSuministrador,
                    c.CodigoContrato,
                    c.PuntoRetiro,
                    c.Contrato,
                    c.Energia_kWh AS Energia_kWh,
                    c.Potencia_kW AS Potencia_kW,
                    c.Sistema,
                    c.DataCreated,
                    b.IdBarra,
                    b.NomBarra,

                    Energia AS Energia,
                    TotalEnergia * f.FactorAjusteEnergia AS  TotalEnergiaAjustado,
                    Potencia AS  Potencia,
                    TotalPotencia * f.FactorAjustePotencia AS TotalPotenciaAjustado,

                    a.NomAgrupacion AS EmpresaDistribuidora,
                    PrecioEnergia AS PrecioEnergia,
                    PrecioEnergia * f.FactorAjusteEnergia AS PrecioEnergiaAjustado,
                    TotalEnergia AS TotalEnergia,
                    
                    PrecioPotencia AS  PrecioPotencia,
                    PrecioPotencia * f.FactorAjustePotencia AS PrecioPotenciaAjustado,
                    TotalPotencia AS  TotalPotencia,
                    c.Fecha
        FROM	dbo.CNE_Contrato c
                INNER JOIN dbo.CNE_Barra b ON c.IdContrato = b.IdContrato
                LEFT JOIN dbo.{AbrevCliente}_AgrupacionEFACT e ON c.DX+'|'+c.NomEmpresaDistribuidora+'|'+c.GX+'|'+c.NomSuministrador+'|'+c.CodigoContrato+'|'+c.Contrato = e.Tag
                LEFT JOIN dbo.Agrupacion a ON e.IdAgrupacion = a.IdAgrupacion
                LEFT JOIN dbo.CNE_FactorAjuste f ON c.Fecha = f.Fecha
        WHERE	(
                    c.GX LIKE '%{Empresa}%'
                    OR c.GX LIKE '%{GX_CNE}%'
                )
                AND c.Contrato LIKE '%{Bloque}%'
                AND a.NomAgrupacion = '{NomAgrupacion}'
        GROUP BY c.IdContrato,
                c.Fecha,
                c.Anho,
                c.Mes,
                c.Codigo,
                c.DX,
                c.NomEmpresaDistribuidora,
                c.GX,
                c.NomSuministrador,
                c.CodigoContrato,
                c.PuntoRetiro,
                c.Contrato,
                c.Energia_kWh,
                c.Potencia_kW,
                c.Sistema,
                c.DataCreated,
                b.IdBarra,
                b.NomBarra,
                Energia,
                PrecioEnergia,
                f.FactorAjusteEnergia,
                TotalEnergia,
                Potencia,
                PrecioPotencia,
                f.FactorAjustePotencia,
                TotalPotencia,
                a.NomAgrupacion

        UNION

        SELECT	DISTINCT
                NULL AS IdContrato,
                CONVERT(VARCHAR(100), FechaEfact, 105) AS MesDevengado,
                YEAR(FechaEfact) AS Anho,
                MONTH(FechaEfact) AS Mes,
                e.PuntoRetiro AS Codigo,
                e.Dx,
                e.NomEmpresaDistribuidora,
                e.Gx,
                e.Gx AS NomSuministrador,
                e.CodigoContrato,
                PuntoRetiro,
                NULL AS Contrato,
                Energia AS Energia_kWh,
                Potencia AS Potencia_kW,
                SistemaZonal,
                GETDATE() AS DataCreated,
                NULL AS IdBarra,
                BarraNAcional AS NomBarra,

                e.EnergiaPC AS  Energia,
                e.EnergiaRecPeso * f.FactorAjusteEnergia AS TotalEnergiaAjustado,
                e.PotenciaPC AS Potencia,
                PotenciaRecPeso * f.FactorAjustePotencia AS TotalPotenciaAjustado,

                a.NomAgrupacion,
                (e.Pe * e.Dolar/1000.0) AS  PrecioEnergia,
                (e.Pe * e.Dolar/1000.0) * f.FactorAjusteEnergia AS PrecioEnergiaAjustado,
                e.EnergiaRecPeso AS TotalEnergia,
                
                (Pp*Dolar) AS PrecioPotencia,
                (Pp*Dolar) * f.FactorAjustePotencia AS PrecioPotenciaAjustado,
                PotenciaRecPeso AS TotalPotencia,
                FechaEfact AS Fecha
        FROM	dbo.CNE_EfactPNP e
                LEFT JOIN dbo.{AbrevCliente}_AgrupacionEfactNew ae ON e.Dx+'|'+e.NomEmpresaDistribuidora+'|'+e.Gx+'|'+e.CodigoContrato = ae.Tag
                LEFT JOIN dbo.Agrupacion a ON ae.IdAgrupacion = a.IdAgrupacion
                LEFT JOIN dbo.CNE_FactorAjuste f ON e.FechaEfact = f.Fecha
        WHERE	e.tipoPNP = 'ITD'
                AND (
                    e.GX LIKE '%{Empresa}%'
                    OR e.GX LIKE '%{GX_CNE}%'
                )
                AND e.CodigoContrato LIKE '%{Bloque}%'
                AND a.NomAgrupacion = '{NomAgrupacion}'
        ORDER BY Fecha,
                CodigoContrato,
                PuntoRetiro,
                NomBarra
        """, engine)


        EfactCNE_BT = pandas.read_sql(f"""
            SELECT	c.IdContrato,
                    CONVERT(VARCHAR(100), c.Fecha, 105) AS MesDevengado,
                    c.Anho,
                    c.Mes,
                    c.Codigo,
                    c.DX,
                    c.NomEmpresaDistribuidora,
                    c.GX,
                    c.NomSuministrador,
                    c.CodigoContrato,
                    c.PuntoRetiro,
                    c.Contrato,
                    c.Energia_kWh AS Energia_kWh,
                    c.Potencia_kW AS Potencia_kW,
                    c.Sistema,
                    c.Fecha
            FROM	dbo.CNE_Contrato c
                    LEFT JOIN dbo.{AbrevCliente}_AgrupacionEFACT e ON c.DX+'|'+c.NomEmpresaDistribuidora+'|'+c.GX+'|'+c.NomSuministrador+'|'+c.CodigoContrato+'|'+c.Contrato = e.Tag
                    LEFT JOIN dbo.Agrupacion a ON e.IdAgrupacion = a.IdAgrupacion
            WHERE	(
                        c.GX LIKE '%{Empresa}%'
                        OR c.GX LIKE '%{GX_CNE}%'
                    )
                    AND c.Contrato LIKE '%{Bloque}%'
                    AND a.NomAgrupacion = '{NomAgrupacion}'
            GROUP BY c.IdContrato,
                    c.Fecha,
                    c.Anho,
                    c.Mes,
                    c.Codigo,
                    c.DX,
                    c.NomEmpresaDistribuidora,
                    c.GX,
                    c.NomSuministrador,
                    c.CodigoContrato,
                    c.PuntoRetiro,
                    c.Contrato,
                    c.Energia_kWh,
                    c.Potencia_kW,
                    c.Sistema

            UNION

            SELECT	DISTINCT
                    NULL AS IdContrato,
                    CONVERT(VARCHAR(100), FechaEfact, 105) AS MesDevengado,
                    YEAR(FechaEfact) AS Anho,
                    MONTH(FechaEfact) AS Mes,
                    e.PuntoRetiro AS Codigo,
                    e.Dx,
                    e.NomEmpresaDistribuidora,
                    e.Gx,
                    e.Gx AS NomSuministrador,
                    e.CodigoContrato,
                    PuntoRetiro,
                    NULL AS Contrato,
                    Energia AS Energia_kWh,
                    Potencia AS Potencia_kW,
                    SistemaZonal,
                    FechaEfact AS Fecha
            FROM	dbo.CNE_EfactPNP e
                    LEFT JOIN dbo.{AbrevCliente}_AgrupacionEfactNew ae ON e.Dx+'|'+e.NomEmpresaDistribuidora+'|'+e.Gx+'|'+e.CodigoContrato = ae.Tag
                    LEFT JOIN dbo.Agrupacion a ON ae.IdAgrupacion = a.IdAgrupacion
            WHERE	e.tipoPNP = 'ITD'
                    AND (
                        e.GX LIKE '%{Empresa}%'
                        OR e.GX LIKE '%{GX_CNE}%'
                    )
                    AND e.CodigoContrato LIKE '%{Bloque}%'
                    AND a.NomAgrupacion = '{NomAgrupacion}'
            ORDER BY Fecha,
                    CodigoContrato,
                    PuntoRetiro
        """, engine)

        
        QuerySigge = f"""
            SELECT	CONVERT(VARCHAR(100), f.Periodo, 105) AS Periodo,
                    CONVERT(VARCHAR(100), m.MesDevengado, 105) AS MesDevengado,
                    f.Nombre,
                    f.Concepto,
                    dc.Agrupacion,
                    f.Glosa,
                    f.FechaCarga,
                    f.TipoDocumento,
                    f.Folio,
                    f.Vendedor,
                    f.Comprador,
                    f.Bloque,
                    f.Empresa,
                    f.ClaveEEDD,
                    f.Licitacion,
                    a.NomAgrupacion,
					f.Num,
                    f.Barra,
                    CASE
                        WHEN dc.Agrupacion = 'Energía' THEN f.Cantidad
                        ELSE NULL
                    END Energia,
                    CASE
                        WHEN dc.Agrupacion = 'Energía' THEN f.Monto
                        ELSE NULL
                    END MontoEnergia,
                    CASE
                        WHEN dc.Agrupacion = 'Potencia' THEN f.Cantidad
                        ELSE NULL
                    END Potencia,
                    CASE
                        WHEN dc.Agrupacion = 'Potencia' THEN f.Monto
                        ELSE NULL
                    END MontoPotencia,
                    f.Precio
            FROM    dbo.{AbrevCliente}_SIGGE_Fact f
                    LEFT JOIN dbo.{AbrevCliente}_AgrupacionSIGGE asi ON f.Nombre = asi.CodigoContrato
                    LEFT JOIN dbo.Agrupacion a ON asi.IdAgrupacion = a.IdAgrupacion
                    LEFT JOIN dbo.{AbrevCliente}_MesDevengado m ON f.Glosa = m.Glosa
                    LEFT JOIN dbo.SIGGE_DiccionarioConceptos dc ON f.Concepto = dc.Concepto
            WHERE	( f.Empresa LIKE '%{Empresa}%' OR f.Empresa LIKE '%{GX_Sigge}%' )
                    AND f.Bloque = '{Bloque}'
                    AND a.NomAgrupacion = '{NomAgrupacion}'
                    --AND f.Folio IS NOT NULL		--ACL 2021-05-20
                    AND dc.Agrupacion IN (|Agrupador|)
            GROUP BY f.Periodo,
                    m.MesDevengado,
                    f.Nombre,
                    f.Num,
                    f.Barra,
                    f.Cantidad,
                    f.Precio,
                    f.Monto,
                    f.Concepto,
                    dc.Agrupacion,
                    f.Glosa,
                    f.FechaCarga,
                    f.TipoDocumento,
                    f.Folio,
                    f.Vendedor,
                    f.Comprador,
                    f.Bloque,
                    f.Empresa,
                    f.ClaveEEDD,
                    f.Licitacion,
                    a.NomAgrupacion
            ORDER BY f.Periodo,
                    f.Nombre,
                    f.Barra
        """
        Agrupador = "'Energía', 'Potencia'"
        # print(QuerySigge.replace("|Agrupador|", Agrupador))
        SiggeFact = pandas.read_sql(QuerySigge.replace("|Agrupador|", Agrupador), engine)

        Agrupador = "'Energía'"
        SiggeFactE = pandas.read_sql(QuerySigge.replace("|Agrupador|", Agrupador), engine)

        Agrupador = "'Potencia'"
        SiggeFactP = pandas.read_sql(QuerySigge.replace("|Agrupador|", Agrupador), engine)

        QueryCoordinador_Energia = f"""
            SELECT	ce.Num,
                    CONVERT(VARCHAR(100), ce.Fecha, 105) AS MesDevengado,
                    ce.Anho,
                    ce.Mes,
                    ce.Empresa,
                    ce.Propietario,
                    ce.Turno,
                    ce.Distribuidora,
                    ce.Clave,
                    ce.PuntoRetiro,
                    ce.Medida_kWh,
                    ce.PuntoRetiroBT,
                    ce.BarraEnAT,
                    ce.FactorRef_BT_AT,
                    ce.ValorAT,
                    ce.Bloque,
                    ce.Incremento,
                    NULL AS kWhAnho,
                    NULL AS PorcenEnergiaAnho,
                    NULL AS kWh_AnhoMes,
                    NULL AS PorcenBarraAnho,
                    ce.EnergiaBloqueAnhoMesBarra_kWh,
                    ce.Amplificacion24h,
                    ce.Denominador,
                    ce.Prorrata,
                    ce.EnergiaBloqueTurnoAnhoMesBarra_kWh,

                    pv.DecretoPNPVigente AS V_DecretoPNPVigente,
                    pv.DecretoPNCPAsociado AS V_DecretoPNCPAsociado,
                    pv.DecretoPNCPLicitacion AS V_DecretoPNCPLicitacion,
                    pv.Barra AS V_Barra,
                    pv.PrecioOferta_EnergiaPolpaico220 AS V_PrecioOferta_EnergiaPolpaico220,
                    pv.Indexacion AS V_Indexacion,
                    pv.PrecioEnergiaUSDMWh AS V_PrecioEnergiaIndexado_USDMWh,
                    pv.DolarPNPVigente AS V_DolarPNPVigente,
                    pv.FactorAjusteEnergia AS V_FactorAjusteEnergia,
                    pv.PrecioEnergiaAjustado_CLP_kWh_Polpaico220 AS V_PrecioEnergiaAjustado_CLP_kWh_Polpaico220,
                    pv.FPEnergia AS V_FPEnergia,
                    pv.PrecioEnergiaAjustado_CLP_kWh AS V_PrecioEnergiaAjustado_CLP_kWh,
                    pv.PrecioPotenciaUSD_kW_Mes AS V_PrecioPotenciaUSD_kW_Mes,
                    pv.PrecioPotenciaUSD_kW_Mes*pv.Indexacion AS V_PrecioPotenciaIndexadoUSD_kW_Mes,
                    pv.FactorAjustePotencia AS V_FactorAjustePotencia,
                    pv.PrecioPotenciaAjustado_CLP_kW_Mes_Polpaico220 AS V_PrecioPotenciaAjustado_CLP_kW_Mes_Polpaico220,
                    pv.FPPotencia AS V_FPPotencia,
                    pv.PrecioPotenciaAjustado_CLP_kW_Mes AS V_PrecioPotenciaAjustado_CLP_kW_Mes,
                    ce.EnergiaBloqueTurnoAnhoMesBarra_kWh * pv.PrecioEnergiaAjustado_CLP_kWh AS V_TotalEnergiaMesBarraGxDxBloque_CLP_kW_Mes,

                    pr.DecretoPNPReliquidacion AS R_DecretoPNPReliquidacion,
                    pr.DecretoPNCPAsociado AS R_DecretoPNCPAsociado,
                    pr.DecretoPNCPLicitacion AS R_DecretoPNCPLicitacion,
                    pr.Barra AS R_Barra,
                    pr.PrecioOfertaEnergiaPolpaico220 AS R_PrecioOfertaEnergiaPolpaico220,
                    pr.Indexacion AS R_Indexacion,
                    pr.PrecioEnergia_USD_MWh AS R_PrecioEnergiaIndexado_USDMWh,
                    pr.DolarPNPReliquidacion AS R_DolarPNPReliquidacion,
                    pr.FactorAjusteEnergia AS R_FactorAjusteEnergia,
                    pr.PrecioEnergiaAjustado_CLP_kWh_Polpaico220 AS R_PrecioEnergiaAjustado_CLP_kWh_Polpaico220,
                    pr.FPEnergia AS R_FPEnergia,
                    pr.PrecioEnergiaAjustado_CLP_kWh AS R_PrecioEnergiaAjustado_CLP_kWh,
                    pr.PrecioPotenciaUSD_kW_Mes AS R_PrecioPotenciaUSD_kW_Mes,
                    pr.FactorAjustePotencia AS R_FactorAjustePotencia,
                    pr.PrecioPotenciaAjustado_CLP_kW_Mes_Polpaico220 AS R_PrecioPotenciaAjustado_CLP_kW_Mes_Polpaico220,
                    pr.FPPotencia AS R_FPPotencia,
                    pr.PrecioPotenciaAjustado_CLP_kW_Mes AS R_PrecioPotenciaAjustado_CLP_kW_Mes,
                    ce.EnergiaBloqueTurnoAnhoMesBarra_kWh * pr.PrecioEnergiaAjustado_CLP_kWh AS R_TotalEnergiaMesBarraGxDxBloque_CLP_kW_Mes

            FROM    dbo.{AbrevCliente}_ConsolidadaTurnoBloqueEnergia ce
                    LEFT JOIN dbo.{AbrevCliente}_PrecioGxBloqueBarraVigente pv
                        ON	ce.Fecha = pv.Fecha
                            AND ce.Empresa = pv.Generador
                            AND ce.Bloque = pv.Bloque
                            AND ce.BarraEnAT = pv.BarraEnAT
                            AND pv.Barra LIKE '%220%'
                    LEFT JOIN dbo.{AbrevCliente}_PrecioGxBloqueBarraReliquidacion pr
                        ON	ce.Fecha = pr.Fecha
                            AND ce.Empresa = pr.Generador
                            AND ce.Bloque = pr.Bloque
                            AND ce.BarraEnAT = pr.BarraEnAT
                            AND pr.Barra LIKE '%220%'
            WHERE   (
                        ce.Empresa LIKE '%{Empresa}%'
                        OR ce.Empresa LIKE '%{GX_CEN}%'
                    )
                    AND ce.Bloque = '{Bloque}'
                    AND ce.Distribuidora = '{NomAgrupacion}'
            ORDER BY 2,5,6,10,13,16
        """
        Coordinador_Energia = pandas.read_sql(QueryCoordinador_Energia, engine)

        QueryCEN_EnergiaBT = f"""
            SELECT  DISTINCT
                    'R_G' AS Tipo,
                    CONVERT(VARCHAR(100), e.Fecha, 105) AS MesDevengado,
                    e.Anho,
                    e.Mes,
                    e.Empresa,
                    e.Propietario,
                    e.Turno,
                    e.Distribuidora,
                    e.Clave,
                    e.PuntoRetiro,
                    e.Medida_kWh*-1 AS ConsumoEnergiaBT,
                    e.PuntoRetiroBT,
                    e.Bloque,
                    e.Incremento
            FROM	dbo.LAP_ConsolidadaTurnoBloqueEnergia e
            WHERE	Distribuidora IS NOT NULL
                    AND (
                        e.Empresa LIKE '%{Empresa}%'
                        OR e.Empresa LIKE '%{GX_CEN}%'
                    )
                    AND e.Bloque = '{Bloque}'
                    AND e.Distribuidora = '{NomAgrupacion}'
            ORDER BY 2,5,6,9
        """
        CEN_EnergiaBT = pandas.read_sql(QueryCEN_EnergiaBT, engine)

        QueryCoordinador_Potencia = f"""
            SELECT	cp.Num,
                    CONVERT(VARCHAR(100), cp.Fecha, 105) AS MesDevengado,
                    cp.Anho,
                    cp.Mes,
                    cp.Empresa,
                    cp.Cliente,
                    cp.Distribuidora,
                    cp.BarraBT,
                    cp.Turno,
                    cp.PotenciaConsumo_kW,
                    cp.BloqueCompleto,
                    cp.BarraEnAT,
                    cp.FactorRefPotencia_BT_AT,
                    cp.Potencia_kW_AT,
                    cp.MaxPPA,
                    cp.Denominador,
                    cp.Prorrata,
                    cp.PotenciaBloqueTurnoAnhoMesBarra_kW,

                    pv.DecretoPNPVigente AS V_DecretoPNPVigente,
                    pv.DecretoPNCPAsociado AS V_DecretoPNCPAsociado,
                    pv.DecretoPNCPLicitacion AS V_DecretoPNCPLicitacion,
                    pv.Barra AS V_Barra,
                    pv.PrecioOferta_EnergiaPolpaico220 AS V_PrecioOferta_EnergiaPolpaico220,
                    pv.Indexacion AS V_Indexacion,
                    pv.PrecioEnergiaUSDMWh AS V_PrecioEnergiaIndexado_USDMWh,
                    pv.DolarPNPVigente AS V_DolarPNPVigente,
                    pv.FactorAjusteEnergia AS V_FactorAjusteEnergia,
                    pv.PrecioEnergiaAjustado_CLP_kWh_Polpaico220 AS V_PrecioEnergiaAjustado_CLP_kWh_Polpaico220,
                    pv.FPEnergia AS V_FPEnergia,
                    pv.PrecioEnergiaAjustado_CLP_kWh AS V_PrecioEnergiaAjustado_CLP_kWh,
                    pv.PrecioPotenciaUSD_kW_Mes AS V_PrecioPotenciaUSD_kW_Mes,
                    pv.PrecioPotenciaUSD_kW_Mes*pv.Indexacion AS V_PrecioPotenciaIndexadoUSD_kW_Mes,
                    pv.FactorAjustePotencia AS V_FactorAjustePotencia,
                    pv.PrecioPotenciaAjustado_CLP_kW_Mes_Polpaico220 AS V_PrecioPotenciaAjustado_CLP_kW_Mes_Polpaico220,
                    pv.FPPotencia AS V_FPPotencia,
                    pv.PrecioPotenciaAjustado_CLP_kW_Mes AS V_PrecioPotenciaAjustado_CLP_kW_Mes,
                    cp.PotenciaBloqueTurnoAnhoMesBarra_kW * pv.PrecioPotenciaAjustado_CLP_kW_Mes AS V_TotalPotenciaMesBarraGxDxBloque_CLP_kW_Mes,

                    pr.DecretoPNPReliquidacion AS R_DecretoPNPReliquidacion,
                    pr.DecretoPNCPAsociado AS R_DecretoPNCPAsociado,
                    pr.DecretoPNCPLicitacion AS R_DecretoPNCPLicitacion,
                    pr.Barra AS R_Barra,
                    pr.PrecioOfertaEnergiaPolpaico220 AS R_PrecioOfertaEnergiaPolpaico220,
                    pr.Indexacion AS R_Indexacion,
                    pr.PrecioEnergia_USD_MWh AS R_PrecioEnergiaIndexado_USDMWh,
                    pr.DolarPNPReliquidacion AS R_DolarPNPReliquidacion,
                    pr.FactorAjusteEnergia AS R_FactorAjusteEnergia,
                    pr.PrecioEnergiaAjustado_CLP_kWh_Polpaico220 AS R_PrecioEnergiaAjustado_CLP_kWh_Polpaico220,
                    pr.FPEnergia AS R_FPEnergia,
                    pr.PrecioEnergiaAjustado_CLP_kWh AS R_PrecioEnergiaAjustado_CLP_kWh,
                    pr.PrecioPotenciaUSD_kW_Mes AS R_PrecioPotenciaUSD_kW_Mes,
                    pr.FactorAjustePotencia AS R_FactorAjustePotencia,
                    pr.PrecioPotenciaAjustado_CLP_kW_Mes_Polpaico220 AS R_PrecioPotenciaAjustado_CLP_kW_Mes_Polpaico220,
                    pr.FPPotencia AS R_FPPotencia,
                    pr.PrecioPotenciaAjustado_CLP_kW_Mes AS R_PrecioPotenciaAjustado_CLP_kW_Mes,
                    cp.PotenciaBloqueTurnoAnhoMesBarra_kW * pr.PrecioPotenciaAjustado_CLP_kW_Mes AS R_TotalPotenciaMesBarraGxDxBloque_CLP_kW_Mes

            FROM	dbo.{AbrevCliente}_ConsolidadaTurnoBloquePotencia cp
                    LEFT JOIN dbo.{AbrevCliente}_PrecioGxBloqueBarraVigente pv
                        ON	cp.Fecha = pv.Fecha
                            AND cp.Empresa = pv.Generador
                            AND cp.BloqueCompleto = pv.Bloque
                            AND cp.BarraEnAT = pv.BarraEnAT
                            AND pv.Barra LIKE '%220%'
                    LEFT JOIN dbo.{AbrevCliente}_PrecioGxBloqueBarraReliquidacion pr
                        ON	cp.Fecha = pr.Fecha
                            AND cp.Empresa = pr.Generador
                            AND cp.BloqueCompleto = pr.Bloque
                            AND cp.BarraEnAT = pr.BarraEnAT
                            AND pr.Barra LIKE '%220%'
            WHERE	(
                        cp.Empresa LIKE '%{Empresa}%'
                        OR cp.Empresa LIKE '%{GX_CEN}%'
                    )
                    AND cp.BloqueCompleto = '{Bloque}'
                    AND cp.Distribuidora = '{NomAgrupacion}'
            ORDER BY 2,5,6,10,13,16
        """
        Coordinador_Potencia = pandas.read_sql(QueryCoordinador_Potencia, engine)

        if( IdCliente == 1 ):
            FactEstimada = pandas.read_sql(f"""
                SELECT	c.IdContrato,
                        CONVERT(VARCHAR(100), c.Fecha, 105) AS MesDevengado,
                        Anho,
                        Mes,
                        c.Codigo,
                        c.DX,
                        c.NomEmpresaDistribuidora,
                        c.GX,
                        c.NomSuministrador,
                        c.CodigoContrato,
                        c.PuntoRetiro,
                        c.Contrato,
                        Energia_kWh AS Energia_kWh,
                        Potencia_kW AS Potencia_kW,
                        Sistema,
                        DataCreated,
                        IdBarra,
                        NomBarra,

                        Energia AS Energia,
                        (Energia * d.PrecioEnergiaAjustado_CLP_kWh) AS TotalEnergiaAjustado,
                        Potencia AS Potencia,
                        (Potencia * d.PrecioPotenciaAjustado_CLP_kW_Mes) AS TotalPotenciaAjustado,
                        --d.DecretoPNPVigente, d.DecretoPNCPAsociado,

                        a.NomAgrupacion AS EmpresaDistribuidora,
                        d.PrecioEnergiaAjustado_CLP_kWh AS PrecioEnergia,
                        d.PrecioEnergiaAjustado_CLP_kWh AS PrecioEnergiaAjustado,
                        (Energia * d.PrecioEnergiaAjustado_CLP_kWh) AS TotalEnergia,
                        
                        d.PrecioPotenciaAjustado_CLP_kW_Mes AS PrecioPotencia,
                        (d.PrecioPotenciaAjustado_CLP_kW_Mes) AS PrecioPotenciaAjustado,
                        (Potencia * d.PrecioPotenciaAjustado_CLP_kW_Mes) AS TotalPotencia,
                        
                        c.Fecha
                FROM	dbo.CNE_Contrato c
                        INNER JOIN dbo.CNE_Barra b ON c.IdContrato = b.IdContrato
                        LEFT JOIN dbo.{AbrevCliente}_AgrupacionEFACT e ON c.DX+'|'+c.NomEmpresaDistribuidora+'|'+c.GX+'|'+c.NomSuministrador+'|'+c.CodigoContrato+'|'+c.Contrato = e.Tag
                        LEFT JOIN dbo.Agrupacion a ON e.IdAgrupacion = a.IdAgrupacion
                        LEFT JOIN dbo.LAP_PrecioGxBloqueBarraVigente d
                            ON	d.Generador = '{GX_CEN}'
                                AND d.Bloque = '{Bloque}'
                                AND b.NomBarra = d.Barra
                                AND c.Fecha = d.Fecha
                WHERE       (
                                    c.GX LIKE '%{Empresa}%'
                                    OR c.GX LIKE '%{GX_CNE}%'
                            )
                            AND c.Contrato LIKE '%{Bloque}%'
                            AND a.NomAgrupacion = '{NomAgrupacion}'
                GROUP BY c.IdContrato,
                        c.Fecha,
                        Anho,
                        Mes,
                        Codigo,
                        c.DX,
                        c.NomEmpresaDistribuidora,
                        c.GX,
                        c.NomSuministrador,
                        c.CodigoContrato,
                        c.PuntoRetiro,
                        c.Contrato,
                        Energia_kWh,
                        Potencia_kW,
                        Sistema,
                        DataCreated,
                        IdBarra,
                        NomBarra,
                        Energia,
                        d.PrecioEnergiaAjustado_CLP_kWh,
                        d.FPEnergia,
                        TotalEnergia,
                        --d.DecretoPNPVigente, d.DecretoPNCPAsociado,
                        Potencia,
                        d.PrecioPotenciaAjustado_CLP_kW_Mes,
                        d.FPPotencia,
                        TotalPotencia,
                        a.NomAgrupacion
                ORDER BY --MesDevengado,
                        c.Fecha,
                        c.CodigoContrato,
                        NomBarra
            """, engine)
            
            QueryReal = f"""
                SELECT	f.Empresa,
                        CONVERT(VARCHAR(100), f.MesDevengado, 105) AS MesDevengado,
                        f.IdAgrupacion,
                        a.NomAgrupacion,
                        f.Bloque,
                        f.NombreCliente,
                        f.FechaDeContabilizacion,
                        f.PrefijoFolio,
                        f.CuentaDeMayor,
                        f.CodigoDeArticulo,
                        f.NumDeFolio,
                        SUM(f.TotalCLP) AS MontoCLP,
                        SUM(f.TotalUSD) AS TotalUSD
                FROM	dbo.LAP_Facturacion2 f
                        LEFT JOIN dbo.Agrupacion a ON f.IdAgrupacion = a.IdAgrupacion
                        LEFT JOIN dbo.DolarMes d ON f.MesDevengado = d.Fecha
                WHERE	f.IdAgrupacion != -1
                        --AND MesDevengado IS NOT NULL
                        AND (
                                f.Empresa LIKE '%{Empresa}%'
                                OR f.Empresa LIKE '%{GX_CEN}%'
                        )
                        AND f.Bloque = '{Bloque}'
                        AND a.NomAgrupacion = '{NomAgrupacion}'
                        AND f.CodigoDeArticulo IN (|Agrupador|)
                GROUP BY f.Empresa,
                        f.MesDevengado,
                        f.IdAgrupacion,
                        a.NomAgrupacion,
                        f.Bloque,
                        f.NombreCliente,
                        f.FechaDeContabilizacion,
                        f.PrefijoFolio,
                        f.CuentaDeMayor,
                        f.CodigoDeArticulo,
                        f.NumDeFolio
                ORDER BY 1,2,3,4
            """

            Agrupador = "'VENTA PPA ENERGIA'"
            FactReal_E = pandas.read_sql(QueryReal.replace("|Agrupador|", Agrupador), engine)

            Agrupador = "'VENTA PPA POTENCIA'"
            FactReal_P = pandas.read_sql(QueryReal.replace("|Agrupador|", Agrupador), engine)

        # Escribe cada DataFrame en diferentes Hojas de Excel.
        TemplateRM.to_excel(writer, sheet_name='README', index=False)
        # if( IdCliente == 1 ):
        #     Mapa_Data_SQL.to_excel(writer, sheet_name='MAPA_DATA', index=False)
        TemplateRE.to_excel(writer, sheet_name='ReliquidacionEFACT', index=False)
        if( IdCliente == 1 ):
                TemplateRC.to_excel(writer, sheet_name='ReliquidacionCEN', index=False)
        TemplateRT.to_excel(writer, sheet_name='ResumenRetiros', index=False)
        
        SiggeFact.to_excel(writer, sheet_name='FACT_EMITIDA', index=False)
        SiggeFactE.to_excel(writer, sheet_name='SIGGE_E', index=False)
        SiggeFactP.to_excel(writer, sheet_name='SIGGE_P', index=False)
        if( IdCliente == 1 ):
            FactEstimada.to_excel(writer, sheet_name='FACT_EST_EFACT', index=False)
            FactReal_E.to_excel(writer, sheet_name='FactRealE', index=False)
            FactReal_P.to_excel(writer, sheet_name='FactRealP', index=False)
        
        Coordinador_Energia.to_excel(writer, sheet_name='CEN_Energia', index=False) #Actualizar CEN
        Coordinador_Potencia.to_excel(writer, sheet_name='CEN_Potencia', index=False)
        CEN_EnergiaBT.to_excel(writer, sheet_name='CEN_EnergiaBT', index=False)

        EfactCNE.to_excel(writer, sheet_name='EFACT_CNE', index=False)
        EfactCNE_BT.to_excel(writer, sheet_name='EfactCNE_BT', index=False)
        
        TemplateS.to_excel(writer, sheet_name='EFACT_EST_CEN', index=False) #Actualizar CEN
        TemplateS.to_excel(writer, sheet_name='SIN_DATA', index=False)
        TemplateI.to_excel(writer, sheet_name='INTERESES', index=False)

        # Close the Pandas Excel writer and output the Excel file.
        writer.save()

        
        DocumentoExcel = openpyxl.load_workbook(NomExcel, data_only=True)
        print( "Archivo", NomExcel, "cargado..." )
        
        
        #**********     HOJA "Reliquidacion EFACT" **********#
        ReliquidacionEFACT = DocumentoExcel["ReliquidacionEFACT"]
        for i, rowOfCellObjects in enumerate(ReliquidacionEFACT['A1':'AE1']):
            for n, cellObj in enumerate(rowOfCellObjects):
                cellObj.value = ''
        ReliquidacionEFACT['B2'] = Empresa
        ReliquidacionEFACT['B3'] = NomAgrupacion
        ReliquidacionEFACT['B5'] = Bloque

        #Fix Formato Fechas
        for i, rowOfCellObjects in enumerate(ReliquidacionEFACT['B9':'B65']):
            for n, cellObj in enumerate(rowOfCellObjects):
                cellObj.number_format = 'DD-MM-YYYY'
        
        # # LO QUE SE FACTURÓ
        # #Columna C: ORIGEN DATA
        # for i, rowOfCellObjects in enumerate(ReliquidacionEFACT['C9':'C65']):
        #     for n, cellObj in enumerate(rowOfCellObjects):
        #         Valor = f"""={cellObj.value}""" #{i+9}
        #         # Valor = f"""=IF(VLOOKUP(B{i+9},MAPA_DATA!D:G,2,0)=1,"FACT_EMITIDA",IF(VLOOKUP(B{i+9},MAPA_DATA!D:G,3,0)=1,"FACT_EST_EFACT",(IF(VLOOKUP(B{i+9},MAPA_DATA!D:G,4,0)=1,"FACT_EST_CEN","SIN_DATA"))))"""
        #         # print(Valor)
        #         cellObj.value = Valor
        
        for i, rowOfCellObjects in enumerate(ReliquidacionEFACT['D9':'J65']):
            for n, cellObj in enumerate(rowOfCellObjects):
                Valor = f"""={cellObj.value}"""
                cellObj.value = Valor
        

        # LO QUE SE DEBIÓ HABER FACTURADO
        # #Columna L: ORIGEN DATA
        # for i, rowOfCellObjects in enumerate(ReliquidacionEFACT['L9':'L65']):
        #     for n, cellObj in enumerate(rowOfCellObjects):
        #         # Valor = f"""={cellObj.value}""" #{i+9}
        #         Valor = f"""=IF(VLOOKUP(B{i+9},MAPA_DATA!D:I,5,0)=1,"EFACT_CNE",IF(VLOOKUP(B{i+9},MAPA_DATA!D,I,6,0)=1,"EFACT_EST_CNE","SIN_DATA"))"""
        #         cellObj.value = Valor
        
        for i, rowOfCellObjects in enumerate(ReliquidacionEFACT['M9':'S65']):
            for n, cellObj in enumerate(rowOfCellObjects):
                Valor = f"""={cellObj.value}"""
                cellObj.value = Valor
        
        #Columna V: Reliquidación mensual ($)
        for i, rowOfCellObjects in enumerate(ReliquidacionEFACT['V9':'V65']):
            for n, cellObj in enumerate(rowOfCellObjects):
                Valor = f"""={cellObj.value}"""
                cellObj.value = Valor
        #Columna AB: N° de días de intereses
        for i, rowOfCellObjects in enumerate(ReliquidacionEFACT['AB9':'AB65']):
            for n, cellObj in enumerate(rowOfCellObjects):
                Valor = f"""={cellObj.value}"""
                cellObj.value = Valor
        #Columna AC: Interés total según n° de días
        for i, rowOfCellObjects in enumerate(ReliquidacionEFACT['AC9':'AC65']):
            for n, cellObj in enumerate(rowOfCellObjects):
                Valor = f"""={cellObj.value}"""
                cellObj.value = Valor
        #Columna AD: Intereses ($)
        for i, rowOfCellObjects in enumerate(ReliquidacionEFACT['AD9':'AD65']):
            for n, cellObj in enumerate(rowOfCellObjects):
                Valor = f"""={cellObj.value}"""
                cellObj.value = Valor
        #Columna AE: Reliquidación Total ($)
        for i, rowOfCellObjects in enumerate(ReliquidacionEFACT['AE9':'AE66']):
            for n, cellObj in enumerate(rowOfCellObjects):
                Valor = f"""={cellObj.value}"""
                cellObj.value = Valor


        
        #**********     HOJA "Reliquidacion CEN" **********#
        if( IdCliente == 1 ):
                ReliquidacionCEN = DocumentoExcel["ReliquidacionCEN"]
                for i, rowOfCellObjects in enumerate(ReliquidacionCEN['A1':'AE1']):
                    for n, cellObj in enumerate(rowOfCellObjects):
                        cellObj.value = ''
                ReliquidacionCEN['B2'] = Empresa
                ReliquidacionCEN['B3'] = NomAgrupacion
                ReliquidacionCEN['B5'] = Bloque

                #Fix Formato Fechas
                for i, rowOfCellObjects in enumerate(ReliquidacionCEN['B9':'B65']):
                    for n, cellObj in enumerate(rowOfCellObjects):
                        cellObj.number_format = 'DD-MM-YYYY'

                for i, rowOfCellObjects in enumerate(ReliquidacionCEN['D9':'J65']):
                    for n, cellObj in enumerate(rowOfCellObjects):
                        Valor = f"""={cellObj.value}"""
                        cellObj.value = Valor

                for i, rowOfCellObjects in enumerate(ReliquidacionCEN['M9':'S65']):
                    for n, cellObj in enumerate(rowOfCellObjects):
                        Valor = f"""={cellObj.value}"""
                        cellObj.value = Valor

                #Columna V: Reliquidación mensual ($)
                for i, rowOfCellObjects in enumerate(ReliquidacionCEN['V9':'V65']):
                    for n, cellObj in enumerate(rowOfCellObjects):
                        Valor = f"""={cellObj.value}"""
                        cellObj.value = Valor
                #Columna AB: N° de días de intereses
                for i, rowOfCellObjects in enumerate(ReliquidacionCEN['AB9':'AB65']):
                    for n, cellObj in enumerate(rowOfCellObjects):
                        Valor = f"""={cellObj.value}"""
                        cellObj.value = Valor
                #Columna AC: Interés total según n° de días
                for i, rowOfCellObjects in enumerate(ReliquidacionCEN['AC9':'AC65']):
                    for n, cellObj in enumerate(rowOfCellObjects):
                        Valor = f"""={cellObj.value}"""
                        cellObj.value = Valor
                #Columna AD: Intereses ($)
                for i, rowOfCellObjects in enumerate(ReliquidacionCEN['AD9':'AD65']):
                    for n, cellObj in enumerate(rowOfCellObjects):
                        Valor = f"""={cellObj.value}"""
                        cellObj.value = Valor
                #Columna AE: Reliquidación Total ($)
                for i, rowOfCellObjects in enumerate(ReliquidacionCEN['AE9':'AE66']):
                    for n, cellObj in enumerate(rowOfCellObjects):
                        Valor = f"""={cellObj.value}"""
                        cellObj.value = Valor
        


        #**********     HOJA "ResumenRetiros" **********#
        ResumenRetiros = DocumentoExcel["ResumenRetiros"]
        for i, rowOfCellObjects in enumerate(ResumenRetiros['A1':'AE1']):
            for n, cellObj in enumerate(rowOfCellObjects):
                cellObj.value = ''
        ResumenRetiros['B2'] = Empresa
        ResumenRetiros['B3'] = NomAgrupacion
        ResumenRetiros['B5'] = Bloque

        #Fix Formato Fechas
        for i, rowOfCellObjects in enumerate(ResumenRetiros['B9':'B65']):
            for n, cellObj in enumerate(rowOfCellObjects):
                cellObj.number_format = 'DD-MM-YYYY'
        
        for i, rowOfCellObjects in enumerate(ResumenRetiros['D9':'H66']):
            for n, cellObj in enumerate(rowOfCellObjects):
                Valor = f"""={cellObj.value}"""
                cellObj.value = Valor


        DocumentoExcel.save(NomExcel)

        wbxl = xw.Book(NomExcel)
        app = xw.apps.active
        ReliquidacionEFACT = wbxl.sheets['ReliquidacionEFACT'].range('AE66').value
        
        ReliquidacionCEN = 0
        if( IdCliente == 1 ):
                ReliquidacionCEN = wbxl.sheets['ReliquidacionCEN'].range('AE66').value
        if( ReliquidacionCEN is None ):
            ReliquidacionCEN = 0
        print("ReliquidacionEFACT:", ReliquidacionEFACT, "| ReliquidacionCEN:", ReliquidacionCEN ) #, "| Diferencia:", (ReliquidacionEFACT-ReliquidacionCEN))
        
        #if( IdCliente == 2 ):
        Energia_SIGGE_kWh = wbxl.sheets['ResumenRetiros'].range('D66').value
        Energia_EFACT_AT_kWh = wbxl.sheets['ResumenRetiros'].range('E66').value
        Energia_EFACT_BT_kWh = wbxl.sheets['ResumenRetiros'].range('F66').value
        Energia_Coordinador_AT_kWh = wbxl.sheets['ResumenRetiros'].range('G66').value
        Energia_Coordinador_BT_kWh = wbxl.sheets['ResumenRetiros'].range('H66').value
        conn2 = pyodbc.connect('DRIVER={SQL Server};SERVER='+Server+';DATABASE='+Database+';UID='+Username+';PWD='+ Password)
        cursor2 = conn2.cursor()
        cursor2.execute(f"""
        DELETE
        FROM    dbo.{AbrevCliente}_ResumenRetiros
        WHERE   Licitacion = '{Licitacion}'
                AND Generadora = '{Empresa}'
                AND Distribuidora = '{NomAgrupacion}'
                AND Bloque = '{Bloque}'

        INSERT INTO dbo.{AbrevCliente}_ResumenRetiros ( Licitacion, Generadora, Distribuidora, Bloque, Energia_SIGGE_kWh, Energia_EFACT_AT_kWh, Energia_EFACT_BT_kWh, Energia_Coordinador_AT_kWh, Energia_Coordinador_BT_kWh )
        VALUES ( '{Licitacion}', '{Empresa}', '{NomAgrupacion}', '{Bloque}', {Energia_SIGGE_kWh}, {Energia_EFACT_AT_kWh}, {Energia_EFACT_BT_kWh}, {Energia_Coordinador_AT_kWh}, {Energia_Coordinador_BT_kWh} )
        """)
        conn2.commit()
        cursor2.close()
        conn2.close()
        
        wbxl.close()
        app.kill()

        if( ReliquidacionEFACT is not None ):
            conn3 = pyodbc.connect('DRIVER={SQL Server};SERVER='+Server+';DATABASE='+Database+';UID='+Username+';PWD='+ Password)
            cursor3 = conn3.cursor()
            cursor3.execute(f"""
                DELETE
                FROM    dbo.{AbrevCliente}_Reliquidacion
                WHERE   Licitacion = '{Licitacion}'
                        AND Generadora = '{Empresa}'
                        AND Distribuidora = '{NomAgrupacion}'
                        AND Bloque = '{Bloque}'

                INSERT INTO dbo.{AbrevCliente}_Reliquidacion ( Licitacion, Generadora, Distribuidora, Bloque, ReliquidacionEFACT, ReliquidacionCEN )
                VALUES ( '{Licitacion}', '{Empresa}', '{NomAgrupacion}', '{Bloque}', {ReliquidacionEFACT}, {ReliquidacionCEN} )
            """)
            conn3.commit()
            cursor3.close()
            conn3.close()

        
        # #Lee la hoja de reliquidación, obtiene si hay algun dato de EFECT en 0 y lo alerta
        # wbxl = xw.Book(NomExcel)
        # app = xw.apps.active
        # EfactValores = wbxl.sheets['ReliquidacionEFACT'].range('S65:S65').value
        # # print("EfactValores:", EfactValores)
        # i=1
        # for Valor in EfactValores:
        #     if( Valor == 0.0 ):
        #         i = i+1
        # if( i>1 ):
        #     print( Empresa, "|", Bloque, "|", GX_CNE, "| ", i, "celdas a revisar" )
        
        # wbxl.close()
        # app.kill()

