import ExportaExcel
import os
from time import time

Comienzo = time()
# Directorio = os.path.abspath(os.path.join(os.path.dirname( __file__ ), 'data')) #, '..'
Directorio = 'data'
Contenido = os.listdir(Directorio)
# print(Contenido)

ArchivosExcel = []
for fichero in Contenido:
    if os.path.isfile(os.path.join(Directorio, fichero)) and fichero.endswith('.xlsx'):
        ArchivosExcel.append(fichero)
# print(ArchivosExcel)

debug = False
j = 1
for Archivo in ArchivosExcel:
    print(f"{j}/{len(ArchivosExcel)}: {Archivo}")
    CargaExcelCNE.CargaExcelCNE(Archivo, debug)
    j += 1

# Archivo = 'EFACT CNE 2018-12.xlsx' #Energia EEDD CNE 2017-09
# CargaExcelCNE.CargaExcelCNE(Archivo, debug)

Transcurrido = time() - Comienzo
print("Tiempo transcurrido de ejecución: %0.10f seconds." % Transcurrido)
