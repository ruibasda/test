# Este proceso genera un fichero csv leyendo todos los archivos xlsx que se encuentren en la ruta
# input 'str_pathCarpeta' y se filtran las lineas con valor 'NOK' en la columna 'Reactance diff < 1 Ω'
# El fichero resultante se guarda en la carpeta Informes que se creará en el directorio proporcionado
import datetime
import pandas as pd
import os

str_pathCarpeta = "C:/Users/ruibasda/OneDrive - REDEIA/Documentos/TYNDP2024_draft_models_imedances_check"
print(f'La carpeta actual es {str_pathCarpeta}.')

while True:  # input str_pathCarpeta
    str_inputRespuesta = input('¿Quieres modificar la ruta donde se encuentran los archivos? (y/n)').lower()
    if str_inputRespuesta == 'y':
        while True:
            str_pathCarpeta = input('Escriba la ruta de la carpeta donde se encuentran los archivos. (Utilice la opción de "copiar como ruta de acceso").').replace('\\', '/').replace('"', '')
            os.path.exists(os.path.dirname(str_pathCarpeta))
            if os.path.exists(os.path.dirname(str_pathCarpeta)):  # Avoid errors empty nodes
                print(f'Se realizará el procesamiento con los archivos ubicados en la ruta: {str_pathCarpeta}')
                break
            else:
                print('Se debe ingresar una ruta existente.')
        break
    elif str_inputRespuesta == 'n':
        print(f'Se mantiene la ruta: {str_pathCarpeta}.')
        break
    else:
        print('Respuesta no válida, por favor, responda "y" o "n".')

str_nameColumn = 'Reactance diff < 1 Ω'
str_nameValue = 'NOK'
str_fechaFormat = datetime.datetime.now().strftime('%d%m%Y_%H%M')
str_nameResultado = 'Informes/DraftModelIrredancesCheckNOK_{{fecha}}.csv'
str_pathResultado = os.path.join(str_pathCarpeta, str_nameResultado.replace('{{fecha}}', str_fechaFormat))
list_dataFrames = []
if not os.path.exists(os.path.dirname(str_pathResultado)):  # Se crean los directorios que no existan (Informes)
    os.makedirs(os.path.dirname(str_pathResultado))

for archivo in os.listdir(str_pathCarpeta):  # Se recorre todos los documentos que se encuentran en el directorio
    if archivo.endswith('.xlsx'):  # Se leen los ficheros .xlsx
        str_pathArchivo = os.path.join(str_pathCarpeta, archivo)
        df_dataExcel = pd.read_excel(str_pathArchivo)
        df_dataExcel['File Name'] = archivo  # Se añade un columna con el nombre del archivo
        list_dataFrames.append(df_dataExcel[df_dataExcel[str_nameColumn] == str_nameValue])
df_dataResultado = pd.concat(list_dataFrames, ignore_index=True)

df_dataResultado.to_csv(str_pathResultado, index=False)
