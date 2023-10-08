"""
This script works to get the best configuration of a double busbar
substation with multiple bays. The selection criterion is: the best
configuration will be such that minimizes the current through the 
coupler.

To use this script you must (directly in PowerFactory): 

  1. Fill all the variable in 'Basic Options'
      * sRutaTemps
      * User_FileName
      * sTini
      * sTfin
      * sMuestreo

  2. In the 'SEsDobleBarra' folder you should:
      a. Load the substations in each general set.
      b. Change the name of the general sets to reflect
        the name of your each selected substation.

_________________________Important note________________________________

The GitHub version of this script doens't use PowerFactory at all.

This code can be executed directly from VS-Code (or any other IDE).

If you want it to work with PF you should modify the Python script 
in such a way that you can look for the historical data of the 
selected substations (P and Q) directly from PowerFactory.
If you run this code from PowerFactory, it will ignore any choice
you have made in 'SEsDobleBarra' and it will look for the P and Q 
files provided in:
  > GitHub\DBB_conf\Example
_______________________________________________________________________

Authors:
    Jorge Mola & Delio Gomez
    XM S.A E.S.P
    Medellín, Colombia
    Oct-2023
"""
#------------------------------------------------------------------------------
#                                LIBRERÍAS
#------------------------------------------------------------------------------
##------- Librerías para trato de datos
import pandas as pd
import numpy as np

##------- Librerías de sistema:
from os import listdir

#------------------------------------------------------------------------------
#                                FUNCIONES
#------------------------------------------------------------------------------
class app():
    """
    This is only for the GitHub version. 
    The intention is for the code to run both from a PowerFactory execution
    or from a Python exectution. 
    The 'app' refers to the PowerFactory app. But since this is a version 
    intended to showcase the methodology this is done to avoid running 
    errors for the people who doesn't have PowerFactory license.
    """
    def __init__(self) -> None:
        pass
    def PrintInfo(self, Thing2Print):
        return print(Thing2Print)
    def PrintWarn(self, Thing2Print):
        return print(Thing2Print)
        
def Conn2PowerFactory():
    """
    Esta función establece la conexión con PowerFactory, resetea cálculos
    existentes en PF y limpia el outputwindow.
    """
    import powerfactory
    app = powerfactory.GetApplication()  
    project = app.GetActiveProject()  
    if not project: 
        app.PrintError('No project active !!!')
    app.ResetCalculation()   # start always from the same PowerFactory state
    app.ClearOutputWindow()   # clear Output Window (OPTIONAL)

    app.PrintInfo('...Script iniciado...')
    return app

def CreateNewNameForFile(sRutaCompleta_y_NombreArchivo):
    """
    Esta función genera un nuevo nombre para un archivo con base a otros nombres similares.
    ejm: si ya existe 'File1.xlsx', esta función genera 'File1(0).xlsx'
    Args:
        sRutaCompleta_y_NombreArchivo (string): ruta y nombre del archivo con nombre original
    output:
        FileName_out (string): El nuevo nombre del archivo que no generará conflicto al guardar
    """
    NameFile_w_xlsx = sRutaCompleta_y_NombreArchivo.split('\\')[-1]
    idx_Ruta = len(sRutaCompleta_y_NombreArchivo) - len(NameFile_w_xlsx)
    Ruta = sRutaCompleta_y_NombreArchivo[:idx_Ruta]
    NameFile_wo_xlsx = NameFile_w_xlsx.split('.')[0]
    Files_Con_MismoNombre = [f for f in listdir(Ruta) if f.startswith(NameFile_wo_xlsx)]
    for NewName_addendum in range(100):
        NewName = ''.join([NameFile_wo_xlsx,'(',str(NewName_addendum),').xlsx'])
        if NewName not in Files_Con_MismoNombre:
            FileName_out = ''.join([Ruta,NewName])
            break   
    return FileName_out             

def ConsultarDatos(P_or_Q) -> pd.DataFrame: 
    """
    This function has been modified to work 
    with the example in the GitHub repository.

    To adapt this tool to your specific interests
    you should modify this function to look for the 
    historical data about P and Q of your system.

    Args:
        P_or_Q (string)

    Returns:
        DataFrame: containing the historical data

    """
    DIR_P_FILE = r'Example\SE1_220kV_EightBays__P__.csv'
    DIR_Q_FILE = r'Example\SE1_220kV_EightBays__Q__.csv'

    if P_or_Q == 'P':
        dfResultPi = pd.read_csv(DIR_P_FILE, index_col=['Time'])
    elif P_or_Q == 'Q':
        dfResultPi = pd.read_csv(DIR_Q_FILE, index_col=['Time'])
    
    return dfResultPi

def Guardar_En_Excel(FileName_out, listDF, listNamesSheets):
    """
    Esta función exporta el resultado a un Excel

    Args:
        FileName_out (string): Ruta y nombre donde se guardará el archivo.
        listDF (list): lista de DataFrames que se exportarán.
        listNamesSheets (list): lista de los nombres para las hojas en Excel.

                                 
    """

    writer = pd.ExcelWriter(FileName_out, engine='openpyxl')
    for idx_Name,DF in enumerate(listDF):
        DF.to_excel(writer, sheet_name = listNamesSheets[idx_Name])    
    writer.close() 


#--------------------------------------------------------------------------------------------------
#----------------------------------------------------------       MAIN()      ---------------------
#--------------------------------------------------------------------------------------------------
RUN_GITHUB_EXAMPLE = True #<-- This doesn't exists in the original version. Only for GitHub.

if RUN_GITHUB_EXAMPLE:    
    #This conditional is merely for distribution purposes 
    #(GitHub version)
    app = app()
    sRutaTemps = r'd:\mis documentos\GitHub\DBB_conf'
    User_FileName = r'SE1__result'
    FolderSEsDoble = ['Example1'] #This is for the code to run (ignore this)    

else:
    #Conectarse a PowerFactory
    app = Conn2PowerFactory()

    #---------- Variables de usuario:
    script=app.GetCurrentScript()
    sRutaTemps = script.sRutaTemps                      # ¿En qué ruta se exportará la inforamación?
    FolderSEsDoble = script.SEsDobleBarra.GetContents() # SETs que representan cada subestación
    User_FileName = script.User_FileName                # ¿Qué nombre quiere para el archivo Excel de salida?
    # sTini = script.sTini                              # Fecha y hora inicial de la consulta
    # sTfin = script.sTfin                              # Fecha y hora final de la consulta
    # sMuestreo = script.sMuestreo.lower()              # Muestreo para la consulta    


#Umbrales importantes
UMBRAL_MIN_BAHIAS_EN_UNA_BARRA = 0.3    # Min bays connected to one busbar
UMBRAL_DIFERENCIA_SUMA_BARRA = 20       # If (ΣP or ΣQ < UMBRAL_DIFERENCIA_SUMA_BARRA) then: bad data

list_df_SEs=[]
list_Hoja_SEs=[]

for Set_SEs in FolderSEsDoble:
    app.PrintInfo("...Inicia Estimación S acople...")

    #Consultamos P y luego Q:
    df_P = ConsultarDatos('P')

    df_Q = ConsultarDatos('Q')                  

    ## En cada instante de tiempo se suman todas las P de los campos y todas la Q de los campos, 
    # conceptualmente esta suma debe ser cero, así que se tienen en cuenta solo aquellas muestras 
    # que en su suma sean menores a un umbral
    df_P = df_P[abs(df_P.sum(axis = 1)) < UMBRAL_DIFERENCIA_SUMA_BARRA].copy()
    df_Q = df_Q[abs(df_Q.sum(axis = 1)) < UMBRAL_DIFERENCIA_SUMA_BARRA].copy()   

    list_CtosName = df_P.columns.to_list()

    #Añadimos el sufijo
    list_CtosName_P = [''.join([cto,'[P]']) for cto in list_CtosName]
    list_CtosName_Q = [''.join([cto,'[Q]']) for cto in list_CtosName]
    #Renombramos las columnas
    df_P.columns = list_CtosName_P
    df_Q.columns = list_CtosName_Q


    #Se unen las dos P y Q ya procesadas en un solo dataframe, para garantizar que todas las muestras
    # tengan la misma estaman de tiempo
    df_PQ = pd.concat([df_P,df_Q], axis=1)
    
    #Botamos las filas que tengan NaN:
    df_PQ.dropna(axis = 0, how = 'any', inplace=True)

    #Se separa de nuevo en P y Q para el procesamiento individual, teniendo ya las series de tiempo 
    # homogéneas
    NumCols_Ctos = len(df_PQ.columns)
    NumCols_Ctos= NumCols_Ctos/2


    df_P = df_PQ[df_PQ.columns[:int(NumCols_Ctos)]].copy()
    df_Q = df_PQ[df_PQ.columns[int(NumCols_Ctos):]].copy()
    NumCols_Ctos = len(df_Q.columns)    

    #Se crean una matriz con todas la posibilidades en función del número de columnas, que corresponden
    #  a los campos, 1 corresponde a que el campo se conecta a una barra  y 0 que se conecta a la otra
    import itertools
    Combinations_01 = list(itertools.product([0, 1], repeat=NumCols_Ctos))
    df_Combinations_01 = pd.DataFrame(Combinations_01).T

    #En una distribución de barras normalmente se espera que haya un número similar de campos en cada barra, 
    # en ese sentido se descartan posibilidades que tengan pocos campos en una barra
    UMBRAL_BALANCE_BARRA = int(UMBRAL_MIN_BAHIAS_EN_UNA_BARRA*NumCols_Ctos) # Mínimo 30% ctos en una barra para posición 0 y para la 1
    bool_mask_2_remove = (df_Combinations_01.sum() >= UMBRAL_BALANCE_BARRA) & \
                            (df_Combinations_01.sum() <=NumCols_Ctos-UMBRAL_BALANCE_BARRA)
    options_2_choose_full = df_Combinations_01.columns*bool_mask_2_remove
    options_2_choose = [idx for idx in options_2_choose_full.to_list() if idx>0]
    df_Combinations_01 = df_Combinations_01[options_2_choose].copy()

    #Para realizar las operaciones matriciales es necesario que el DataFrame sea una matriz
    np_Com = df_Combinations_01.to_numpy()
    np_P = df_P.to_numpy()
    np_Q = df_Q.to_numpy()

    #Se realizan las operaciones matriciales de manera que se obtenga la P y Q estimada por el acople en cada
    #  muestra y para cada configuración, obteniendo matrices de tamaño 
    # (número de muestras x número de configuraciones)
    np_DiffOptions_P = np.dot(np_P,np_Com)
    np_DiffOptions_Q = np.dot(np_Q,np_Com)

    #Los resultados de las matrices se convierten de nuevo a DataFrame para volver a tener las etiquetas 
    # en cada columna que equivalen a las de las configuraciones, por esto se toman las mismas del 
    # DataFrame de combinaciones
    df_DiffOPtions_Q = pd.DataFrame(np_DiffOptions_Q, columns = df_Combinations_01.columns)
    df_DiffOPtions_P = pd.DataFrame(np_DiffOptions_P, columns = df_Combinations_01.columns)

    #Se calcula mediante una operación punto a punto la potencia aparente por el acople en cada muestra y 
    # para cada configuración
    df_DiffOPtions_S = (df_DiffOPtions_P**2 + df_DiffOPtions_Q**2).apply(np.sqrt)


    #Se obtienen los valores promedio y máximo de la corriente por el acople para cada configuración

    df_Sacople_Max=df_DiffOPtions_S.max()
    df_Sacople_Mean=df_DiffOPtions_S.mean()
    df_Indicadores=pd.concat([df_Sacople_Mean, df_Sacople_Max, ], axis=1)
    df_Indicadores.columns=['mean','max']

    df_Comb_fila=df_Combinations_01.T
    df_Comb_fila.columns=list_CtosName

    df_final=pd.concat([df_Comb_fila, df_Indicadores, ], axis=1)
    df_final=df_final.sort_values(by=['mean', 'max'],ascending=[True, True])

    list_df_SEs.append(df_final)
    try:
        list_Hoja_SEs.append(Set_SEs.loc_name)
    except AttributeError:
        list_Hoja_SEs.append(Set_SEs)

app.PrintInfo("...Creación archivo de Excel...")
try:
    FileName_out=sRutaTemps+"\\"+User_FileName+'.xlsx'
    app.PrintInfo("\n\n")
    app.PrintInfo(FileName_out)
    Guardar_En_Excel(FileName_out, list_df_SEs, list_Hoja_SEs)
except PermissionError:
    old_name = FileName_out
    FileName_out = CreateNewNameForFile(FileName_out)
    msg_warn = "El documento \'{}\' (o .csv) está abierto actualmente. Se guardará el resultado con el nombre: \'{}\'"
    app.PrintWarn(msg_warn.format(old_name, FileName_out))
    Guardar_En_Excel(FileName_out, list_df_SEs, list_Hoja_SEs)






