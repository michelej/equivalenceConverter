import pandas as pd # Panda Excel
import numpy as np
import sys 
import os  
import math # funciones matematicas
import json  # JSON libreria
import simplelogging # Libreria de logs
import datetime  # Fechas
import ntpath   # Rutas
import time
import xlwings as xw
import re
import traceback

class NullValue(Exception):   
   pass

class EmptyRow(Exception):      
   pass

class EndOfData(Exception):   
   pass

class CriticalError(Exception):
   pass

class InvalidFormat(Exception):
   pass

#config_file = "fichero_equivalencias.xlsx"  # fichero de equivalencias

log = None
logger_console = False  # log to console

pd.set_option("display.precision", 16)

def main():
    IS_SEDOL = False    

    incidence_file = None
    output_file = None
    input_file = None
    output_dir = None

    process_type = None
    process_id = None
    fields = []
    formulas = []
    field_types = []

    configExcel = None
    fields = None
    validation = None
    files = None

    error_format_match = False

    if len(sys.argv) < 5:
        print("ERROR : faltan argumentos..  <Tipo Proceso> <ID Proceso> <Fichero> <Directorio Destino>")
    else:
        start_time = time.time()
        process_type = sys.argv[1]  # Pagina en fichero
        process_id = sys.argv[2]  # ID del tipo
        input_file = sys.argv[3] # Fichero a procesar
        output_dir = sys.argv[4] # Direccion donde dejar el fichero
        config_file = sys.argv[5] # Fichero de equivalencias
                               
        if not os.path.exists(output_dir):
            print("ERROR : No se encuentra la carpeta de destino: " + output_dir)
            return                 
        if not os.path.exists(input_file):
            print("ERROR : No se encuentra el fichero a procesar: " + input_file)
            return

        now = datetime.datetime.now()
        input_file_name = os.path.splitext(ntpath.basename(input_file))
        log_name = "log_"+input_file_name[0]+"_"+now.strftime("%d%m%Y%H%M%S")+".txt"
        global log
        log = simplelogging.get_logger(file_name=output_dir+log_name,console=logger_console)    

        IS_SEDOL = True if not pd.isna(re.search("SEDOL", str(process_id))) else False        
        
        output_file = "output_"+process_type+".xlsx" if not IS_SEDOL else "output_SEDOL_"+process_type+".xlsx"
        incidence_file = "incidence_"+process_type+".xlsx" if not IS_SEDOL else "incidence_SEDOL_"+process_type+".xlsx"
        
        log.info("Iniciando Proceso")
        log.info("PID: "+ str(os.getpid()))
        log.info(process_type + " - ( " + process_id + " )")
                
        try:
            configExcel = pd.read_excel(config_file, process_type)                        
            configExcel.columns = configExcel.columns.astype(str) 

            if set(['Campos', process_id]).issubset(configExcel.columns):
                campos = configExcel["Campos"]                
                tipo = configExcel["Tipo"]                                               
                fields=campos[1:] 
                fields = fields.dropna(how='all')                                 
                fields = fields.apply(lambda x: x.strip())                               
                
                allFormulas = configExcel[process_id]
                formulas = allFormulas[1:len(fields)+1]                
                field_types=tipo[1:len(fields)+1]                                
                field_types = field_types.apply(lambda x: x.lower() if not pd.isna(x) else x)
                                
                filter_func = allFormulas[0]

                validate_field_types(field_types)                                
                validate_formulas(formulas)
            else:
                raise Exception("No se han encontrado las columnas necesarias : (Campos ," + process_id + ")")              
        except OSError as e:
            log.error("Error: " + str(e))
            print("ERROR : Fichero de equivalencias no encontrado")
            return 
        except Exception as e:
            log.error("Error: " + str(e))
            print("ERROR : Fichero de equivalencias ERRONEO")
            return 
        
        log.info("Procesando fichero: " + input_file)
        
        file_extension = os.path.splitext(input_file)
        dataExcel = None
        try:
            if(file_extension[1] == ".csv"):
                dataExcel = pd.read_csv(input_file,  header=None)
            else:
                app = xw.App(visible=False)
                wb = xw.Book(input_file,ignore_read_only_recommended=True)
                active_sheet_name = wb.sheets.active.name                
                wb.close()
                app.quit()
                dataExcel = pd.read_excel(input_file, sheet_name=active_sheet_name, header=None)
                dataExcel=dataExcel.replace('\n', '',regex=True).replace('\r','',regex=True)
        except Exception as e: 
            print("ERROR : No se pudo abrir el fichero a procesar")
            return        

        firstLineExists = gather_data_from_excel(dataExcel,filter_func,fields,formulas,field_types,IS_SEDOL,True)       
        outputData=[]
        incidenceData=[]
        if((len(firstLineExists["outputData"])>0 or len(firstLineExists["incidenceData"])>0) and firstLineExists["error_format_match"] == False):
            if(not pd.isna(filter_func)):
                filter_data_excel(filter_func, dataExcel)
                log.info("Filtro aplicado.")
            dataFromExcel = gather_data_from_excel(dataExcel,filter_func,fields,formulas,field_types,IS_SEDOL,False)
            
            outputData = dataFromExcel["outputData"]
            incidenceData = dataFromExcel["incidenceData"]
            error_format_match = False
        else:
            error_format_match=True            

        try:
            flag_ok = False
            if(len(outputData)>0 or len(incidenceData)>0):                     
                resultExcel = pd.DataFrame(outputData,columns=fields)                            
                resultExcel.drop_duplicates(keep='first', inplace=True)                                          
                
                #print(resultExcel)
                number_fields = {}    
                for i in range(len(field_types)):
                    if(field_types[i+1] == 'number'):
                        number_fields[fields[i+1]]=2
                resultExcel=resultExcel.round(number_fields)
                
                ratio_fields = {}    
                for i in range(len(field_types)):
                    if(field_types[i+1] == 'ratio'):
                        ratio_fields[fields[i+1]]=12        
                         
                resultExcel=resultExcel.round(ratio_fields)                                

                #print(resultExcel.dtypes)
                #print(resultExcel)   
                historyExcel = load_excel(output_dir+output_file,fields)                
                #print(historyExcel.dtypes)
                #print(historyExcel)   
                finalExcel = None
                equalRows = pd.DataFrame(columns=['ISIN'])
                if(not historyExcel.empty and not IS_SEDOL):
                    historyExcel = historyExcel.dropna(how='all')                      
                    match_dataframes_types(resultExcel,historyExcel)                    
                    resultExcel = dataframe_difference(resultExcel, historyExcel, "left_only")
                    #print(resultExcel)
                    finalExcel=pd.concat([historyExcel,resultExcel])                                                                                               
                else:
                    finalExcel=resultExcel                    

                #print(resultExcel)
                finalExcel.drop_duplicates(keep='first', inplace=True)
                finalExcel.reset_index(drop=True, inplace=True)       
                #print(finalExcel)        
                
                columnId = fields[1] if not IS_SEDOL else fields[2]
                equalRows = finalExcel[finalExcel.duplicated(columnId, keep=False)]            
                equalRows.reset_index(drop=True, inplace=True)                      
                #print(equalRows)  

                columnId = fields[1] if not IS_SEDOL else fields[2]
                finalExcel.drop_duplicates(columnId, inplace=True, keep=False)                        
                finalExcel.reset_index(drop=True, inplace=True)            
                #print(finalExcel)
                
                #print(resultExcel)
                newsExcel = dataframe_difference(resultExcel, finalExcel, "both")
                #print(newsExcel)
                                                            
                flag_data = False
                if(not finalExcel.empty):
                    if(save_excel(output_dir+output_file, finalExcel)):
                        log.info("Guardando fichero: "+output_dir+output_file)
                        log.info("Proceso finalizado...")                        

                        res_ids = newsExcel.iloc[:, 0] if not IS_SEDOL else newsExcel.iloc[:, 1]                        
                        if(len(res_ids)>0):
                            print("OK")
                            flag_ok = True
                            for ids in res_ids:
                                print(str(ids) + ",OK,"+output_file)
                            flag_data = True
                    else:
                        print("ERROR : Fallo en la escritura del fichero de salida.")
                        return
                
                if(len(incidenceData)>0 or not equalRows.empty):
                    incidenceExcel = pd.DataFrame(incidenceData,columns=fields) 
                    #print(equalRows)
                    incidenceExcel=pd.concat([incidenceExcel,equalRows])                
                    #print(incidenceExcel)

                    historyIncidenceExcel = load_excel(output_dir+incidence_file,fields)
                    finalIncidenceExcel = None
                    if(not historyIncidenceExcel.empty and not IS_SEDOL):                        
                        incidenceExcel=incidenceExcel.where(~incidenceExcel.notnull(), incidenceExcel.astype('str'))
                        historyIncidenceExcel=historyIncidenceExcel.where(~historyIncidenceExcel.notnull(), historyIncidenceExcel.astype('str'))
                        #match_dataframes_types(incidenceExcel,historyIncidenceExcel)
                        finalIncidenceExcel = historyIncidenceExcel.append(incidenceExcel)                                                    
                    else:
                        finalIncidenceExcel = incidenceExcel                                                

                    finalIncidenceExcel.drop_duplicates(inplace=True) # Los identicos se ignoran
                    finalIncidenceExcel.reset_index(drop=True,inplace=True)                                                                    
                    #print(finalIncidenceExcel)

                    newsIncidenceExcel = finalIncidenceExcel.append(historyIncidenceExcel)                    
                    newsIncidenceExcel.drop_duplicates(keep=False,inplace=True)                    
                    #print(newsIncidenceExcel)

                    if(save_excel(output_dir+incidence_file,finalIncidenceExcel)):                        
                        log.info("Guardando fichero: "+output_dir+incidence_file)                                                                            
                        err_isins = newsIncidenceExcel.iloc[:, 0]   
                        if(len(err_isins)>0):
                            if(flag_ok ==False):
                                print("OK")                        
                            for isin in err_isins:
                                if(not pd.isna(isin)):                            
                                    print(isin +",ERROR,"+incidence_file)     
                            flag_data = True
                    else:
                        print("ERROR : Fallo en la escritura del fichero de incidencias.")                    
                
                if(flag_data == False):
                    print("NO DATA: No se encontraron datos para procesar")
            else:
                if(error_format_match == True):
                    print("ERROR : El formato para este excel no es correcto: " + process_type + " - " + process_id)
                elif(len(outputData) == 0 and len(incidenceData) == 0):
                    print("NO DATA: No se encontraron datos para procesar")
                else:    
                    print("ERROR") 
        except (InvalidFormat) as e:            
            print("ERROR : " + str(e))
        except Exception as e:                
            log.error(e)
            print("ERROR : Ha ocurrido un error inesperado.")

    elapsed_time = time.time() - start_time
    log.info("Tiempo de ejecución: " + str(elapsed_time) +" segundos")

################################################
##  Obtener los datos del Excel 
##  
################################################
def gather_data_from_excel(dataExcel,filter_func,fields,formulas,field_types,IS_SEDOL,VALIDATE):
    finished = False
    index = 0
    outputData = []
    incidenceData = []
    error_format_match = False
    try:       
        while not finished:           
            i = 1
            rowResult = []
            errorsFound = False
            forcedErrorsFound = False
            row_id = None
            for j in range(len(fields)):
                try:
                    dataResult = float('NaN')
                    if(not pd.isna(formulas[i])):
                        convertResult = convert_function(formulas[i], dataExcel, index, True if j == 0 else False)  # Datos
                        dataResult = convertResult["value"]
                        if(not pd.isna(dataResult)):
                            if(not pd.isna(field_types[i])):
                                # Forzamos todo lo que sea string a texto por cuestiones de simplicidad
                                if(field_types[i] == "string"):
                                    dataResult = str(dataResult)
                                if(not check_field_type(dataResult, field_types[i])):
                                    try:
                                        dataResult = convert_field_to_type(dataResult, field_types[i])
                                    except:
                                        if(convertResult["command"] != "constant"):
                                            raise InvalidFormat("El formato para el campo "+fields[i]+" es invalido (" + field_types[i] + ")")
                                        else:
                                            forcedErrorsFound = True  
                        else:
                            dataResult = float('NaN')

                    rowResult.append(dataResult)
                    if((not IS_SEDOL and i == 1) or (IS_SEDOL and i == 2)):
                        row_id = dataResult
                        if(not IS_SEDOL):
                            if(not valid_isin_code(row_id)):
                                raise InvalidFormat("ISIN no valido")
                except NullValue as e:
                    rowResult.append(float('NaN'))
                    #errorsFound=True
                    log.warning("Advertencia campo vacio ["+fields[i]+"]  -  " + str(e))
                    if((i == 1 and not IS_SEDOL) or (i == 2 and IS_SEDOL)):
                        errorsFound = True
                        #rowResult.append("ERROR")
                except InvalidFormat as e:
                    errorsFound = True
                    log.error("Error en el campo ["+fields[i]+"]  -  " + str(e) + " ISIN(" + str(row_id) + ")")
                    if(i != 1):
                        rowResult.append("ERROR " + str(dataResult))
                except EmptyRow as e:
                    errorsFound = True
                    rowResult.append(float('NaN'))
                    log.error(
                        "Error en el campo ["+fields[i]+"]  -  " + str(e))
                except CriticalError as e:
                    errorsFound = True
                    finished = True
                except EndOfData as e:
                    finished = True
                finally:
                    i = i+1

            if(not finished and not check_empty_row_array(rowResult)):
                # sin error y la fila no esta vacia
                if(not errorsFound and not forcedErrorsFound):
                    outputData.append(rowResult)
                # con error y isin no es null
                elif((errorsFound or forcedErrorsFound) and pd.notna(row_id) and valid_isin_code(row_id)):
                    incidenceData.append(rowResult)

            if((errorsFound and index == 0) or (index == 0 and check_empty_row_array(rowResult))):
                finished = True
                error_format_match = True

            if(VALIDATE == True):
               finished = True 

            if(index % 5 == 0):
                log.info("Se han procesado " + str(index) + " filas")
            index = index + 1
    except (InvalidFormat, CriticalError) as e:
        error_format_match = True
    return {"outputData":outputData,"incidenceData":incidenceData,"error_format_match":error_format_match}    

################################################
##  Guardar fichero excel
##  returns : True / False
################################################
def save_excel(filepath,data):
    try:
        with pd.ExcelWriter(filepath,date_format='DD/MM/YYYY',datetime_format='DD/MM/YYYY') as writer:
            data.to_excel(writer,index=False)
        return True    
    except Exception as e:
        log.error("Fallo la escritura del fichero: " + filepath + " - " +str(e))
        return False
################################################
##  Cargar un fichero excel
##  Devuelve un Dataframe con datos o uno vacio.
################################################
def load_excel(filepath,columns):
    if os.path.exists(filepath):
        try:
            dataExcel = pd.read_excel(filepath, sheet_name=0)    
            return dataExcel
        except Exception as e:
            log.error("Error leyendo el fichero: " + filepath) 
            return pd.DataFrame(columns=columns)  # Empty Dataframe
    else:        
        return pd.DataFrame(columns=columns)  # Empty Dataframe
###################################################
##  Convertir y procesar el JSON en la funcion dada
###################################################
def convert_function(string, dataExcel, index,firstColumn):    
    if(not pd.isna(string)):
        commands = json.loads(string) # TODO:Validar que sea JSON CriticalError
        try: 
            if(commands.get("value")):
                found = eval_value(commands.get("value"), dataExcel, index)
                return {"command":"value","value":found["value"]}
            if(commands.get("constant")):
                return {"command":"constant","value":commands["constant"]}
            if(commands.get("position")):
                return {"command":"position","value":eval_position(commands["position"], dataExcel)}
            if(commands.get("math")):
                return {"command":"math","value":eval_math(commands["math"], dataExcel, index)}
            if(commands.get("date")):
                return {"command":"date","value":eval_date(commands["date"], dataExcel, index)}
        except (NullValue,EndOfData,EmptyRow,CriticalError,InvalidFormat) as e:                  
            raise e                 
        except Exception as e:
            tb = traceback.format_exc()
            log.error(tb)
################################################
##  Aplicar filtro al excel
################################################
def filter_data_excel(filter_command , dataExcel):
    if(not pd.isna(filter_command)):
        commands = json.loads(filter_command) # TODO:Validar que sea JSON CriticalError
        try:
            value = commands["filter"]["column"]["value"]
            query = commands["filter"]["query"]
            query_type = commands["filter"]["type"]
        except Exception as e:
            raise InvalidFormat("[Filter] El formato de filter es invalido - "+ str(e))    

        index = 0
        finished = False            
        rowResult = []               
        while not finished:
            i = 1                                                
            try:                            
                valueFilter=eval_value(value, dataExcel, index)
                if(query_type == "equal"):
                    if(query != valueFilter["value"]):
                        rowResult.append(valueFilter["row"])                                                                    
                elif(query_type == "diff"):
                    if(query == valueFilter["value"]):
                        rowResult.append(valueFilter["row"])                     
            except (EmptyRow,NullValue) as e:                        
                pass            
            except CriticalError as e:                
                raise e    
            except EndOfData as e:                        
                finished = True
            finally:    
                index = index + 1        
    dataExcel.drop(rowResult, inplace=True)
    dataExcel.reset_index(drop=True,inplace=True)
##############################################################################
##  (VALUE)
##    params
##      - col : columna
##      - below-text : busca el texto indicado e inicia en la siguiente fila
##      - row : fila inicial (opcional)
##      - add-rows : agrega Nº filas para indicar la inicial
##      - prepend : agrega al INICIO del valor el texto indicado
##      - append : agrega al FINAL del valor el texto indicado
###############################################################################
def eval_value(commands, dataExcel, index):
    col = 0
    row = 0    
    if(commands.get("col")):
        col = int(commands["col"])-1        
        data_top = dataExcel.head()                 
        if(col not in list(data_top.columns)):
            raise InvalidFormat("[Value] La columna no es valida para el comando: " + str(commands))
    if(commands.get("below-text")):        
        belowText = commands["below-text"]
        belowText = belowText.lower()
        f = dataExcel.index[dataExcel[col].str.lower() == belowText].tolist()
        if(len(f) > 0):
            row = f[0]+1  # Si lo consigue dame la siguiente fila
        else:
            log.error("[Value] No se pudo resolver la formula para encontrar el valor :" + str(commands))
            raise CriticalError("[Value] No se pudo resolver la formula para encontrar el valor :" + str(commands))
    elif(commands.get("row")):
        row = int(commands["row"])-1 # row: = desde row hasta el final en la columna col
    
    if(commands.get("add-rows")):
        row = row+ int(commands.get("add-rows"))
    
    if(row+index in dataExcel.index):
        val = dataExcel.iloc[row+index, col]
        if(not pd.isna(val)):
            val = eval_prepend(commands , val)
            val = eval_append(commands , val)
            val = eval_replace(commands , val)
            val = val.strip() if isinstance(val , str) else val
            return {"value":val,"col":col,"row":row+index}
        else:     
            if(check_row_ifnull(dataExcel,row+index)):       
                raise EmptyRow("[Value] Valor en la posicion FILA="+str(row+index+1) + " COLUMNA=" + str(col+1) +" no es valido.")
            else:
                raise NullValue("[Value] Valor en la posicion FILA="+str(row+index+1) + " COLUMNA=" + str(col+1) +" no es valido.")
    else:             
        raise EndOfData 
################################################
##  (POSITION)
##    params
##      - col : columna
##      - row : fila
################################################
def eval_position(commands, dataExcel):
    col = 0
    row = 0
    if(commands.get("col")):
        col = int(commands["col"])-1
        data_top = dataExcel.head()                 
        if(col not in list(data_top.columns)):
            raise InvalidFormat("[Position] La columna no es valida para el comando: " + str(commands))
    if(commands.get("row")):
        row = int(commands["row"])-1
    if(row in dataExcel.index):
        val = dataExcel.iloc[row, col]
        val = eval_prepend(commands , val)
        val = eval_append(commands , val)
        val = eval_replace(commands , val) 
        val =  val.strip() if isinstance(val , str) else val
        return val
    else:
        return None    
###########################################################
##  (MATH)
##    params
##      - result : operacion matematica a realizar
##      - varXXX : operandos para la operacion , el nombre 
##          va equivale al mismo encontrado en result
###########################################################
def eval_math(commands, dataExcel, index):
    data = {}
    result = None    
    try:
        for key in commands:
            value = commands[key]
            if(key != "result"):            
                val  = eval_value(value["value"], dataExcel, index)                        
                data[key] = val["value"]            
            else:
                result = value            
        res = result    
        for val in data:
            res = res.replace(val, str(data[val]))            
        return eval(res)    
    except (NullValue,EmptyRow) as e:                                  
        raise e                        
############################################################
##  Evaluar comando DATE
##  params:  
##     - format : Formato de la fecha
##     - value o position : La posicion de la fecha 
##     - transform : Aplicar transformacion a la fecha:
##            - addDays : Sumar dias habiles (workdays)
##            - subDays : Restar dias habiles (workdays)
##     - quantity : Cantidad aplicada al operando transform
############################################################
def eval_date(commands, dataExcel, index):
    data = {}
    result = None    
    dateValue = None
    if(commands.get("value")):        
        try:
            dateValue = eval_value(commands.get("value"), dataExcel, index)                               
            dateValue = dateValue["value"]
        except NullValue as e:                          
            log.warning(e)
            raise e       
        except EmptyRow as e:
            raise e        
    elif(commands.get("position")):    
        try:
            dateValue = eval_position(commands.get("position"), dataExcel)              
        except NullValue as e:                          
            log.warning(e)
            raise e       
        except EmptyRow as e:
            raise e
    else:        
        raise InvalidFormat("[Date] Formato invalido, falta la posicion de la fecha.")
    
    if(isinstance(dateValue, str) and commands.get("format")):                              
        try:
            dateValue = datetime.datetime.strptime(dateValue, commands.get("format"))           
        except ValueError as ve:
            raise InvalidFormat("[Date] Formato invalido, el formato de la fecha no es valido - " + commands.get("format"))

    if(not check_field_type(dateValue,'date')):
        raise InvalidFormat("[Date] Formato invalido, no se encontro una fecha en la posicion.")    

    if(commands.get("transform")):
        transform = commands.get("transform")
        quantity = None
        if(commands.get("quantity")):
            quantity = commands.get("quantity")
        else:
            raise InvalidFormat("[Date] Se encontro un transform sin un quantity")          
        
        if(transform=="addDays"):
            dateValue=add_business_days(dateValue,int(quantity))
        elif(transform=="subDays"):
            dateValue=add_business_days(dateValue,-int(quantity))            
        else:
            raise InvalidFormat("[Date] Se encontro un transform no reconocido: " + transform)          
    return dateValue          
################################################
##  Anteponer un string en un valor ( Texto )
################################################
def eval_prepend(command, value):
    if(command.get("prepend")):
        value = command.get("prepend") + str(value)
    return value
################################################
##  Anexar un string a un valor ( Texto )
################################################
def eval_append(command, value):
    if(command.get("append")):
        value = str(value) + command.get("append")
    return value
################################################
##  Reemplaza el pattern con el text
################################################
def eval_replace(command, value):
    if(command.get("replace")):
        replace = command.get("replace")        
        if("pattern" in replace and  "text" in replace):
            return value.replace(replace.get("pattern"),replace.get("text"))
    return value        

def valid_isin_code(isin):
    if(not pd.isna(isin)):
        pattern_isin = "\\b([A-Z]{2})((?![A-Z]{10}\b)[A-Z0-9]{10})\\b"
        if(not pd.isna(re.search(pattern_isin, isin))):
            return True
        else:
            return False    
    else:
        return False
        

#################################################
## Verificar si una fila esta completamente vacia
#################################################
def check_row_ifnull(dataExcel, row):    
    rowData = dataExcel.iloc[row, :]    
    count = 0
    for value in rowData:        
        if(pd.isna(value)):
            count = count + 1        
    if(count==len(rowData)):
        return True
    else:
        return False
#################################################
## Sumar/Restar dias laborales a una fecha
## Params:
##   - from_date : Fecha
##   - add_days : +/- dias a sumar o restar
#################################################
def add_business_days(from_date, add_days):        
    holydays = [{'month':12,'day':25},{'month':1,'day':1}] ## Dias feriados
    business_days_to_add = add_days
    if(add_days<0):
        business_days_to_add = abs(business_days_to_add)      
    current_date = from_date
    while business_days_to_add > 0:
        if(add_days<0):
            current_date -= datetime.timedelta(days=1)
        else:
            current_date += datetime.timedelta(days=1)        

        if(next((x for x in holydays if x["day"] == current_date.day and x["month"] == current_date.month), None) != None):
            continue
        weekday = current_date.weekday()
        if weekday >= 5: # sunday = 6
            continue       
        business_days_to_add -= 1    
    return current_date
################################################
##  Verficar el tipo del campo
################################################
def check_field_type(value,field_type):        
    if(field_type == "string"):
        if(isinstance(value, str)):   
            return True            
    elif(field_type == "number"):
        if(isinstance(value, float) or isinstance(value, int)):    
            return True
    elif(field_type == "date"):               
        if(isinstance(value, pd._libs.tslibs.timestamps.Timestamp) or isinstance(value, datetime.datetime)):
            return True
    elif(field_type == "percentage"):
        if(isinstance(value, float)):
            return True
    elif(field_type == "ratio"):
        if(isinstance(value, float) or isinstance(value, int)):    
            number = str(value).split(".")            
            if(len(number[0])<=5):
                return True
            else:
                return False                
    else:
        if(not pd.isna(field_type)):        
            log.warning("Formato no reconocido: "+ field_type)        
            return False

def convert_field_to_type(value,field_type):
    if(field_type == "number"):
        if(isinstance(value, str)):             
            if(convert_float(value) != None):
                return convert_float(value)
            elif(convert_int(value) != None):
                return convert_int(value)
            else:
                raise InvalidFormat("No se puede convertir a numero")   
    elif(field_type == "ratio"):
        if(isinstance(value, str)):            
            if(convert_float(value) != None):
                return convert_float(value)
            elif(convert_int(value) != None):
                return convert_int(value)
            else:
                raise InvalidFormat("No se puede convertir a numero")           

        number = str(value).split(".")
        if(len(number[0])>5):
            raise InvalidFormat("El formato del ratio tiene mas de 5 valores enteros")               
        return value
    else:
        return value            

def convert_float(value):
    try:
        return float(value)
    except ValueError as e:
        return None               

def convert_int(value):
    try:
        return int(value)
    except ValueError as e:
        return None

################################################################
## Verificar si toda la fila esta comprendida por valores vacios
################################################################
def check_empty_row_array(array):
    count = 0
    for ar in array:
        if(pd.isna(ar)):
            count = count + 1
    if(count == len(array)):
        return True
    else:
        return False

def validate_field_types(field_types):    
    valid = ["string","number","date","percentage","ratio"]
    for tp in field_types:
        if(not pd.isna(tp)):
            if(not tp in valid):
                raise Exception("Se encontro un valor no reconocido para la columna TIPO => ["+tp+"]")
    

def validate_formulas(formulas):
    const_values=["col","below-text","row","add-rows","prepend","append","replace"]
    const_position=["col","row"]
    const_date = ["transform", "quantity", "format", "value", "position"]
    const_filter=["column","query","type"]
    for formula in formulas:
        if(not pd.isna(formula)):
            try:
                commands = json.loads(formula)                
                for com in commands:
                    if(com == "value"):
                        for key in commands["value"]:                        
                            const_values.index(key)
                    elif(com == "position"):
                        for key in commands["position"]:                        
                            const_position.index(key)
                    elif(com == "date"):
                        for key in commands["date"]:                        
                            const_date.index(key)
                            if(key == 'value'):
                                for ikey in commands["date"]["value"]:                        
                                    const_values.index(ikey)                                
                    elif(com == "math"):
                        if(commands["math"]["result"]):
                            pass
                        
                    elif(com == "constant"):
                        pass
                    else:
                        raise Exception("El JSON no es valido => "+ formula)
                
            except (ValueError,Exception) as e:  # includes simplejson.decoder.JSONDecodeError
                raise Exception("El JSON no es valido => "+ formula)


def dataframe_difference(first, second, which=None):  
    df1 = first.copy()           
    df2 = second.copy()          
    filteredColumns = df1.dtypes[df1.dtypes == np.float_]    
    listOfColumnNames = list(filteredColumns.index)
    
    N = 10000000000000000
    for column in listOfColumnNames:
        df1[column] = np.round(df1[column]*N).astype('Int64')        
        df2[column] = np.round(df2[column]*N).astype('Int64')

    comparison_df = df1.merge(df2,indicator=True,how='outer')    
    
    for column in listOfColumnNames:
        comparison_df[column] = comparison_df[column] / N     

    if which is None:
        diff_df = comparison_df[comparison_df['_merge'] != 'both']
    else:
        diff_df = comparison_df[comparison_df['_merge'] == which] 

    return diff_df.drop('_merge',axis=1)

def match_dataframes_types(df1,df2):        
    for col in df2.columns:        
        if(df2[col].dtype != df1[col].dtype):                        
            if(df2[col].isnull().sum() == len(df2[col])):
                if(df1[col].dtype == "int64"):
                    df1[col]=df1[col].astype(df2[col].dtype)
                else:
                    df2[col]=df2[col].astype(df1[col].dtype)
            elif(df1[col].isnull().sum() == len(df1[col])):                                
                if(df2[col].dtype == "int64"):
                    df2[col] = df2[col].astype(df1[col].dtype)
                else:
                    df1[col] = df1[col].astype(df2[col].dtype)
            else:
                df1[col]=df1[col].astype(df2[col].dtype)

if __name__ == '__main__':
    main()
    os._exit(0)
