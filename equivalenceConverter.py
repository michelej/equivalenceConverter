import pandas as pd # Panda Excel
import sys 
import os  
import math # funciones matematicas
import json  # JSON libreria
import simplelogging # Libreria de logs
import datetime  # Fechas
import ntpath   # Rutas
import time
import xlwings as xw

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

def main():    
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
        
        output_file = "output_"+process_type+".xlsx"
        incidence_file = "incidence_"+process_type+".xlsx"
        
        log.info("Iniciando Proceso")
        log.info("PID: "+ str(os.getpid()))
        log.info(process_type + " - ( " + process_id + " )")
                
        try:
            configExcel = pd.read_excel(config_file, process_type)
            if set(['Campos', int(process_id)]).issubset(configExcel.columns):
                campos = configExcel["Campos"]                
                tipo = configExcel["Tipo"]                                

                field_types=tipo[1:]
                fields=campos[1:]                                
                allFormulas = configExcel[int(process_id)]                
                formulas = allFormulas[1:]
                filter_func = allFormulas[0]                                
            else:
                raise Exception("No se han encontrado las columnas necesarias : (Campos ," + process_id + ")")              
        except Exception as e:
            log.error("Error: " + str(e))
            print("ERROR : Fichero de equivalencias erroneo o no encontrado")
            return

        
        log.info("Procesando fichero: " + input_file)
        
        file_extension = os.path.splitext(input_file)
        dataExcel = None
        try:
            if(file_extension[1] == ".csv"):
                dataExcel = pd.read_csv(input_file,  header=None)
            else:
                wb = xw.Book(input_file)
                active_sheet_name = wb.sheets.active.name                
                dataExcel = pd.read_excel(input_file, sheet_name=active_sheet_name, header=None)
        except Exception as e: 
            print("ERROR : No se pudo abrir el fichero")
            return        
        
        finished = False
        index = 0
        outputData = []                                   
        incidenceData = []
        
        try:
            if(not pd.isna(filter_func)):
                filter_data_excel(filter_func,dataExcel)            
            while not finished:
                i = 1
                rowResult = []               
                errorsFound = False 
                row_isin = ""
                for j in range(len(fields)):                                         
                    try:
                        dataResult=None                         
                        if(not pd.isna(formulas[i])):
                            dataResult = convert_function(formulas[i], dataExcel, index ,True if j==0 else False)  # Datos                                                                                            
                            if(not pd.isna(field_types[i])):                            
                                if(field_types[i] == "string"): # Forzamos todo lo que sea string a texto por cuestiones de simplicidad
                                    dataResult = str(dataResult)
                                if(not check_field_type(dataResult,field_types[i])):
                                    raise InvalidFormat("El formato para el campo "+fields[i]+" es invalido ("+ field_types[i] + ")")                        
                        rowResult.append(dataResult)  
                        if(i==1):
                            row_isin = dataResult
                    except NullValue as e: 
                        if(i==1): # Nullvalue en la primera columna es error en toda la fila
                            log.error("Error en el campo ["+fields[i]+"]  -  " + str(e))                            
                            errorsFound=True                                                    
                        rowResult.append(None)                                
                    except InvalidFormat as e:
                        log.error("Error en el campo ["+fields[i]+"]  -  " + str(e) + " ISIN(" + row_isin +")")                                                       
                        rowResult.append("ERROR")
                        errorsFound=True
                    except EmptyRow as e:                        
                        errorsFound=True
                        rowResult.append(None)                                                    
                    except CriticalError as e:
                        errorsFound=True                        
                        finished = True
                    except EndOfData as e:                        
                        finished = True
                    finally:    
                        i = i+1                                              
                if(not errorsFound and not finished):
                    outputData.append(rowResult) 
                if(errorsFound and not finished):                    
                    if(not check_empty_row_array(rowResult)):
                        incidenceData.append(rowResult) 

                if(errorsFound and index == 0):
                    finished = True      
                    error_format_match = True                          
                index = index + 1     
        except (InvalidFormat,CriticalError) as e:
            error_format_match = True      
                    
        if(len(outputData)>0):
            resultExcel = pd.DataFrame(outputData,columns=fields)
            resultExcel.drop_duplicates(keep='first', inplace=True)       

            historyExcel = load_excel(output_dir+output_file,fields)                         
            finalExcel = None
            equalRows = pd.DataFrame(columns=['ISIN'])            
            if(not historyExcel.empty):
                historyExcel = historyExcel.dropna(how='all')                
                finalExcel=historyExcel.append(resultExcel)                                                                        
            else:
                finalExcel=resultExcel                    
                        
            finalExcel.drop_duplicates(keep='first', inplace=True)
            finalExcel.reset_index(drop=True, inplace=True)             

            equalRows = finalExcel[finalExcel.duplicated(fields[1], keep=False)]
            equalRows.reset_index(drop=True, inplace=True)                        


            finalExcel.drop_duplicates(fields[1], inplace=True, keep=False)            
            finalExcel.reset_index(drop=True, inplace=True)            
            

            copyFinalExcel = finalExcel
            newsExcel = copyFinalExcel.append(resultExcel)
            newsExcel = newsExcel[newsExcel.duplicated()]
            
                                 
            flag_data = False
            if(not finalExcel.empty):
                if(save_excel(output_dir+output_file, finalExcel)):
                    log.info("Guardando fichero: "+output_dir+output_file)
                    log.info("Proceso finalizado...")                        

                    res_isins = newsExcel.iloc[:, 0]
                    if(len(res_isins)>0):
                        print("OK")
                        for isin in res_isins:
                            print(isin + ";OK;"+output_file)
                        flag_data = True
                else:
                    print("ERROR : Fallo en la escritura del fichero de salida.")
                    return
            
            if(len(incidenceData)>0 or not equalRows.empty):
                incidenceExcel = pd.DataFrame(incidenceData,columns=fields) 
                incidenceExcel.append(equalRows)
                historyIncidenceExcel = load_excel(output_dir+incidence_file,fields)
                finalIncidenceExcel = None
                if(not historyIncidenceExcel.empty):                        
                    finalIncidenceExcel = historyIncidenceExcel.append(incidenceExcel)                                                    
                else:
                    finalIncidenceExcel = incidenceExcel                                                

                finalIncidenceExcel.drop_duplicates(inplace=True) # Los identicos se ignoran
                finalIncidenceExcel.reset_index(drop=True,inplace=True)                                                                    

                newsIncidenceExcel = finalIncidenceExcel.append(historyIncidenceExcel)                    
                newsIncidenceExcel.drop_duplicates(keep=False,inplace=True)                    

                if(save_excel(output_dir+incidence_file,finalIncidenceExcel)):                        
                    log.info("Guardando fichero: "+output_dir+incidence_file)                                                                            
                    err_isins = newsIncidenceExcel.iloc[:, 0]   
                    if(len(err_isins)>0):
                        for isin in err_isins:
                            if(not pd.isna(isin)):                            
                                print(isin +";ERROR;"+incidence_file)     
                        flag_data = True
                else:
                    print("ERROR : Fallo en la escritura del fichero de incidencias.")                    
            
            if(flag_data == False):
                print("ERROR: No se encontraron datos para procesar")
        else:
            if(error_format_match == True):
                print("ERROR : El formato para este excel no es correcto: " + process_type + " - " + process_id)
            else:    
                print("ERROR")            
    elapsed_time = time.time() - start_time
    log.info("Tiempo de ejecución: " + str(elapsed_time) +" segundos")
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
                return found["value"]
            if(commands.get("constant")):
                return commands["constant"]
            if(commands.get("position")):
                return eval_position(commands["position"], dataExcel)
            if(commands.get("math")):
                return eval_math(commands["math"], dataExcel, index)            
            if(commands.get("date")):
                return eval_date(commands["date"], dataExcel, index)            
        except (NullValue,EndOfData,EmptyRow,CriticalError,InvalidFormat) as e:                  
            raise e                 
        except Exception as e:
            log.error(e)
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
            errorsFound = False                         
            try:                            
                valueFilter=eval_value(value, dataExcel, index)
                if(query_type == "equal"):
                    if(query != valueFilter["value"]):
                        rowResult.append(valueFilter["row"])                                                                    
                elif(query_type == "diff"):
                    if(query == valueFilter["value"]):
                        rowResult.append(valueFilter["row"])                     
            except EmptyRow as e:                        
                errorsFound=True       
            except NullValue as e:  
                log.error(e)
                raise CriticalError()             
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
    if(commands.get("below-text")):        
        belowText = commands["below-text"]
        f = dataExcel.index[dataExcel[col] == belowText].tolist()
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
            return {"value":val,"col":col,"row":row+index}
        else:     
            if(check_row_ifnull(dataExcel,row+index)):       
                raise EmptyRow()
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
    if(commands.get("row")):
        row = int(commands["row"])-1
    if(row in dataExcel.index):    
        return dataExcel.iloc[row, col]
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

    if(not check_field_type(dateValue,'date')):
        raise InvalidFormat("[Date] Formato invalido, no se encontro una fecha en la posicion.")

    if(isinstance(dateValue, str) and commands.get("format")):                              
        dateValue = datetime.datetime.strptime(dateValue, commands.get("format"))           

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
        if(isinstance(value, pd._libs.tslibs.timestamps.Timestamp)):
            return True
    else:
        if(not pd.isna(field_type)):        
            log.warning("Formato no reconocido: "+ field_type)
        
    return False
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

if __name__ == '__main__':
    main()
