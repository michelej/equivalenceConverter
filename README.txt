Argumentos:

- "Tipo"
- "Columna de la equivalencia en el excel"
- "Fichero a procesar"
- "Directorio donde se va a dejar el fichero de salida, finalizado con //"
- "Fichero de equivalencias"

Como ejecutar el EXE?
> equivalenceConverter.exe "FINALES_UK" "16" "FICHERO" "DESTINO"
> py equivalenceConverter.py


Para programar se necesita:
Instalar:
> pip install panda
> pip install openpyxl
> pip install simplelogging


Salidas a consola:

ERROR : faltan argumentos..  <Tipo Proceso> <ID Proceso> <Fichero> <Directorio Destino>
ERROR : No se encuentra la carpeta de destino: 
ERROR : No se encuentra el fichero a procesar:
ERROR : Fichero de equivalencias erroneo o no encontrado
ERROR : No se pudo abrir el fichero a procesar
ERROR : Fallo en la escritura del fichero de salida.
ERROR : Fallo en la escritura del fichero de incidencias.
NO DATA: No se encontraron datos para procesar
ERROR : El formato para este excel no es correcto: " + process_type + " - " + process_id
ERROR