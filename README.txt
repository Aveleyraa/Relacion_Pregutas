Instrucciones:
Descargar todos los archivos que terminan con .py
Los archivos excel son a manera de ejemplo de lo que se hace.

Asegurarse de contar con un cuestionario previamente marcado con las preguntas que se
relacionan. También contar con el mismo cuestionario en blanco (o con otras validaciones)
dentro de la misma carpeta en caso de que se requiera meter validaciones.

Primer paso: abrir archivo "Encontrar.py", sustituir string de variable "libro" con
el path realtivo (el nombre del archivo) del cuestionario marcado previamente. A continuación
ejecutar el script.
Esto va a generar un archivo de salida llamado "PR_(nombre en variable libro).csv"
Ese archivo generado será útil para los otros dos procesos siguientes. Se trata 
del archivo guía.

Proceso para meter validaciones en cuestionario:
Abrir script llamado "validar_p_rel". En variable "guia" sustituir la string por otra que 
contenga el nombre (o path relativo) del archivo guia previamente generado con script 
"encontrar.py". Sustituir también string de variable "archivo" por el nombre del 
cuestionario en el que se imprimirán las validaciones. 
Con el nombre de ambos documentos actualizados en sus respectivas variables, ejecutar el 
script. Esto generará un nuevo archivo llamado: 
"(lo que se puso en variable archivo)_vprel.xlsx"
Ese cuestionario ya cuenta con las formulas puestas en cada pregunta que fue marcada con la 
letra "W" en el proceso de marcar el cuestionario que se hizo incluso antes de ejecutar
el scrpit llamado "encontrar.py"


Proceso para revisión de cuestionario contestado:
Una vez que los informantes han regresado el cuestionario contestado, pasar tal documento a
la carpeta dodne están los scripts de este proceso. Abrir el script llamado "revision.py".
En seguida actualizar string de variable "documento" con el nombre o path relativo del
cuestionario contestado por los informantes. También actualizar la string de variabe "guia"
con el nombre o path relativo del documento guia que fue generado con el script 
"encontrar.py" en un proceso previo.
Cuando ambas variables fueron actualizadas, se ejectua el script. Esto genera un archivo con
el nombre "resultado(lo que se escribió en variable 'documento').csv"
En este documento generado, hay una columna llamada "resultado" que está expresada en
ceros y unos. Los 1 significan error mientras que los 0 significan que se cumple
la condición de comprobación. También hay valores NA en esta columna, pero ellos 
corresponden a las celdas que son referentes y no necesitan comparación.

 
 