# Comparador de versiones de archivos Excel
Para cuando te descargas un Excel de un servidor, por ejemplo, realizas cambios en el archivo y luego quieres compararlo con la versión del Excel original en el servidor antes de subirlo y sustituirlo.
En este caso primero se convertirán los Excel(.xlsx, .xls) a formato (.csv) y luego se compararán los archivos usando el software [WinMerge](https://winmerge.org/downloads/?lang=es)

## Guía del usuario

1. Descargar e instalar [Python](https://www.python.org/downloads/)
   - Descargar y ejecutar el archivo python-3.12.5-amd64.exe eligiendo la "Customize installation"
     ![imagen](https://github.com/user-attachments/assets/ad9ecfb2-20cf-49bb-ad15-801cff87d9e0)
     ![imagen](https://github.com/user-attachments/assets/c78de7e1-6865-41bc-8bc9-d1234ea6322a)
     ![imagen](https://github.com/user-attachments/assets/25537d81-58d1-43d2-b500-7ef04dbed05f)

3. Convertir Excel a CSV usando el convert_excel_to_csv.exe
   - Descargar el convert_excel_to_csv.exe
   - Ejecutarlo. Va a pedir que selecciones 1 carpeta y va a generar otra carpeta dentro de esa que seleccionaste con todos los archivos .xlsx y .xls convertidos a .csv. Lo va a pedir 2 veces.

   - Otra alternativa es ejecutar el archivo convert_excel_to_csv.py que está en la carpeta "code" desde la terminal directamente. Para ello hay que ir a la carpeta dónde lo tenemos descargado, pulsar el botón derecho del ratón, elegir "Abrir en Terminal" y poner el siguiente comando: python convert_excel_to_csv.py

4. Comparar Excels usando WinMerge
   - Descargar el WinMerge-2.16.42.1-x64-Setup.exe
   - Instalarlo
   - Abrir las 2 carpetas que generamos antes con el convert_excel_to_csv.exe

## Información técnica del convert_excel_to_csv.exe
Este script convierte archivos Excel en archivos CSV. Realiza las siguientes tareas:

 * Instala automáticamente las dependencias necesarias (pandas y openpyxl) si no están presentes.
 * Permite al usuario seleccionar una carpeta que contenga archivos Excel.
 * Convierte cada hoja de cálculo en cada archivo Excel a un archivo CSV separado.
 * Guarda los archivos CSV en una subcarpeta creada dentro de la carpeta seleccionada.

Propósito de las Bibliotecas Utilizadas

 * pandas: Para la manipulación de datos, leer archivos Excel y escribir CSV.
 * openpyxl: Para leer archivos Excel .xlsx con pandas.
 * tkinter: Para la interfaz gráfica de selección de carpetas.
 * os y sys: Para gestionar archivos y controlar la ejecución del programa.
 * subprocess: Para instalar automáticamente dependencias si es necesario.

El código está disponible en la carpeta 'code'
