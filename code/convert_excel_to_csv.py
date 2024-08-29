import os
import subprocess
import sys
import pandas as pd
import openpyxl
import tkinter as tk
from tkinter import filedialog, messagebox

# Instalación de pandas y openpyxl
def install_package(package):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
    except subprocess.CalledProcessError:
        print(f"Error al instalar {package}")
        sys.exit(1)

try:
    import pandas as pd
    import openpyxl
except ImportError:
    print("Instalando dependencias necesarias...")
    install_package('pandas')
    install_package('openpyxl')
    import pandas as pd
    import openpyxl

# Selección de carpeta
def select_directory():
    root = tk.Tk()
    root.withdraw()
    folder_selected = filedialog.askdirectory()
    if not folder_selected:
        messagebox.showwarning("Advertencia", "No se ha seleccionado ninguna carpeta. El programa se cerrará.")
        sys.exit(1)
    return folder_selected

# Conversor
def process_files():
    input_folder = select_directory()
    print(f"Directorio de entrada: {input_folder}", file=sys.stderr)

    input_folder_name = os.path.basename(input_folder)
    output_folder_name = f"{input_folder_name}_csv"
    output_folder = os.path.join(input_folder, output_folder_name)
    print(f"Directorio de salida: {output_folder}", file=sys.stderr)

    try:
        os.makedirs(output_folder, exist_ok=True)
    except Exception as e:
        print(f"Error al crear la carpeta de salida: {e}", file=sys.stderr)
        sys.exit(1)

    files_processed = False
    for filename in os.listdir(input_folder):
        file_path = os.path.join(input_folder, filename)
        if filename.endswith('.xlsx') or filename.endswith('.xls'):
            files_processed = True
            print(f"Procesando archivo Excel: {filename}", file=sys.stderr)
            try:
                excel_file = pd.ExcelFile(file_path)
                for sheet_name in excel_file.sheet_names:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    csv_filename = f"{os.path.splitext(filename)[0]}_{sheet_name}.csv"
                    csv_path = os.path.join(output_folder, csv_filename)
                    df.to_csv(csv_path, index=False)
                    print(f"Archivo CSV guardado: {csv_path}", file=sys.stderr)
            except Exception as e:
                print(f"Error al procesar el archivo {filename}: {e}", file=sys.stderr)

    if not files_processed:
        print("No se encontraron archivos Excel en la carpeta de entrada.", file=sys.stderr)

    print("Conversión completada.", file=sys.stderr)

# Ejecutarlo 2 veces
process_files()
process_files()
