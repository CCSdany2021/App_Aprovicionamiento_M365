import pandas as pd
import sys

file_path = r'c:\aprovisionamientoEstudiantes\archivos_procesar\plantilla_creacion_teams_jardin_2026.xlsx'

try:
    xl = pd.ExcelFile(file_path)
    print(f"Sheet names: {xl.sheet_names}")
    
    for sheet in xl.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet, nrows=0) # Read only headers
        print(f"\n--- Sheet: {sheet} ---")
        print("Columns found:")
        for col in df.columns:
            print(f" - '{col}'")
            
except Exception as e:
    print(f"Error reading file: {e}")
