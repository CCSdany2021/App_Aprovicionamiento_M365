import pandas as pd
import os
import sys
import subprocess
import time
import json
from datetime import datetime

sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config

class CreadorReglasReenvio:
    """Clase para gestionar la creación de reglas de reenvío masivas vía PowerShell"""
    
    def __init__(self):
        self.ps_script_path = os.path.join(os.path.dirname(__file__), 'powershell', 'Crear-ReglasReenvio.ps1')
        self.resultados = {
            "total": 0,
            "procesados": 0,
            "errores": 0,
            "detalles": []
        }

    def ejecutar(self, filepath):
        """Generador para SSE"""
        yield {"status": "info", "message": "Iniciando proceso de Reglas de Reenvío CCS...", "progress": 5}
        
        try:
            # Validar archivo
            if filepath.endswith('.xlsx'):
                df = pd.read_excel(filepath, dtype=str)
            else:
                df = pd.read_csv(filepath, sep=None, engine='python', dtype=str)
            
            # Normalizar columnas y datos
            df.columns = df.columns.str.strip().upper()
            
            # Mapeo de columnas corregido para priorizar UPN
            # Buscamos columnas que contengan estas palabras clave
            target_cols = {
                'UPN': ['UPN', 'USERPRINCIPALNAME', 'CORREO', 'EMAIL'],
                'CODIGO': ['CODIGO', 'ID'],
                'CORREOMAMA': ['MAMA', 'MADRE'],
                'CORREOPAPA': ['PAPA', 'PADRE'],
                'CORREOACUDIENTE': ['ACUDIENTE', 'GUARDIAN']
            }
            
            new_df = pd.DataFrame()
            for key, keywords in target_cols.items():
                found_col = next((c for c in df.columns if any(k in c for k in keywords)), None)
                if found_col:
                    new_df[key] = df[found_col].astype(str).str.strip().str.lower()
            
            # Si no hay UPN pero hay Código, el UPN es el Código + Dominio
            if 'UPN' not in new_df.columns and 'CODIGO' in new_df.columns:
                new_df['UPN'] = new_df['CODIGO'].apply(lambda x: f"{x}@{config.COLEGIO_DOMINIO}" if x and '@' not in x else x)
            
            if 'UPN' not in new_df.columns:
                raise Exception("No se pudo identificar la columna de UPN o Código en el archivo")

            self.resultados["total"] = len(new_df)
            new_df.to_csv(temp_csv, index=False)
            
            yield {"status": "info", "message": f"Archivo validado: {len(df)} registros encontrados.", "progress": 15}
            yield {"status": "process", "message": "Conectando con Exchange Online...", "progress": 20}

            # Ejecutar PowerShell
            cmd = [
                'powershell.exe', '-ExecutionPolicy', 'Bypass', '-File', 
                self.ps_script_path, 
                '-ArchivoEntrada', temp_csv, 
                '-Dominio', config.COLEGIO_DOMINIO
            ]

            process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding='cp850', errors='replace')
            
            count_success = 0
            while True:
                line = process.stdout.readline()
                if not line and process.poll() is not None:
                    break
                if line:
                    line = line.strip()
                    if '✅' in line:
                        count_success += 1
                        yield {"status": "log", "message": line}
                    elif '❌' in line:
                        self.resultados["errores"] += 1
                        yield {"status": "log", "message": line}
                    elif 'Procesando buzón:' in line:
                        self.resultados["procesados"] += 1
                        progress = int(20 + (self.resultados["procesados"] / self.resultados["total"]) * 75)
                        yield {"status": "process", "message": line, "progress": progress}

            process.wait()
            
            # Limpiar
            try: os.remove(temp_csv)
            except: pass

            yield {"status": "complete", "message": "Proceso de reglas CCS finalizado.", "progress": 100, "results": self.resultados}

        except Exception as e:
            yield {"status": "error", "message": f"Error crítico: {str(e)}"}

if __name__ == "__main__":
    pass
