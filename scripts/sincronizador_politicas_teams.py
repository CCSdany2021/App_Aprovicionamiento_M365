import pandas as pd
import requests
import urllib3
import os
import sys
import json
import subprocess
import time
from datetime import datetime

sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class SincronizadorPoliticasTeams:
    """Sincroniza paquetes de políticas de Teams (automático o manual)"""
    
    def __init__(self, departamento_filtro=None):
        try:
            config.validar_configuracion()
        except:
            pass
        
        self.token = None
        # Obtener periodo de la configuración dinámica
        default_dept = f"Estudiantes {config.PERIODO_ACTUAL}"
        self.departamento_filtro = departamento_filtro if departamento_filtro else default_dept
        
        self.package_student = "Education_SecondaryStudent"
        self.ps_script_path = os.path.join(os.path.dirname(__file__), 'powershell', 'Asignar-PoliticasEstudiantes.ps1')
        
        self.resultados = {
            "total": 0,
            "exitosos": 0,
            "ya_asignados": 0,
            "errores": 0,
            "archivo_log": "",
            "detalles": []
        }

    def obtener_token(self) -> bool:
        """Obtiene token de acceso"""
        url = f"https://login.microsoftonline.com/{config.TENANT_ID}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": config.CLIENT_ID,
            "client_secret": config.CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default"
        }
        try:
            response = requests.post(url, data=data, verify=False, timeout=10)
            response.raise_for_status()
            self.token = response.json()["access_token"]
            return True
        except Exception as e:
            return False

    def _consultar_usuarios(self, filtro):
        """Helper paginado"""
        headers = {"Authorization": f"Bearer {self.token}", "Content-Type": "application/json"}
        estudiantes = []
        url = f"{config.GRAPH_ENDPOINT}/users?{filtro}&$select=userPrincipalName&$top=999"
        
        try:
            while url:
                response = requests.get(url, headers=headers, verify=False, timeout=20)
                if response.status_code == 200:
                    data = response.json()
                    estudiantes.extend([u['userPrincipalName'] for u in data.get('value', [])])
                    url = data.get('@odata.nextLink')
                else: break
            return estudiantes
        except:
            return []

    def obtener_estudiantes_del_tenant(self) -> list:
        """Obtiene UPNs con búsqueda inteligente"""
        if not self.obtener_token():
            return []
            
        # 1. Intento original
        filtro = f"$filter=department eq '{self.departamento_filtro}'"
        print(f"🔍 Buscando políticas para: '{self.departamento_filtro}'")
        estudiantes = self._consultar_usuarios(filtro)
        
        # 2. Intento Plural/Singular si falló
        if not estudiantes:
            alternativa = ""
            if self.departamento_filtro.lower().startswith("estudiante "):
                alternativa = self.departamento_filtro.replace("Estudiante ", "Estudiantes ")
            elif self.departamento_filtro.lower().startswith("estudiantes "):
                alternativa = self.departamento_filtro.replace("Estudiantes ", "Estudiante ")
            
            if alternativa:
                 print(f"⚠️ Probando alternativa: '{alternativa}'")
                 filtro_alt = f"$filter=department eq '{alternativa}'"
                 estudiantes = self._consultar_usuarios(filtro_alt)
                 if estudiantes:
                     self.departamento_filtro = alternativa
        
        return estudiantes

    def ejecutar(self, filepath=None):
        """Generador para streaming SSE. Si filepath es None, hace escaneo automático."""
        modo = "Automático (Tenant)" if filepath is None else f"Manual ({os.path.basename(filepath)})"
        yield {"status": "info", "message": f"Iniciando proceso en modo: {modo}", "progress": 5}
        
        upns = []
        if filepath:
            # Modo Manual: Cargar del archivo
            try:
                if filepath.endswith('.xlsx'):
                    df = pd.read_excel(filepath, dtype=str)
                else:
                    df = pd.read_csv(filepath, sep=None, engine='python', dtype=str)
                
                # Detectar columna de identificación
                col_id = next((c for c in df.columns if c.upper() in ['CODIGO', 'USERPRINCIPALNAME', 'UPN', 'EMAIL']), None)
                if not col_id:
                    raise Exception("No se encontró columna de Código o UPN en el archivo")
                
                for val in df[col_id].dropna():
                    val = str(val).strip()
                    if '@' not in val:
                        val = f"{val}@{config.COLEGIO_DOMINIO}"
                    upns.append(val)
            except Exception as e:
                yield {"status": "error", "message": f"Error leyendo archivo: {str(e)}"}
                return
        else:
            # Modo Automático: Escanear Tenant
            yield {"status": "info", "message": f"Buscando estudiantes en el departamento '{self.departamento_filtro}'...", "progress": 15}
            upns = self.obtener_estudiantes_del_tenant()
            if not upns:
                yield {"status": "warning", "message": "No se encontraron estudiantes para procesar automáticamente."}
                yield {"status": "complete", "results": self.resultados, "progress": 100}
                return

        self.resultados["total"] = len(upns)
        yield {"status": "info", "message": f"Se procesarán {len(upns)} usuarios.", "progress": 20}

        # Crear archivo temporal para el script de PowerShell
        temp_dir = os.path.join(config.CARPETA_RESULTADOS, 'temp')
        os.makedirs(temp_dir, exist_ok=True)
        temp_csv = os.path.join(temp_dir, f"temp_upns_{int(time.time())}.csv")
        pd.DataFrame({'UserPrincipalName': upns}).to_csv(temp_csv, index=False)

        # Ejecutar PowerShell
        yield {"status": "process", "message": "Lanzando motor de políticas (PowerShell)...", "progress": 25}
        
        # Comando para ejecutar el script .ps1
        # Usamos -File para que sepa que es un script y le pasamos los parámetros
        cmd = [
            'powershell.exe', '-ExecutionPolicy', 'Bypass', '-File', 
            self.ps_script_path, 
            '-ArchivoEstudiantes', temp_csv, 
            '-Dominio', config.COLEGIO_DOMINIO
        ]

        try:
            # Ejecutamos y capturamos la salida en tiempo real
            process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding='cp850', errors='replace')
            
            while True:
                line = process.stdout.readline()
                if not line and process.poll() is not None:
                    break
                if line:
                    line = line.strip()
                    # Parsear la salida del script para actualizar progreso y logs
                    if '✅ Política asignada' in line:
                        self.resultados["exitosos"] += 1
                        yield {"status": "log", "message": line}
                    elif '⚠️  Ya tiene la política' in line:
                        self.resultados["ya_asignados"] += 1
                        yield {"status": "log", "message": line}
                    elif '❌' in line:
                        self.resultados["errores"] += 1
                        yield {"status": "log", "message": line}
                    elif 'Procesando:' in line:
                        # Extraer progreso si es posible (ej: [1/150])
                        yield {"status": "process", "message": line, "progress": 25 + int((self.resultados["exitosos"] + self.resultados["ya_asignados"] + self.resultados["errores"]) / len(upns) * 70)}

            process.wait()
            
            # Limpiar temporal
            try: os.remove(temp_csv)
            except: pass

            yield {"status": "complete", "message": "Proceso de políticas finalizado", "progress": 100, "results": self.resultados}
            
        except Exception as e:
            yield {"status": "error", "message": f"Error en ejecución: {str(e)}"}

if __name__ == "__main__":
    # Test
    sinc = SincronizadorPoliticasTeams()
    # for update in sinc.ejecutar(): print(update)
