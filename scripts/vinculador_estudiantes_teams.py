import pandas as pd
import requests
import urllib3
from datetime import datetime
import os
import sys
import time

sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class VinculadorEstudiantesTeams:
    """Aprovisiona estudiantes a equipos de Teams basándose en su curso"""
    
    def __init__(self):
        try:
            config.validar_configuracion()
        except:
            pass
        
        self.token = None
        self.teams_encontrados = []
        self.usuarios_cache = {}
        
        self.resultados = {
            "total_estudiantes": 0,
            "total_equipos": 0,
            "estudiantes_vinculados": 0,
            "estudiantes_ya_en_equipo": 0,
            "estudiantes_no_encontrados": 0,
            "equipos_no_encontrados": 0,
            "errores_vinculacion": 0,
            "errores": [],
            "detalles_equipos": [],
            "estudiantes_procesados": []
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
            self.resultados["errores"].append(f"Error de token: {str(e)}")
            return False

    def obtener_todos_los_teams(self) -> bool:
        """Obtiene todos los grupos/teams que siguen el patrón 'CURSO: MATERIA'"""
        if not self.token:
            return False
        
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        print("\n🔍 Buscando equipos de Teams con formato 'CURSO: MATERIA'...")
        
        try:
            teams = []
            # Obtenemos grupos. Graph API no permite filtrar fácilmente por 'contains' y ':' al mismo tiempo con patrones complejos,
            # así que traemos los que parecen ser equipos y filtramos en Python.
            url = f"{config.GRAPH_ENDPOINT}/groups?$select=id,displayName&$top=999"
            
            while url:
                response = requests.get(url, headers=headers, verify=False, timeout=15)
                if response.status_code == 200:
                    data = response.json()
                    for group in data.get('value', []):
                        display_name = group.get('displayName', '')
                        if ": " in display_name:
                            teams.append(group)
                    url = data.get('@odata.nextLink')
                else:
                    break
            
            self.teams_encontrados = teams
            self.resultados["total_equipos"] = len(teams)
            print(f"✅ {len(teams)} equipos encontrados con el formato requerido")
            return True
        except Exception as e:
            self.resultados["errores"].append(f"Error obteniendo teams: {str(e)}")
            return False

    def cargar_archivo(self, ruta_archivo: str) -> pd.DataFrame:
        """Carga el archivo Excel/CSV de estudiantes"""
        print(f"DEBUG: Cargando archivo desde: {ruta_archivo}")
        try:
            if not ruta_archivo:
                raise ValueError("Ruta de archivo vacía")
            ruta_lower = ruta_archivo.lower()
            if ruta_lower.endswith(".xlsx"):
                df = pd.read_excel(ruta_archivo, dtype=str)
            elif ruta_lower.endswith(".csv"):
                # Intentar varios delimitadores comunes
                try:
                    df = pd.read_csv(ruta_archivo, sep=';', dtype=str, encoding="utf-8")
                except:
                    df = pd.read_csv(ruta_archivo, sep=',', dtype=str, encoding="utf-8")
            else:
                raise ValueError(f"Formato no soportado: {os.path.basename(ruta_archivo)}")
            
            df.columns = df.columns.str.strip()
            df = df.fillna("")
            self.resultados["total_estudiantes"] = len(df)
            return df
        except Exception as e:
            raise Exception(f"Error cargando archivo: {e}")

    def detectar_columnas(self, df: pd.DataFrame) -> tuple:
        """Detecta las columnas de Código/UPN y Curso/Grado"""
        col_id = None
        col_curso = None
        
        # Posibles nombres para identificación del estudiante
        ids = ['CODIGO', 'codigo', 'UserPrincipalName', 'UPN', 'upn', 'Email', 'Mail', 'CODIGO_ESTUDIANTE']
        # Posibles nombres para el curso
        cursos = ['CURSO', 'curso', 'jobtitle', 'JOBTITLE', 'GRADO', 'grado']
        
        for col in df.columns:
            if col in ids and not col_id:
                col_id = col
            if col in cursos and not col_curso:
                col_curso = col
        
        if not col_id or not col_curso:
            # Si no los encuentra por nombre exacto, intentar por contenido parcial
            for col in df.columns:
                if 'CODIGO' in col.upper() or 'UPN' in col.upper(): col_id = col
                if 'CURSO' in col.upper() or 'GRADO' in col.upper() or 'JOBTITLE' in col.upper(): col_curso = col
        
        if not col_id or not col_curso:
            raise ValueError(f"No se detectaron las columnas necesarias. Encontradas: {list(df.columns)}")
            
        return col_id, col_curso

    def obtener_user_id(self, identifier: str) -> str or None:
        """Obtiene el ID de objeto de un usuario por su UPN o Código"""
        if not self.token or not identifier:
            return None
        
        # Asegurar formato UPN si es solo código
        upn = identifier if '@' in identifier else f"{identifier}@{config.COLEGIO_DOMINIO}"
        upn = upn.strip().lower()
        
        if upn in self.usuarios_cache:
            return self.usuarios_cache[upn]
        
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        try:
            url = f"{config.GRAPH_ENDPOINT}/users/{upn}?$select=id"
            response = requests.get(url, headers=headers, verify=False, timeout=10)
            if response.status_code == 200:
                uid = response.json().get("id")
                self.usuarios_cache[upn] = uid
                return uid
        except:
            pass
        return None

    def agregar_a_team(self, team_id: str, user_id: str) -> tuple:
        """Agrega al usuario como Member al Team"""
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        # Primero intentamos añadirlo a los miembros del grupo
        url = f"{config.GRAPH_ENDPOINT}/groups/{team_id}/members/$ref"
        body = {"@odata.id": f"https://graph.microsoft.com/v1.0/directoryObjects/{user_id}"}
        
        try:
            response = requests.post(url, json=body, headers=headers, verify=False, timeout=10)
            if response.status_code == 204:
                return True, "Vinculado"
            elif response.status_code == 400:
                # Comprobar si es porque ya existe
                error = response.json().get('error', {})
                if "already exists" in error.get('message', '').lower():
                    return True, "Ya en equipo"
                return False, f"Error: {error.get('message')}"
            else:
                return False, f"Status {response.status_code}"
        except Exception as e:
            return False, str(e)

    def ejecutar(self, ruta_archivo: str):
        """Proceso principal - GENERADOR para progreso en tiempo real"""
        yield {"status": "info", "message": "Iniciando vinculación manual...", "progress": 0}
        
        try:
            if not self.obtener_token():
                raise Exception("No se pudo obtener token")
            
            yield {"status": "info", "message": "Obteniendo lista de equipos...", "progress": 10}
            if not self.obtener_todos_los_teams():
                raise Exception("No se pudieron obtener equipos de Teams")
            
            yield {"status": "info", "message": "Cargando archivo de estudiantes...", "progress": 15}
            df = self.cargar_archivo(ruta_archivo)
            col_id, col_curso = self.detectar_columnas(df)
            
            # Organizar estudiantes por curso
            estudiantes_por_curso = {}
            for _, row in df.iterrows():
                id_val = str(row[col_id]).strip()
                curso_val = str(row[col_curso]).strip()
                if id_val and curso_val:
                    if curso_val not in estudiantes_por_curso:
                        estudiantes_por_curso[curso_val] = []
                    estudiantes_por_curso[curso_val].append(id_val)
            
            total_equipos = len(self.teams_encontrados)
            yield {"status": "process", "message": f"Procesando {total_equipos} equipos...", "progress": 20}
            
            # Procesar cada equipo encontrado
            for idx, team in enumerate(self.teams_encontrados):
                display_name = team['displayName']
                team_id = team['id']
                curso_equipo = display_name.split(":")[0].strip()
                
                percent = int(20 + (idx / total_equipos) * 75)
                yield {"status": "process", "message": f"[{idx+1}/{total_equipos}] Trabajando en: {display_name}", "progress": percent}
                
                if curso_equipo in estudiantes_por_curso:
                    lista_estudiantes = estudiantes_por_curso[curso_equipo]
                    vinculados_equipo = 0
                    errores_equipo = 0
                    
                    for est_id in lista_estudiantes:
                        u_id = self.obtener_user_id(est_id)
                        if not u_id:
                            self.resultados["estudiantes_no_encontrados"] += 1
                            self.resultados["estudiantes_procesados"].append({"Estudiante": est_id, "Equipo": display_name, "Resultado": "No encontrado"})
                            continue
                            
                        exito, msg = self.agregar_a_team(team_id, u_id)
                        if exito:
                            if msg == "Ya en equipo":
                                self.resultados["estudiantes_ya_en_equipo"] += 1
                            else:
                                self.resultados["estudiantes_vinculados"] += 1
                                vinculados_equipo += 1
                                if vinculados_equipo % 5 == 0: # Log cada 5 estudiantes para no saturar
                                    yield {"status": "log", "message": f"   ✅ {est_id} vinculado a {curso_equipo}"}
                        else:
                            self.resultados["errores_vinculacion"] += 1
                            errores_equipo += 1
                        
                        self.resultados["estudiantes_procesados"].append({"Estudiante": est_id, "Equipo": display_name, "Resultado": msg})
                    
                    self.resultados["detalles_equipos"].append({
                        "Equipo": display_name, "Estudiantes": len(lista_estudiantes), "Vinculados": vinculados_equipo, "Errores": errores_equipo
                    })
            
            yield {"status": "info", "message": "Finalizando proceso...", "progress": 98}
            self.guardar_logs()
            yield {"status": "complete", "message": "Vinculación finalizada con éxito", "progress": 100, "results": self.resultados}
            
        except Exception as e:
            yield {"status": "error", "message": str(e)}
            self.resultados["errores"].append(str(e))
            yield {"status": "complete", "results": self.resultados}

    def guardar_logs(self):
        """Guarda los resultados en un archivo de log"""
        try:
            os.makedirs(config.CARPETA_LOGS, exist_ok=True)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            log_file = os.path.join(config.CARPETA_LOGS, f'vinculacion_teams_{timestamp}.log')
            
            with open(log_file, 'w', encoding='utf-8') as f:
                f.write("REPORTE DE VINCULACIÓN DE ESTUDIANTES A TEAMS\n")
                f.write(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("="*70 + "\n\n")
                f.write(f"Total Estudiantes en Archivo: {self.resultados['total_estudiantes']}\n")
                f.write(f"Total Equipos Identificados: {self.resultados['total_equipos']}\n")
                f.write(f"Vínculos Exitantes/Nuevos: {self.resultados['estudiantes_vinculados']}\n")
                f.write(f"Ya eran miembros: {self.resultados['estudiantes_ya_en_equipo']}\n")
                f.write(f"Estudiantes no encontrados: {self.resultados['estudiantes_no_encontrados']}\n")
                f.write(f"Errores de vinculación: {self.resultados['errores_vinculacion']}\n\n")
                
                f.write("RESUMEN POR EQUIPO:\n")
                for det in self.resultados["detalles_equipos"]:
                    f.write(f"- {det['Equipo']}: {det['Vinculados']} vinculados, {det['Errores']} errores de {det['Estudiantes']} totales.\n")
                
                if self.resultados["errores"]:
                    f.write("\nERRORES CRÍTICOS:\n")
                    for err in self.resultados["errores"]:
                        f.write(f"!!! {err}\n")
            
            print(f"📝 Log guardado en: {log_file}")
        except Exception as e:
            print(f"⚠️ No se pudo guardar el log: {e}")

if __name__ == "__main__":
    # Prueba rápida si se ejecuta directamente
    vinculador = VinculadorEstudiantesTeams()
    # ruta = "archivos/Estudiantes_Prueba.xlsx"
    # vinculador.ejecutar(ruta)
