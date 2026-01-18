import requests
import urllib3
from datetime import datetime
import os
import sys
import time

sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class SincronizadorAutomaticoTeams:
    """Sincroniza estudiantes directamente desde el Tenant a sus equipos de Teams"""
    
    def __init__(self, departamento_filtro=None):
        try:
            config.validar_configuracion()
        except:
            pass
        
        self.token = None
        # Obtener periodo de la configuración dinámica
        default_dept = f"Estudiantes {config.PERIODO_ACTUAL}"
        self.departamento_filtro = departamento_filtro if departamento_filtro else default_dept
        
        print(f"DEBUG: Sincronizador inicializado con Departamento='{self.departamento_filtro}'")
        
        self.teams_encontrados = []
        
        self.resultados = {
            "total_estudiantes_tenant": 0,
            "total_equipos": 0,
            "estudiantes_vinculados": 0,
            "estudiantes_ya_en_equipo": 0,
            "estudiantes_sin_curso": 0,
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
        """Obtiene todos los equipos que siguen el patrón 'CURSO: MATERIA'"""
        if not self.token:
            return False
        
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        try:
            teams = []
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
            return True
        except Exception as e:
            self.resultados["errores"].append(f"Error obteniendo teams: {str(e)}")
            return False


    def _consultar_usuarios_por_filtro(self, filtro, headers):
        """Helper para consultar usuarios paginados"""
        estudiantes = []
        url = f"{config.GRAPH_ENDPOINT}/users?{filtro}&$select=id,userPrincipalName,jobTitle,displayName&$top=999"
        
        while url:
            response = requests.get(url, headers=headers, verify=False, timeout=20)
            if response.status_code == 200:
                data = response.json()
                estudiantes.extend(data.get('value', []))
                url = data.get('@odata.nextLink')
            else:
                break
        return estudiantes

    def obtener_estudiantes_del_tenant(self) -> list:
        """Obtiene todos los usuarios (Intento inteligente Singular/Plural)"""
        try:
            if not self.token: return []
                
            headers = {
                "Authorization": f"Bearer {self.token}",
                "Content-Type": "application/json",
                "ConsistencyLevel": "eventual"
            }
            
            # 1. Intento original
            filtro_original = f"$filter=department eq '{self.departamento_filtro}'"
            print(f"🔍 Buscando estudiantes: '{self.departamento_filtro}'")
            estudiantes = self._consultar_usuarios_por_filtro(filtro_original, headers)
            
            # 2. Intento Plural/Singular si falló
            if not estudiantes:
                alternativa = ""
                if self.departamento_filtro.lower().startswith("estudiante "):
                    alternativa = self.departamento_filtro.replace("Estudiante ", "Estudiantes ")
                elif self.departamento_filtro.lower().startswith("estudiantes "):
                    alternativa = self.departamento_filtro.replace("Estudiantes ", "Estudiante ")
                
                if alternativa:
                    print(f"⚠️ No encontrados. Probando alternativa: '{alternativa}'")
                    filtro_alt = f"$filter=department eq '{alternativa}'"
                    estudiantes = self._consultar_usuarios_por_filtro(filtro_alt, headers)
                    if estudiantes:
                        print(f"✅ ¡Encontrados con '{alternativa}'!")
                        self.departamento_filtro = alternativa # Actualizamos para log
            
            self.resultados["total_estudiantes_tenant"] = len(estudiantes)
            print(f"✅ {len(estudiantes)} estudiantes totales encontrados")
            return estudiantes
        except Exception as e:
            self.resultados["errores"].append(f"Error obteniendo estudiantes: {str(e)}")
            return []

    def agregar_a_team(self, team_id: str, user_id: str) -> tuple:
        """Agrega al usuario como Member al Team"""
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        url = f"{config.GRAPH_ENDPOINT}/groups/{team_id}/members/$ref"
        body = {"@odata.id": f"https://graph.microsoft.com/v1.0/directoryObjects/{user_id}"}
        
        try:
            response = requests.post(url, json=body, headers=headers, verify=False, timeout=10)
            if response.status_code == 204:
                return True, "Vinculado"
            elif response.status_code == 400:
                error = response.json().get('error', {})
                if "already exists" in error.get('message', '').lower():
                    return True, "Ya en equipo"
                return False, f"Error: {error.get('message')}"
            else:
                return False, f"Status {response.status_code}"
        except Exception as e:
            return False, str(e)

    def ejecutar(self):
        """Proceso principal - GENERADOR para progreso en tiempo real"""
        yield {"status": "info", "message": "Iniciando sincronización automática...", "progress": 0}
        
        try:
            if not self.token and not self.obtener_token():
                raise Exception("No se pudo obtener token")
            
            yield {"status": "info", "message": "Obteniendo equipos del Tenant...", "progress": 10}
            if not self.obtener_todos_los_teams():
                raise Exception("No se pudieron obtener equipos de Teams")
            
            yield {"status": "info", "message": "Buscando estudiantes en el Tenant...", "progress": 20}
            estudiantes = self.obtener_estudiantes_del_tenant()
            if not estudiantes:
                yield {"status": "warning", "message": "No se encontraron estudiantes en el departamento especificado."}
                yield {"status": "complete", "results": self.resultados, "progress": 100}
                return

            # Organizar equipos por el código del curso
            equipos_por_curso = {}
            for team in self.teams_encontrados:
                curso = team['displayName'].split(":")[0].strip()
                if curso not in equipos_por_curso:
                    equipos_por_curso[curso] = []
                equipos_por_curso[curso].append(team)
            
            total = len(estudiantes)
            yield {"status": "process", "message": f"Sincronizando {total} estudiantes...", "progress": 25}
            
            # Procesar cada estudiante
            for idx, est in enumerate(estudiantes):
                upn = est.get('userPrincipalName')
                u_id = est.get('id')
                nombre = est.get('displayName')
                curso_est = str(est.get('jobTitle', '')).strip()
                
                percent = int(25 + (idx / total) * 70)
                if idx % 10 == 0: # Evitar saturar el canal SSE con logs de cada estudiante
                    yield {"status": "process", "message": f"[{idx+1}/{total}] Sincronizando: {nombre}", "progress": percent}
                
                if not curso_est:
                    self.resultados["estudiantes_sin_curso"] += 1
                    continue
                
                if curso_est in equipos_por_curso:
                    for team in equipos_por_curso[curso_est]:
                        team_id = team['id']
                        team_name = team['displayName']
                        
                        exito, msg = self.agregar_a_team(team_id, u_id)
                        
                        if exito:
                            if msg == "Ya en equipo":
                                self.resultados["estudiantes_ya_en_equipo"] += 1
                            else:
                                self.resultados["estudiantes_vinculados"] += 1
                                yield {"status": "log", "message": f"   ✅ {nombre} -> {team_name}"}
                        else:
                            self.resultados["errores_vinculacion"] += 1
                            yield {"status": "log", "message": f"   ❌ Error con {nombre} en {team_name}"}
                        
                        self.resultados["estudiantes_procesados"].append({
                            "Estudiante": upn, "Nombre": nombre, "Curso": curso_est, "Equipo": team_name, "Resultado": msg
                        })

            yield {"status": "info", "message": "Guardando reporte final...", "progress": 98}
            self.guardar_logs()
            yield {"status": "complete", "message": "Sincronización finalizada", "progress": 100, "results": self.resultados}
            
        except Exception as e:
            yield {"status": "error", "message": str(e)}
            self.resultados["errores"].append(str(e))
            yield {"status": "complete", "results": self.resultados}

    def guardar_logs(self):
        """Guarda los resultados en un archivo de log"""
        try:
            os.makedirs(config.CARPETA_LOGS, exist_ok=True)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            log_file = os.path.join(config.CARPETA_LOGS, f'sincronizacion_automatica_{timestamp}.log')
            
            with open(log_file, 'w', encoding='utf-8') as f:
                f.write("REPORTE DE SINCRONIZACIÓN AUTOMÁTICA DE TEAMS\n")
                f.write(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Departamento Filtro: {self.departamento_filtro}\n")
                f.write("="*70 + "\n\n")
                f.write(f"Total Estudiantes encontrados en Tenant: {self.resultados['total_estudiantes_tenant']}\n")
                f.write(f"Total Equipos Identificados: {self.resultados['total_equipos']}\n")
                f.write(f"Estudiantes sin JobTitle (Curso): {self.resultados['estudiantes_sin_curso']}\n")
                f.write(f"Vínculos Exitantes/Nuevos realizados: {self.resultados['estudiantes_vinculados']}\n")
                f.write(f"Ya eran miembros: {self.resultados['estudiantes_ya_en_equipo']}\n")
                f.write(f"Errores de vinculación: {self.resultados['errores_vinculacion']}\n\n")
                
                if self.resultados["errores"]:
                    f.write("ERRORES CRÍTICOS:\n")
                    for err in self.resultados["errores"]:
                        f.write(f"!!! {err}\n")
            
            print(f"📝 Log guardado en: {log_file}")
        except Exception as e:
            print(f"⚠️ No se pudo guardar el log: {e}")

if __name__ == "__main__":
    # Prueba
    sinc = SincronizadorAutomaticoTeams()
    # sinc.ejecutar()
