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
    
    def __init__(self, departamento_filtro=None, target_upn=None):
        try:
            config.validar_configuracion()
        except:
            pass
        
        self.token = None
        # Obtener periodo de la configuración dinámica
        default_dept = f"Estudiantes {config.PERIODO_ACTUAL}"
        self.departamento_filtro = departamento_filtro if departamento_filtro else default_dept
        self.target_upn = target_upn
        
        mode_msg = f"UPN='{self.target_upn}'" if self.target_upn else f"Departamento='{self.departamento_filtro}'"
        print(f"DEBUG: Sincronizador inicializado con {mode_msg}")
        
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
            self.token_expires_at = time.time() + 3500 # 58 minutos aprox de validez segura
            return True
        except Exception as e:
            self.resultados["errores"].append(f"Error de token: {str(e)}")
            return False

    def _ensure_valid_token(self):
        """Verifica si el token va a expirar y lo renueva"""
        if not self.token or time.time() > self.token_expires_at:
            print("🔄 Renovando token de acceso...")
            self.obtener_token()

    def _request_with_retry(self, method, url, **kwargs):
        """Wrapper para requests con manejo de Rate Limit (429) y Re-Auth (401)"""
        self._ensure_valid_token()
        headers = kwargs.get('headers', {})
        headers['Authorization'] = f"Bearer {self.token}"
        kwargs['headers'] = headers
        
        intentos = 0
        max_intentos = 5
        wait_time = 2
        
        while intentos < max_intentos:
            try:
                if method == 'POST':
                    response = requests.post(url, **kwargs)
                elif method == 'GET':
                    response = requests.get(url, **kwargs)
                else:
                    return None
                
                # Manejo de expiración de token en vuelo
                if response.status_code == 401:
                    print("⚠️ Token expiró (401). Renovando y reintentando...")
                    self.obtener_token()
                    headers['Authorization'] = f"Bearer {self.token}" # Actualizar header
                    kwargs['headers'] = headers
                    intentos += 1
                    continue

                # Manejo de Throttling
                if response.status_code == 429 or response.status_code == 503:
                    retry_after = int(response.headers.get('Retry-After', wait_time))
                    print(f"⏳ Throttling (429/503). Esperando {retry_after}s...")
                    time.sleep(retry_after)
                    wait_time *= 2
                    intentos += 1
                    continue
                    
                return response
                
            except requests.exceptions.RequestException as e:
                print(f"⚠️ Error de red: {e}. Reintentando...")
                time.sleep(wait_time)
                intentos += 1
                wait_time *= 2
        
        return None

    def obtener_todos_los_teams(self) -> bool:
        """Obtiene todos los equipos que siguen el patrón 'CURSO: MATERIA'"""
        if not self.token:
            return False
        
        try:
            teams = []
            url = f"{config.GRAPH_ENDPOINT}/groups?$select=id,displayName&$top=999"
            
            # Usamos el wrapper para listar equipos también
            while url:
                response = self._request_with_retry('GET', url, verify=False, timeout=15)
                if response and response.status_code == 200:
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

    def _consultar_usuarios_por_filtro(self, filtro):
        """Helper para consultar usuarios paginados"""
        estudiantes = []
        url = f"{config.GRAPH_ENDPOINT}/users?{filtro}&$select=id,userPrincipalName,jobTitle,displayName&$top=999"
        
        while url:
            response = self._request_with_retry('GET', url, verify=False, timeout=20)
            if response and response.status_code == 200:
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
            
            # 0. Si hay target_upn, buscar solo ese usuario
            if self.target_upn:
                print(f"🔍 Buscando usuario específico: '{self.target_upn}'")
                filtro_upn = f"$filter=userPrincipalName eq '{self.target_upn}'"
                estudiantes = self._consultar_usuarios_por_filtro(filtro_upn)
                if not estudiantes:
                   print(f"⚠️ Usuario '{self.target_upn}' no encontrado.")
            else:
                # 1. Intento original
                filtro_original = f"$filter=department eq '{self.departamento_filtro}'"
                print(f"🔍 Buscando estudiantes: '{self.departamento_filtro}'")
                estudiantes = self._consultar_usuarios_por_filtro(filtro_original)
                
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
                        estudiantes = self._consultar_usuarios_por_filtro(filtro_alt)
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
        url = f"{config.GRAPH_ENDPOINT}/groups/{team_id}/members/$ref"
        body = {"@odata.id": f"https://graph.microsoft.com/v1.0/directoryObjects/{user_id}"}
        
        try:
            # Usar el wrapper robusto
            response = self._request_with_retry('POST', url, json=body, verify=False, timeout=10)
            
            if response is None:
                return False, "Error de Red/Timeout"

            if response.status_code == 204:
                return True, "Vinculado"
            elif response.status_code == 400:
                error = response.json().get('error', {})
                msg_lower = error.get('message', '').lower()
                if "already exists" in msg_lower or "one or more added object references already exist" in msg_lower:
                    return True, "Ya en equipo"
                return False, f"Error 400: {error.get('message')}"
            else:
                try:
                    err_msg = response.json().get('error', {}).get('message', response.text)
                except:
                    err_msg = response.text
                return False, f"Status {response.status_code}: {err_msg}"
                
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
            
            # Filtrar estudiantes que NO tienen curso (JobTitle) para que no salgan en el log ni en el total
            estudiantes = [e for e in estudiantes if e.get('jobTitle')]
            
            total = len(estudiantes)
            self.resultados["total_filtrado"] = total
            yield {"status": "process", "message": f"Sincronizando {total} estudiantes válidos (con curso)...", "progress": 25}
            
            # Procesar cada estudiante
            for idx, est in enumerate(estudiantes):
                upn = est.get('userPrincipalName')
                u_id = est.get('id')
                nombre = est.get('displayName')
                curso_est = str(est.get('jobTitle', '')).strip()
                
                percent = int(25 + (idx / total) * 70)
                if idx % 10 == 0: # Evitar saturar el canal SSE con logs de cada estudiante
                    yield {"status": "process", "message": f"[{idx+1}/{total}] Sincronizando: {nombre} ({upn})", "progress": percent}
                
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
                                # Opcional: No llenar el log visual si ya existe, o usar nivel info
                            else:
                                self.resultados["estudiantes_vinculados"] += 1
                                yield {"status": "log", "message": f"   ✅ {nombre} -> {team_name}"}
                        else:
                            self.resultados["errores_vinculacion"] += 1
                            # CRITICAL FIX: Show the actual error message 'msg'
                            yield {"status": "log", "message": f"   ❌ Error con {nombre} en {team_name}: {msg}"}
                        
                        self.resultados["estudiantes_procesados"].append({
                            "Estudiante": upn, "Nombre": nombre, "Curso": curso_est, "Equipo": team_name, "Resultado": msg, "Exito": exito
                        })

            yield {"status": "info", "message": "Guardando reporte final...", "progress": 98}
            self.guardar_logs()
            yield {"status": "complete", "message": "Sincronización finalizada", "progress": 100, "results": self.resultados}
            
        except Exception as e:
            import traceback
            trace_str = traceback.format_exc()
            yield {"status": "error", "message": f"CRITICAL ERROR: {str(e)}"}
            self.resultados["errores"].append(f"{str(e)} | {trace_str}")
            self.guardar_logs() # Intentar guardar lo que se pueda
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
                f.write(f"Vínculos Exitosos: {self.resultados['estudiantes_vinculados']}\n")
                f.write(f"Ya eran miembros: {self.resultados['estudiantes_ya_en_equipo']}\n")
                f.write(f"Errores de vinculación: {self.resultados['errores_vinculacion']}\n\n")
                
                if self.resultados["errores"]:
                    f.write("ERRORES DE SISTEMA:\n")
                    for err in self.resultados["errores"]:
                        f.write(f"!!! {err}\n")
                    f.write("\n")
                
                f.write("DETALLE DE OPERACIONES (Muestra):\n")
                f.write("Estudiante | Equipo | Resultado\n")
                for item in self.resultados["estudiantes_procesados"]:
                    # Escribir solo errores o cambios efectivos para no hacer logs de 1GB si son muchos
                    if not item['Exito'] or item['Resultado'] == 'Vinculado': 
                         f.write(f"{item['Estudiante']} | {item['Equipo']} | {item['Resultado']}\n")
            
            print(f"📝 Log guardado en: {log_file}")
        except Exception as e:
            print(f"⚠️ No se pudo guardar el log: {e}")

if __name__ == "__main__":
    # Prueba
    sinc = SincronizadorAutomaticoTeams()
    # sinc.ejecutar()
