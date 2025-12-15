import pandas as pd
import requests
import urllib3
from datetime import datetime, timedelta
import os
import sys
import time

# Añadir la carpeta scripts al path para importar configuración
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class VaciadorEquipos:
    """Clase para vaciar Equipos (Teams): Eliminar miembros y owners (excepto CAP)"""

    CUENTA_CAP = "cap@calasanzsuba.edu.co"
    MAX_REINTENTOS = 3  # Número máximo de reintentos por operación

    def __init__(self):
        config.validar_configuracion()
        self.token = None
        self.token_expiracion = None  # Timestamp de expiración del token
        self.token_renovaciones = 0  # Contador de renovaciones
        self.resultados = {
            "total": 0,
            "total_equipos": 0,
            "equipos_procesados": 0,
            "miembros_eliminados": 0,
            "owners_eliminados": 0,
            "errores": 0,
            "detalles": [],
            "log_detallado": [],  # Lista para registro detallado
            "token_renovaciones": 0  # Contador de renovaciones de token
        }

    def obtener_token(self) -> bool:
        """Obtiene token de acceso a Microsoft Graph API"""
        url = f"https://login.microsoftonline.com/{config.TENANT_ID}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": config.CLIENT_ID,
            "client_secret": config.CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default"
        }

        try:
            response = requests.post(url, data=data, verify=False)
            response.raise_for_status()
            token_data = response.json()
            self.token = token_data["access_token"]
            
            # Calcular tiempo de expiración (normalmente 3600 segundos = 1 hora)
            # Renovamos 5 minutos antes para estar seguros
            expires_in = token_data.get("expires_in", 3600)
            self.token_expiracion = datetime.now() + timedelta(seconds=expires_in - 300)
            
            if self.token_renovaciones > 0:
                msg = f"✓ Token renovado exitosamente (renovación #{self.token_renovaciones})"
                print(msg)
                self.resultados["log_detallado"].append(msg)
            
            return True
        except requests.RequestException as e:
            self.resultados["detalles"].append(f"Error obteniendo token: {e}")
            return False

    def token_valido(self) -> bool:
        """Verifica si el token actual sigue siendo válido"""
        if not self.token or not self.token_expiracion:
            return False
        return datetime.now() < self.token_expiracion

    def renovar_token_si_necesario(self) -> bool:
        """Renueva el token si está próximo a expirar o ya expiró"""
        if not self.token_valido():
            print("⚠️ Token expirado o próximo a expirar, renovando...")
            self.token_renovaciones += 1
            self.resultados["token_renovaciones"] = self.token_renovaciones
            return self.obtener_token()
        return True

    def obtener_id_equipo(self, identificador: str) -> tuple[str, str]:
        """Busca el ID de un equipo por Email o ID
        
        Returns:
            tuple: (team_id, mensaje_error) - team_id es None si hay error
        """
        if not self.token:
            return None, "Token no disponible"

        identificador = identificador.strip()
        
        # Si parece un ID (GUID), validarlo y devolverlo directamente
        if len(identificador) == 36 and identificador.count('-') == 4:
            # Validar formato GUID: 8-4-4-4-12 caracteres
            partes = identificador.split('-')
            if len(partes) == 5 and len(partes[0]) == 8 and len(partes[1]) == 4:
                print(f"   ℹ️ Usando Team ID directamente: {identificador[:8]}...")
                return identificador, None

        # Renovar token si es necesario antes de hacer la búsqueda
        if not self.renovar_token_si_necesario():
            return None, "No se pudo renovar el token"

        # Buscar por mail
        headers = {"Authorization": f"Bearer {self.token}"}
        url = f"{config.GRAPH_ENDPOINT}/groups?$filter=mail eq '{identificador}'&$select=id,displayName"

        try:
            response = requests.get(url, headers=headers, verify=False)
            
            # Manejar error 401 específicamente
            if response.status_code == 401:
                print("   ⚠️ Token expirado durante búsqueda, renovando...")
                if self.renovar_token_si_necesario():
                    # Reintentar con nuevo token
                    headers = {"Authorization": f"Bearer {self.token}"}
                    response = requests.get(url, headers=headers, verify=False)
                else:
                    return None, "No se pudo renovar token después de error 401"
            
            response.raise_for_status()
            data = response.json()
            
            if data['value']:
                team_id = data['value'][0]['id']
                display_name = data['value'][0].get('displayName', 'N/A')
                print(f"   ✓ Equipo encontrado: {display_name}")
                return team_id, None
            else:
                return None, f"No se encontró equipo con email: {identificador}"
                
        except requests.RequestException as e:
            return None, f"Error en búsqueda: {str(e)}"
        except Exception as e:
            return None, f"Error inesperado: {str(e)}"

    def obtener_usuarios_grupo(self, group_id: str, rol: str = 'members') -> list:
        """Obtiene miembros u owners de un grupo"""
        # Renovar token si es necesario
        if not self.renovar_token_si_necesario():
            return []
            
        usuarios = []
        # Endpoint para owners es /owners, para miembros es /members
        endpoint = "owners" if rol == 'owners' else "members"
        url = f"{config.GRAPH_ENDPOINT}/groups/{group_id}/{endpoint}?$select=id,userPrincipalName,mail"
        headers = {"Authorization": f"Bearer {self.token}"}

        while url:
            try:
                response = requests.get(url, headers=headers, verify=False)
                
                # Manejar token expirado
                if response.status_code == 401:
                    if self.renovar_token_si_necesario():
                        headers = {"Authorization": f"Bearer {self.token}"}
                        response = requests.get(url, headers=headers, verify=False)
                    else:
                        break
                
                if response.status_code == 200:
                    data = response.json()
                    usuarios.extend(data.get('value', []))
                    url = data.get('@odata.nextLink')
                else:
                    break
            except Exception:
                break
        return usuarios

    def eliminar_miembro(self, group_id: str, user_id: str, es_owner: bool = False) -> tuple[bool, str]:
        """Elimina un usuario del grupo (endpoint cambia si es owner)
        
        Incluye lógica de reintento automático en caso de error 401
        """
        endpoint = "owners" if es_owner else "members"
        url = f"{config.GRAPH_ENDPOINT}/groups/{group_id}/{endpoint}/{user_id}/$ref"
        
        for intento in range(self.MAX_REINTENTOS):
            # Renovar token si es necesario antes de cada intento
            if not self.renovar_token_si_necesario():
                return False, "No se pudo renovar el token"
            
            headers = {"Authorization": f"Bearer {self.token}"}
            
            try:
                response = requests.delete(url, headers=headers, verify=False)
                
                if response.status_code == 204:
                    return True, ""
                elif response.status_code == 401:
                    # Token expirado, forzar renovación
                    print(f"   ⚠️ Error 401 en intento {intento + 1}/{self.MAX_REINTENTOS}, renovando token...")
                    self.token = None  # Forzar renovación
                    self.token_expiracion = None
                    if intento < self.MAX_REINTENTOS - 1:
                        time.sleep(1)  # Pequeña pausa antes de reintentar
                        continue
                    else:
                        return False, f"Error 401 persistente después de {self.MAX_REINTENTOS} intentos"
                elif response.status_code == 404:
                    # Usuario ya no existe o no es miembro
                    return True, "Usuario no encontrado (posiblemente ya eliminado)"
                else:
                    return False, f"Status: {response.status_code}, Body: {response.text}"
                    
            except Exception as e:
                if intento < self.MAX_REINTENTOS - 1:
                    time.sleep(1)
                    continue
                return False, str(e)
        
        return False, f"Falló después de {self.MAX_REINTENTOS} intentos"

    def procesar(self, ruta_archivo: str, confirmacion: bool = False) -> dict:
        """Proceso principal"""
        if not self.obtener_token():
            return self.resultados

        # Cargar archivo
        try:
            if ruta_archivo.endswith(".xlsx"):
                df = pd.read_excel(ruta_archivo, dtype=str)
            elif ruta_archivo.endswith(".csv"):
                df = pd.read_csv(ruta_archivo, dtype=str, encoding="utf-8")
            else:
                raise ValueError("Formato no soportado")
            
            # Buscar columna de identificador
            col = next((c for c in df.columns if c.lower() in ['groupid', 'teamid', 'id', 'primarysmtpaddress', 'email', 'correo']), None)
            
            if not col:
                self.resultados["errores"] += 1
                self.resultados["detalles"].append("No se encontró columna ID o Email")
                return self.resultados

            equipos = df[col].dropna().unique().tolist()
            self.resultados["total_equipos"] = len(equipos)
            self.resultados["total"] = len(equipos)

        except Exception as e:
            self.resultados["errores"] += 1
            self.resultados["detalles"].append(f"Error leyendo archivo: {e}")
            return self.resultados
        
        print(f"🔄 Iniciando vaciado de para {len(equipos)} equipos...")


        for ident in equipos:
            ident = ident.strip()
            print(f"🔍 Procesando: {ident}")
            
            group_id, error_msg = self.obtener_id_equipo(ident)
            if not group_id:
                msg = error_msg if error_msg else f"Equipo no encontrado: {ident}"
                print(f"❌ {msg}")
                self.resultados["detalles"].append(msg)
                self.resultados["errores"] += 1
                continue


            # 1. Eliminar Miembros (Estudiantes)
            miembros = self.obtener_usuarios_grupo(group_id, 'members')
            for m in miembros:
                uid = m['id']
                upn = m.get('userPrincipalName', 'unknown')
                
                # Opcional: ignorar CAP si está como miembro (aunque no debería importar)
                if upn.lower() == self.CUENTA_CAP:
                    continue

                ok, err = self.eliminar_miembro(group_id, uid, es_owner=False)
                if ok:
                    self.resultados["miembros_eliminados"] += 1
                    self.resultados["log_detallado"].append(f"✓ Miembro eliminado: {upn} del equipo {ident}")
                else:
                    self.resultados["detalles"].append(f"Error borrando miembro {upn} de {ident}: {err}")

            # 2. Eliminar Owners (Docentes) EXCEPTO CAP
            owners = self.obtener_usuarios_grupo(group_id, 'owners')
            for o in owners:
                uid = o['id']
                upn = o.get('userPrincipalName', 'unknown')
                mail = o.get('mail', 'unknown')

                # VALIDACIÓN CRÍTICA: NO BORRAR A CAP
                if upn.lower() == self.CUENTA_CAP or mail.lower() == self.CUENTA_CAP:
                    print(f"   🛡️ Se protege al owner CAP: {upn}")
                    continue
                
                ok, err = self.eliminar_miembro(group_id, uid, es_owner=True)
                if ok:
                    self.resultados["owners_eliminados"] += 1
                    self.resultados["log_detallado"].append(f"✓ Owner eliminado: {upn} del equipo {ident}")
                else:
                    self.resultados["detalles"].append(f"Error borrando owner {upn} de {ident}: {err}")

            self.resultados["equipos_procesados"] += 1
            print(f"   ✅ Equipo procesado.")

        self.guardar_log()
        return self.resultados

    def generar_inventario(self, carpeta_salida: str) -> str:
        """Genera un Excel con TODOS los equipos del tenant"""
        if not self.obtener_token():
            return None

        print("🔍 Buscando todos los equipos en el tenant...")
        
        # Filtro para obtener solo TEAMS
        url = f"{config.GRAPH_ENDPOINT}/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName,mail,description,visibility"
        headers = {"Authorization": f"Bearer {self.token}"}
        
        todos_equipos = []
        
        while url:
            try:
                response = requests.get(url, headers=headers, verify=False)
                if response.status_code == 200:
                    data = response.json()
                    todos_equipos.extend(data.get('value', []))
                    url = data.get('@odata.nextLink')
                    print(f"   ... {len(todos_equipos)} equipos encontrados")
                else:
                    print(f"Error obteniendo equipos: {response.text}")
                    break
            except Exception as e:
                print(f"Excepción buscando equipos: {e}")
                break
        
        if not todos_equipos:
            return None
            
        # Crear DataFrame
        datos = []
        for equipo in todos_equipos:
            datos.append({
                "DisplayName": equipo.get('displayName'),
                "Email": equipo.get('mail'),
                "Id": equipo.get('id'),
                "PrimarySmtpAddress": equipo.get('mail'), # Duplicado útil para compatibilidad
                "Visibility": equipo.get('visibility')
            })
            
        df = pd.DataFrame(datos)
        
        # Guardar
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        nombre_archivo = f"Inventario_Teams_{timestamp}.xlsx"
        ruta_completa = os.path.join(carpeta_salida, nombre_archivo)
        
        df.to_excel(ruta_completa, index=False)
        print(f"✅ Inventario guardado en: {ruta_completa}")
        
        return ruta_completa

    def guardar_log(self):
        """Guarda log del proceso"""
        try:
            os.makedirs(config.CARPETA_LOGS, exist_ok=True)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            log_file = os.path.join(config.CARPETA_LOGS, f'vaciado_equipos_{timestamp}.log')
            
            with open(log_file, 'w', encoding='utf-8') as f:
                f.write(f"VACIADO DE EQUIPOS (TEAMS) - {config.COLEGIO_NOMBRE}\n")
                f.write(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("="*50 + "\n")
                f.write(f"Total Equipos: {self.resultados['total_equipos']}\n")
                f.write(f"Equipos Procesados: {self.resultados['equipos_procesados']}\n")
                f.write(f"Miembros Eliminados: {self.resultados['miembros_eliminados']}\n")
                f.write(f"Owners Eliminados: {self.resultados['owners_eliminados']}\n")
                f.write(f"Errores: {self.resultados['errores']}\n")
                f.write(f"Renovaciones de Token: {self.resultados['token_renovaciones']}\n\n")
                
                # Sección de errores
                if self.resultados["detalles"]:
                    f.write("ERRORES Y ADVERTENCIAS:\n")
                    for detalle in self.resultados["detalles"]:
                        f.write(f"- {detalle}\n")
                    f.write("\n")
                
                # Sección de log detallado
                if self.resultados["log_detallado"]:
                    f.write("REGISTRO DETALLADO DE ELIMINACIONES:\n")
                    f.write("="*50 + "\n")
                    for entrada in self.resultados["log_detallado"]:
                        f.write(f"{entrada}\n")
                    
        except Exception as e:
            print(f"Error guardando log: {e}")
