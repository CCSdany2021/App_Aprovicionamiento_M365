import pandas as pd
import requests
import urllib3
from datetime import datetime
import os
import sys

# Añadir la carpeta scripts al path para importar configuración
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class VaciadorEquipos:
    """Clase para vaciar Equipos (Teams): Eliminar miembros y owners (excepto CAP)"""

    CUENTA_CAP = "cap@calasanzsuba.edu.co"

    def __init__(self):
        config.validar_configuracion()
        self.token = None
        self.resultados = {
            "total": 0,
            "total_equipos": 0,
            "equipos_procesados": 0,
            "miembros_eliminados": 0,
            "owners_eliminados": 0,
            "errores": 0,
            "detalles": [],
            "log_detallado": []  # Lista para registro detallado
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
            self.token = response.json()["access_token"]
            return True
        except requests.RequestException as e:
            self.resultados["detalles"].append(f"Error obteniendo token: {e}")
            return False

    def obtener_id_equipo(self, identificador: str) -> str:
        """Busca el ID de un equipo por Email o ID"""
        if not self.token:
            return None

        # Si parece un ID (GUID), devolverlo directamente
        if len(identificador) == 36 and '-' in identificador:
             return identificador

        # Buscar por mail
        headers = {"Authorization": f"Bearer {self.token}"}
        url = f"{config.GRAPH_ENDPOINT}/groups?$filter=mail eq '{identificador}'&$select=id"

        try:
            response = requests.get(url, headers=headers, verify=False)
            response.raise_for_status()
            data = response.json()
            if data['value']:
                return data['value'][0]['id']
            return None
        except Exception:
            return None

    def obtener_usuarios_grupo(self, group_id: str, rol: str = 'members') -> list:
        """Obtiene miembros u owners de un grupo"""
        usuarios = []
        # Endpoint para owners es /owners, para miembros es /members
        endpoint = "owners" if rol == 'owners' else "members"
        url = f"{config.GRAPH_ENDPOINT}/groups/{group_id}/{endpoint}?$select=id,userPrincipalName,mail"
        headers = {"Authorization": f"Bearer {self.token}"}

        while url:
            try:
                response = requests.get(url, headers=headers, verify=False)
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
        """Elimina un usuario del grupo (endpoint cambia si es owner)"""
        endpoint = "owners" if es_owner else "members"
        url = f"{config.GRAPH_ENDPOINT}/groups/{group_id}/{endpoint}/{user_id}/$ref"
        headers = {"Authorization": f"Bearer {self.token}"}
        
        try:
            response = requests.delete(url, headers=headers, verify=False)
            if response.status_code == 204:
                return True, ""
            else:
                return False, f"Status: {response.status_code}, Body: {response.text}"
        except Exception as e:
            return False, str(e)

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
            
            group_id = self.obtener_id_equipo(ident)
            if not group_id:
                msg = f"Equipo no encontrado: {ident}"
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
                f.write(f"Errores: {self.resultados['errores']}\n\n")
                
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
