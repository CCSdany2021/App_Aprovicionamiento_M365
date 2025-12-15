import pandas as pd
import requests
import urllib3
from datetime import datetime
import os
import sys
import time

# Añadir la carpeta scripts al path para importar configuración
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class DesvinculadorGrupos:
    """Clase para desvincular miembros de grupos de distribución/seguridad"""

    def __init__(self):
        config.validar_configuracion()
        self.token = None
        self.resultados = {
            "total": 0,
            "total_grupos": 0,
            "grupos_procesados": 0,
            "miembros_eliminados": 0,
            "errores": 0,
            "detalles": []
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

    def obtener_id_grupo(self, email_grupo: str) -> str:
        """Busca el ID de un grupo por su correo electrónico"""
        if not self.token:
            return None

        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        # Filtrar por mail o proxyAddresses
        url = f"{config.GRAPH_ENDPOINT}/groups?$filter=mail eq '{email_grupo}' or proxyAddresses/any(x:x eq 'smtp:{email_grupo}')&$select=id,displayName"

        try:
            response = requests.get(url, headers=headers, verify=False)
            response.raise_for_status()
            data = response.json()
            
            if data['value']:
                return data['value'][0]['id']
            return None
        except Exception as e:
            print(f"Error buscando grupo {email_grupo}: {e}")
            return None

    def obtener_miembros_grupo(self, group_id: str) -> list:
        """Obtiene todos los miembros de un grupo"""
        miembros = []
        url = f"{config.GRAPH_ENDPOINT}/groups/{group_id}/members?$select=id,userPrincipalName,displayName"
        headers = {"Authorization": f"Bearer {self.token}"}

        while url:
            try:
                response = requests.get(url, headers=headers, verify=False)
                if response.status_code == 200:
                    data = response.json()
                    miembros.extend(data.get('value', []))
                    url = data.get('@odata.nextLink') # Paginación
                else:
                    break
            except Exception:
                break
        
        return miembros

    def eliminar_miembro(self, group_id: str, member_id: str) -> tuple[bool, str]:
        """Elimina un miembro del grupo"""
        url = f"{config.GRAPH_ENDPOINT}/groups/{group_id}/members/{member_id}/$ref"
        headers = {"Authorization": f"Bearer {self.token}"}
        
        try:
            response = requests.delete(url, headers=headers, verify=False)
            if response.status_code == 204:
                return True, ""
            else:
                return False, f"Status: {response.status_code}, Body: {response.text}"
        except Exception as e:
            return False, str(e)

    def procesar_desvinculacion(self, ruta_archivo: str, confirmacion: bool = False) -> dict:
        """Proceso principal de desvinculación"""
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
            
            # Buscar columna de email
            columna_email = next((col for col in df.columns if col.lower() in ['primarysmtpaddress', 'email', 'correo']), None)
            
            if not columna_email:
                self.resultados["errores"] += 1
                self.resultados["detalles"].append("No se encontró columna 'PrimarySmtpAddress' o equivalente")
                return self.resultados

            grupos = df[columna_email].dropna().unique().tolist()
            self.resultados["total_grupos"] = len(grupos)
            self.resultados["total"] = len(grupos)

        except Exception as e:
            self.resultados["errores"] += 1
            self.resultados["detalles"].append(f"Error leyendo archivo: {e}")
            return self.resultados

        print(f"🔄 Iniciando desvinculación para {len(grupos)} grupos...")
        
        for email in grupos:
            email = email.strip()
            print(f"🔍 Procesando grupo: {email}")
            
            group_id = self.obtener_id_grupo(email)
            
            if not group_id:
                msg = f"Grupo no encontrado en Azure AD: {email}"
                print(f"❌ {msg}")
                self.resultados["detalles"].append(msg)
                self.resultados["errores"] += 1
                continue
                
            # Obtener miembros
            miembros = self.obtener_miembros_grupo(group_id)
            print(f"   👥 Encontrados {len(miembros)} miembros")
            
            count_removed = 0
            for miembro in miembros:
                member_id = miembro['id']
                member_upn = miembro.get('userPrincipalName', 'Unknown')
                
                exito, error_msg = self.eliminar_miembro(group_id, member_id)
                if exito:
                    count_removed += 1
                    # print(f"      - Desvinculado: {member_upn}") # Verbose
                else:
                    self.resultados["detalles"].append(f"Error desvinculando {member_upn} de {email}: {error_msg}")
            
            self.resultados["miembros_eliminados"] += count_removed
            self.resultados["grupos_procesados"] += 1
            print(f"   ✅ Desvinculados {count_removed} miembros")

        self.guardar_log()
        return self.resultados

    def guardar_log(self):
        """Guarda log del proceso"""
        try:
            os.makedirs(config.CARPETA_LOGS, exist_ok=True)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            log_file = os.path.join(config.CARPETA_LOGS, f'desvinculacion_grupos_{timestamp}.log')
            
            with open(log_file, 'w', encoding='utf-8') as f:
                f.write(f"DESVINCULACIÓN DE MIEMBROS DE GRUPOS\n")
                f.write(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("="*50 + "\n")
                f.write(f"Total Grupos: {self.resultados['total_grupos']}\n")
                f.write(f"Grupos Procesados: {self.resultados['grupos_procesados']}\n")
                f.write(f"Total Miembros Desvinculados: {self.resultados['miembros_eliminados']}\n")
                f.write(f"Errores Generales: {self.resultados['errores']}\n\n")
                f.write("DETALLES:\n")
                for detalle in self.resultados["detalles"]:
                    f.write(f"- {detalle}\n")
                    
        except Exception as e:
            print(f"Error guardando log: {e}")
