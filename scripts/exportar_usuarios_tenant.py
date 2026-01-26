
import requests
import json
import pandas as pd
import os
from datetime import datetime
import sys

# Ajustar path para importar configuración
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config

class ExportadorUsuariosTenant:
    def __init__(self):
        config.validar_configuracion()
        self.token = None
        self.output_folder = config.CARPETA_RESULTADOS
        
    def obtener_token(self):
        url = f"https://login.microsoftonline.com/{config.TENANT_ID}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": config.CLIENT_ID,
            "client_secret": config.CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default"
        }
        try:
            resp = requests.post(url, data=data)
            resp.raise_for_status()
            self.token = resp.json()["access_token"]
            return True
        except Exception as e:
            print(f"Error obteniendo token: {e}")
            return False

    def exportar_usuarios(self, departamento_filtro=None):
        """
        Exporta usuarios del tenant.
        Si departamento_filtro está presente, filtra por ese departamento.
        Si es None o vacío, exporta TODOS los usuarios.
        """
        mensajes_inicio = f"Iniciando exportación de usuarios..."
        if departamento_filtro:
            mensajes_inicio = f"Iniciando exportación de usuarios del departamento: '{departamento_filtro}'..."
            
        yield {"status": "info", "message": mensajes_inicio, "progress": 5}
        
        if not self.obtener_token():
            yield {"status": "error", "message": "No se pudo autenticar con Microsoft Graph"}
            return

        headers = {"Authorization": f"Bearer {self.token}"}
        
        # Propiedades solicitadas
        select_fields = "id,displayName,userPrincipalName,jobTitle,department,officeLocation,mail,proxyAddresses,accountEnabled"
        
        url = f"{config.GRAPH_ENDPOINT}/users?$select={select_fields}&$top=999"
        
        if departamento_filtro:
            # Nota: codificación URL básica para espacios
            clean_dept = departamento_filtro.replace("'", "''") # Escape single quotes
            url += f"&$filter=department eq '{clean_dept}'"
            print(f"Buscando usuarios con department eq '{clean_dept}'")
        else:
            print("Exportando TODOS los usuarios (sin filtro de departamento)")
        
        usuarios_encontrados = []
        
        yield {"status": "info", "message": "Consultando Microsoft Graph API...", "progress": 10}
        
        total_paginas = 0
        while url:
            try:
                resp = requests.get(url, headers=headers)
                if resp.status_code != 200:
                    yield {"status": "error", "message": f"Error API: {resp.text}"}
                    break
                    
                data = resp.json()
                batch = data.get('value', [])
                usuarios_encontrados.extend(batch)
                
                total_paginas += 1
                yield {"status": "info", "message": f"Recuperados {len(usuarios_encontrados)} usuarios...", "progress": 20 + (total_paginas % 60)}
                
                # Paginación
                url = data.get('@odata.nextLink')
                
            except Exception as e:
                yield {"status": "error", "message": f"Error de conexión: {e}"}
                return

        if not usuarios_encontrados:
            yield {"status": "info", "message": "No se encontraron usuarios con los criterios especificados.", "progress": 100}
            return

        yield {"status": "info", "message": f"Procesando {len(usuarios_encontrados)} registros...", "progress": 85}

        # Procesar datos (Email Alias, Renombrar columnas)
        datos_procesados = []
        for user in usuarios_encontrados:
            # Lógica para Alias
            email_alias = ""
            mail = user.get('mail')
            proxy_addresses = user.get('proxyAddresses', [])
            
            if mail:
                email_alias = mail.split("@")[0]
            elif proxy_addresses:
                # Buscar SMTP: principal
                smtp = next((addr for addr in proxy_addresses if addr.startswith("SMTP:")), None)
                if smtp:
                    email_alias = smtp.replace("SMTP:", "").split("@")[0]
                else:
                    # Buscar smtp: secundario
                    smtp = next((addr for addr in proxy_addresses if addr.startswith("smtp:")), None)
                    if smtp:
                        email_alias = smtp.replace("smtp:", "").split("@")[0]

            item = {
                "Usuario CCS": user.get('userPrincipalName', ''),
                "Nombre Completo": user.get('displayName', ''),
                "Puesto": user.get('jobTitle', ''),
                "Departamento": user.get('department', ''),
                "Ubicacion Office": user.get('officeLocation', ''),
                "Correo": mail if mail else '',
                "Alias de Correo": email_alias,
                "Estado Cuenta": "Activa" if user.get('accountEnabled') else "Inactiva"
            }
            datos_procesados.append(item)

        # Convertir a DataFrame
        df = pd.DataFrame(datos_procesados)
        
        # Generar nombre de archivo
        fecha = datetime.now().strftime("%Y%m%d_%H%M%S")
        prefix = "Usuarios_Tenant"
        if departamento_filtro:
            sanitized_dept = departamento_filtro.replace(" ", "_").replace("/", "-")
            prefix = f"Exportacion_{sanitized_dept}"
        
        filename = f"{prefix}_{fecha}.xlsx"
        filepath = os.path.join(self.output_folder, filename)
        
        os.makedirs(self.output_folder, exist_ok=True)
        
        # Guardar Excel
        try:
            # Intento robusto de guardar con o sin formato bonito
            engine = 'openpyxl' # Default fallback
            try:
                import xlsxwriter
                engine = 'xlsxwriter'
            except ImportError:
                yield {"status": "log", "message": "⚠️ Librería 'xlsxwriter' no encontrada. Usando formato simple."}

            if engine == 'xlsxwriter':
                with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Usuarios')
                    # Autoajuste de columnas simple
                    worksheet = writer.sheets['Usuarios']
                    for i, col in enumerate(df.columns):
                        # Evitar error si hay valores nulos o extraños
                        try:
                            max_len = df[col].astype(str).map(len).max()
                            width = max(max_len, len(col)) + 2
                            # Limitar ancho razonable
                            width = min(width, 50) 
                            worksheet.set_column(i, i, width)
                        except:
                            pass
            else:
                 # Fallback standard
                 df.to_excel(filepath, index=False)
                    
            yield {"status": "log", "message": f"✅ Archivo generado: {filename}"}
        except Exception as e:
             yield {"status": "error", "message": f"Error guardando Excel: {e}"}
             return

        yield {
            "status": "complete", 
            "message": "Exportación finalizada exitosamente.", 
            "progress": 100, 
            "results": {"archivo": filename, "total": len(df)}
        }
