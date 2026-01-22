import requests
import json
import pandas as pd
import os
from datetime import datetime
import sys

# Ajustar path para importar configuración
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config


class ExportadorEstudiantes:
    def __init__(self):
        config.validar_configuracion()
        self.token = None
        self.output_folder = config.CARPETA_RESULTADOS
        
    def obtener_token(self):
        # Lógica de obtención de token (similar a otras clases o centralizada)
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

    def exportar_por_departamento(self, departamento_filtro="Estudiante 2026"):
        """Exporta usuarios cuyo departamento coincida con el filtro"""
        yield {"status": "info", "message": f"Iniciando exportación de estudiantes del departamento: '{departamento_filtro}'...", "progress": 5}
        
        if not self.obtener_token():
            yield {"status": "error", "message": "No se pudo autenticar con Microsoft Graph"}
            return

        headers = {"Authorization": f"Bearer {self.token}"}
        
        # Usamos filter para traer solo los del departamento especifico
        # OJO: M365 a veces es estricto con los filtros string. 
        # Url encode del departamento es importante si tiene espacios.
        # $filter=department eq 'Estudiante 2026'
        
        print(f"Buscando usuarios con department eq '{departamento_filtro}'")
        
        # Seleccionamos campos clave para validar actualizaciones
        select_fields = "id,displayName,userPrincipalName,givenName,surname,department,city,jobTitle"
        url = f"{config.GRAPH_ENDPOINT}/users?$filter=department eq '{departamento_filtro}'&$select={select_fields}&$top=999"
        
        usuarios_encontrados = []
        
        yield {"status": "info", "message": "Consultando Microsoft Graph API...", "progress": 10}
        
        while url:
            try:
                resp = requests.get(url, headers=headers)
                if resp.status_code != 200:
                    yield {"status": "error", "message": f"Error API: {resp.text}"}
                    break
                    
                data = resp.json()
                batch = data.get('value', [])
                usuarios_encontrados.extend(batch)
                
                yield {"status": "info", "message": f"Encontrados {len(usuarios_encontrados)} estudiantes...", "progress": 50}
                
                # Paginación
                url = data.get('@odata.nextLink')
                
            except Exception as e:
                yield {"status": "error", "message": f"Error de conexión: {e}"}
                return

        if not usuarios_encontrados:
            yield {"status": "info", "message": "No se encontraron estudiantes con ese departamento.", "progress": 100}
            return

        # Convertir a DataFrame
        df = pd.DataFrame(usuarios_encontrados)
        
        # Renombrar columnas para que sea más legible en Excel
        columnas_renombre = {
            "displayName": "Nombre Mostrar",
            "userPrincipalName": "User Principal Name",
            "givenName": "Nombres",
            "surname": "Apellidos",
            "department": "Departamento",
            "city": "Ciudad",
            "jobTitle": "Curso (JobTitle)"
        }
        # Solo renombrar las que existan
        df = df.rename(columns=columnas_renombre)
        
        # Generar nombre de archivo
        fecha = datetime.now().strftime("%Y%m%d_%H%M%S")
        sanitized_dept = departamento_filtro.replace(" ", "_")
        filename = f"Exportacion_{sanitized_dept}_{fecha}.xlsx"
        filepath = os.path.join(self.output_folder, filename)
        
        os.makedirs(self.output_folder, exist_ok=True)
        
        # Guardar Excel
        try:
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
