
import requests
import json
import os
import sys

# Añadir path para importar configuración
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config

class GestorUsuarios:
    def __init__(self):
        self.token = None
        self._auth()

    def _auth(self):
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
        except Exception as e:
            print(f"Error auth: {e}")
            self.token = None

    def _ensure_token(self):
        """Verifica y refresca el token si es necesario"""
        if not self.token:
            self._auth()
        return self.token is not None

    def buscar_usuarios(self, query):
        if not self._ensure_token(): return []
        
        headers = {"Authorization": f"Bearer {self.token}", "ConsistencyLevel": "eventual"}
        q = query.replace("'", "")
        url = f"{config.GRAPH_ENDPOINT}/users"
        
        params_search = {
            "$search": f"\"displayName:{q}\" OR \"userPrincipalName:{q}\" OR \"mailNickname:{q}\"",
            "$select": "id,displayName,userPrincipalName,jobTitle,department,mailNickname",
            "$top": 20
        }

        # Lógica de Retry para 401
        for attempt in range(2):
            try:
                # 1. Intento Búsqueda Avanzada
                resp = requests.get(url, headers=headers, params=params_search)
                
                if resp.status_code == 401 and attempt == 0:
                    print("🔄 Token expirado (401). Refrescando...")
                    self._auth()
                    headers["Authorization"] = f"Bearer {self.token}"
                    continue # Reintentar loop
                
                if resp.status_code == 200:
                    results = resp.json().get('value', [])
                    if not results:
                        # 2. Fallback Filter Simple
                        params_simple = {
                            "$filter": f"startswith(displayName, '{q}') or startswith(mailNickname, '{q}')",
                            "$select": "id,displayName,userPrincipalName,jobTitle,department,mailNickname",
                            "$top": 10
                        }
                        resp2 = requests.get(url, headers=headers, params=params_simple)
                        if resp2.status_code == 200:
                            results = resp2.json().get('value', [])
                    return results
                else:
                    print(f"Error search: {resp.status_code} - {resp.text}")
                    return []
                    
            except Exception as e:
                print(f"Excepcion buscar: {e}")
                return []
        return []

    def restablecer_contrasena(self, user_id, nueva_password):
        if not self.token: return False, "No token"
        
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        url = f"{config.GRAPH_ENDPOINT}/users/{user_id}"
        
        data = {
            "passwordProfile": {
                "forceChangePasswordNextSignIn": True,
                "password": nueva_password
            }
        }
        
        try:
            resp = requests.patch(url, headers=headers, json=data)
            if resp.status_code == 204:
                return True, "Contraseña restablecida correctamente"
            else:
                return False, f"Error API: {resp.text}"
        except Exception as e:
            return False, str(e)
