import msal
import os
import sys
import webbrowser
from scripts.configuracion import config

class AzureInteractiveAuth:
    """Clase para manejar el flujo de autenticación interactiva de Azure AD"""
    
    def __init__(self):
        self.client_id = config.CLIENT_ID
        self.tenant_id = config.TENANT_ID
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self.scopes = ["https://graph.microsoft.com/.default"]
        
        self.app = msal.PublicClientApplication(
            self.client_id,
            authority=self.authority
        )

    def get_token(self):
        """Obtiene un token de acceso mediante flujo interactivo"""
        accounts = self.app.get_accounts()
        result = None

        if accounts:
            # Intentar obtener token silenciosamente
            result = self.app.acquire_token_silent(self.scopes, account=accounts[0])

        if not result:
            # Flujo interactivo si no hay token en cache
            print("Iniciando flujo de autenticación interactiva...")
            result = self.app.acquire_token_interactive(scopes=self.scopes)

        if "access_token" in result:
            return result["access_token"]
        else:
            error_msg = result.get("error_description", "Error desconocido en la autenticación")
            raise Exception(f"No se pudo obtener el token: {error_msg}")

if __name__ == "__main__":
    try:
        auth = AzureInteractiveAuth()
        token = auth.get_token()
        print("✅ Token obtenido exitosamente")
    except Exception as e:
        print(f"❌ Error: {e}")
