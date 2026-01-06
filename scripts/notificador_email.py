import requests
import json
import os
import sys
from jinja2 import Environment, FileSystemLoader
from scripts.configuracion import config

class NotificadorEmail:
    """Clase para manejar el envío de notificaciones usando Microsoft Graph API"""
    
    def __init__(self, token=None):
        self.token = token
        self.email_sender = config.EMAIL_SENDER
        
        # Configurar Jinja2 para cargar plantillas
        template_dir = os.path.join(os.path.dirname(__file__), '..', 'templates', 'emails')
        self.jinja_env = Environment(loader=FileSystemLoader(template_dir))

    def obtener_token(self):
        """Obtiene token usando las credenciales de la aplicación (Client Credentials)"""
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
        except Exception as e:
            print(f"❌ Error obteniendo token para envío de correo: {e}")
            return False

    def enviar_credenciales(self, email_destino, nombre_estudiante, upn, password_temporal):
        """Envía el correo con las credenciales usando Microsoft Graph API"""
        if not self.email_sender:
            print("⚠️ EMAIL_SENDER no configurado en el archivo .env. El correo no será enviado.")
            return False
            
        if not self.token:
            if not self.obtener_token():
                return False

        try:
            # Renderizar plantilla
            template = self.jinja_env.get_template('credenciales.html')
            html_content = template.render(
                nombre_aspirante=nombre_estudiante,
                email_institucional=upn,
                password_temporal=password_temporal
            )
            
            # Construir JSON para Microsoft Graph
            email_data = {
                "message": {
                    "subject": f"🔐 Credenciales de Acceso - {nombre_estudiante}",
                    "body": {
                        "contentType": "HTML",
                        "content": html_content
                    },
                    "toRecipients": [
                        {
                            "emailAddress": {
                                "address": email_destino
                            }
                        }
                    ]
                },
                "saveToSentItems": "true"
            }
            
            # Enviar vía Graph API
            url = f"{config.GRAPH_ENDPOINT}/users/{self.email_sender}/sendMail"
            headers = {
                "Authorization": f"Bearer {self.token}",
                "Content-Type": "application/json"
            }
            
            response = requests.post(url, headers=headers, json=email_data, verify=False)
            
            if response.status_code == 202:
                print(f"✅ Correo enviado exitosamente vía Graph API a {email_destino}")
                return True
            else:
                print(f"❌ Error enviando correo vía Graph API: {response.status_code} - {response.text}")
                return False
                
        except Exception as e:
            print(f"❌ Error crítico enviando correo: {e}")
            return False

if __name__ == "__main__":
    # Prueba rápida si se ejecuta directamente
    notificador = NotificadorEmail()
    # notificador.enviar_credenciales("test@ejemplo.com", "Estudiante de Prueba", "prueba@dominio.com", "Temp123!")
