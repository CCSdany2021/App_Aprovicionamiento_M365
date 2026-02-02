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
            elif response.status_code == 404:
                print(f"❌ Error 404: El remitente '{self.email_sender}' NO EXISTE en este tenant.")
                print("   >> Verifique en Configuración que el 'Correo Remitente' sea válido en esta organización.")
                return False
            elif response.status_code == 403:
                print(f"❌ Error 403: Permisos denegados para enviar correo como '{self.email_sender}'.")
                print("   >> Asegúrese de que la App en Azure tenga el permiso 'Mail.Send' (Application).")
                return False
            else:
                print(f"❌ Error enviando correo vía Graph API: {response.status_code} - {response.text}")
                return False
                
        except Exception as e:
            print(f"❌ Error crítico enviando correo: {e}")
            return False
    def enviar_reset_password(self, email_destino, nombre_usuario, password_nueva):
        """Envía el correo de restablecimiento de contraseña"""
        if not self.email_sender:
             return False, "Remitente no configurado"

        if not self.token:
            if not self.obtener_token():
                return False, "Error de token"

        try:
            # Usaremos una plantilla simple si no existe archivo
            # Idealmente crear templates/emails/reset_password.html
            asunto = f"🔑 Recuperación de Contraseña - {nombre_usuario}"
            
            # HTML simple inline para no depender de archivo nuevo ahora mismo
            html_content = f"""
            <html>
            <body style="font-family: 'Segoe UI', Arial, sans-serif; color: #333; line-height: 1.6;">
                <div style="max-width: 600px; margin: 0 auto; border: 1px solid #e5e7eb; padding: 0; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                    <div style="background-color: #2563eb; padding: 20px; text-align: center;">
                        <h2 style="color: white; margin: 0; font-size: 24px;">🔐 Recuperación de Acceso</h2>
                    </div>
                    
                    <div style="padding: 30px 25px;">
                        <h3 style="color: #1f2937; margin-top: 0;">Hola, {nombre_usuario}</h3>
                        <p style="color: #4b5563;">Has solicitado restablecer tu contraseña para la cuenta institucional de Microsoft 365.</p>
                        
                        <p style="margin-bottom: 20px;">Tu nueva contraseña temporal es:</p>
                        
                        <div style="background-color: #eff6ff; border: 2px dashed #bfdbfe; padding: 20px; text-align: center; border-radius: 8px; margin: 25px 0;">
                            <span style="font-size: 24px; font-weight: bold; letter-spacing: 2px; color: #1e40af; font-family: monospace;">{password_nueva}</span>
                        </div>
                        
                        <div style="background-color: #fffbeb; border-left: 4px solid #f59e0b; padding: 15px; margin-bottom: 25px; font-size: 14px; color: #92400e;">
                            <strong>⚠️ Importante:</strong> Esta contraseña es temporal. Al iniciar sesión se te pedirá crear una nueva contraseña PERSONAL que solo tú conozcas.
                        </div>

                        <div style="text-align: center; margin-top: 30px;">
                            <a href="https://www.office.com" style="background-color: #2563eb; color: white; padding: 12px 25px; text-decoration: none; border-radius: 6px; font-weight: bold; display: inline-block;">Iniciar Sesión en Office.com</a>
                        </div>
                    </div>
                    
                    <div style="background-color: #f9fafb; padding: 15px; text-align: center; border-top: 1px solid #e5e7eb; font-size: 12px; color: #9ca3af;">
                        <p style="margin: 0;">Este es un mensaje automático del sistema de Tecnología.</p>
                        <p style="margin: 5px 0 0 0;">No respondas a este correo.</p>
                    </div>
                </div>
            </body>
            </html>
            """
            
            
            # Construir JSON para Microsoft Graph
            email_data = {
                "message": {
                    "subject": asunto,
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
                print(f"✅ Correo de reset enviado a {email_destino}")
                return True, "Correo enviado correctamente"
            else:
                return False, f"Error API: {response.text}"
                
        except Exception as e:
            return False, f"Error: {str(e)}"
