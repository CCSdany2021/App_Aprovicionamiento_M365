import requests
import os
import sys
from dotenv import load_dotenv

# Añadir directorio actual al path
sys.path.append(os.getcwd())

from scripts.configuracion import config
from scripts.notificador_email import NotificadorEmail

def test_graph_email():
    load_dotenv(override=True)
    print("--- DEPURAClÓN DE GRAPH API ---")
    
    # 1. Probar obtención de Token
    print("\n1. Intentando obtener token...")
    url = f"https://login.microsoftonline.com/{config.TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": config.CLIENT_ID,
        "client_secret": config.CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }
    
    try:
        resp = requests.post(url, data=data, verify=False)
        if resp.status_code == 200:
            token = resp.json().get("access_token")
            print("✅ Token obtenido exitosamente.")
        else:
            print(f"❌ Error en Token (HTTP {resp.status_code}):")
            print(resp.text)
            return
    except Exception as e:
        print(f"❌ Error de conexión: {e}")
        return

    # 2. Probar envío de correo
    print(f"\n2. Intentando enviar correo desde: {config.EMAIL_SENDER}...")
    notificador = NotificadorEmail(token=token)
    
    # Intentamos enviar a la misma cuenta emisora
    url_mail = f"{config.GRAPH_ENDPOINT}/users/{config.EMAIL_SENDER}/sendMail"
    email_data = {
        "message": {
            "subject": "TEST GRAPH API",
            "body": {"contentType": "Text", "content": "Prueba de conexión"},
            "toRecipients": [{"emailAddress": {"address": config.EMAIL_SENDER}}]
        }
    }
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    try:
        resp_mail = requests.post(url_mail, headers=headers, json=email_data, verify=False)
        if resp_mail.status_code == 202:
            print("✅ ENVÍO EXITOSO!")
        else:
            print(f"❌ Error en Envío (HTTP {resp_mail.status_code}):")
            print(resp_mail.text)
    except Exception as e:
        print(f"❌ Error de conexión al enviar: {e}")

if __name__ == "__main__":
    test_graph_email()
