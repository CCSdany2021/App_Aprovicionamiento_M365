import os
import sys
from dotenv import load_dotenv

# Añadir directorio actual al path
sys.path.append(os.getcwd())

from scripts.configuracion import config
from scripts.notificador_email import NotificadorEmail

def test_graph_email():
    load_dotenv(override=True)
    print("--- Verificando Configuración Graph API ---")
    print(f"TENANT_ID: {config.TENANT_ID[:8]}...")
    print(f"CLIENT_ID: {config.CLIENT_ID[:8]}...")
    print(f"EMAIL_SENDER: {config.EMAIL_SENDER}")
    
    if not config.EMAIL_SENDER:
        print("\n❌ ERROR: Falta EMAIL_SENDER en el archivo .env")
        return

    notificador = NotificadorEmail()
    print("\nIntento de envío vía Graph API...")
    
    # Prueba enviando a la misma cuenta emisora para verificar permisos
    exito = notificador.enviar_credenciales(
        config.EMAIL_SENDER, 
        "ADMIN GRAPH TEST", 
        "test_graph@calasanzsuba.edu.co", 
        "GraphPass123!"
    )
    
    if exito:
        print("\n✅ PRUEBA EXITOSA: El correo fue enviado vía Graph API.")
    else:
        print("\n❌ PRUEBA FALLIDA: Revisa los errores de permisos en la consola.")

if __name__ == "__main__":
    test_graph_email()
