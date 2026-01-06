import os
import sys
from dotenv import load_dotenv

# Añadir directorio actual al path
sys.path.append(os.getcwd())

from scripts.configuracion import config
from scripts.notificador_email import NotificadorEmail

def test_email_config():
    load_dotenv(override=True)
    print("--- Verificando Configuración SMTP ---")
    print(f"SMTP_SERVER: {config.SMTP_SERVER}")
    print(f"SMTP_PORT: {config.SMTP_PORT}")
    print(f"SMTP_USER: {config.SMTP_USER}")
    print(f"SMTP_PASSWORD: {'******' if config.SMTP_PASSWORD else 'NO CONFIGURADO'}")
    print(f"EMAIL_FROM: {config.EMAIL_FROM}")
    
    if not config.SMTP_USER or not config.SMTP_PASSWORD:
        print("\n❌ ERROR: Faltan variables en el archivo .env")
        return

    notificador = NotificadorEmail()
    print("\nIntento de envío de prueba...")
    # Prueba enviando a la misma cuenta de origen
    exito = notificador.enviar_credenciales(
        config.SMTP_USER, 
        "ADMIN TEST", 
        "test@calasanzsuba.edu.co", 
        "Prueba123!"
    )
    
    if exito:
        print("\n✅ PRUEBA EXITOSA: El correo de prueba fue enviado.")
    else:
        print("\n❌ PRUEBA FALLIDA: Revisa los errores arriba.")

if __name__ == "__main__":
    test_email_config()
