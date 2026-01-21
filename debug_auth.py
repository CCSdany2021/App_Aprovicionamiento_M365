import sys
import os
import requests
import urllib3

# Disable warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Add scripts to path
sys.path.append(os.path.join(os.getcwd(), 'scripts'))

try:
    from scripts.gestor_configuracion import GestorConfiguracion
    
    print("\n--- TEST DE AUTENTICACION AZURE AD ---")
    
    # Load from DB directly to be sure
    gestor = GestorConfiguracion()
    db_config = gestor.obtener_configuracion()
    
    if not db_config:
        print("❌ ERROR CRITICO: No hay configuración en la base de datos.")
        sys.exit(1)
        
    tenant_id = db_config.get('tenant_id')
    client_id = db_config.get('client_id')
    client_secret = db_config.get('client_secret')
    
    print(f"Tenant ID: {tenant_id}")
    print(f"Client ID: {client_id}")
    print(f"Secret Length: {len(client_secret) if client_secret else 0}")
    
    if not tenant_id or not client_id or not client_secret:
        print("❌ Faltan credenciales en la configuración.")
        sys.exit(1)
        
    # Attempt Token Request
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default"
    }
    
    print(f"\n📡 Conectando a: {url}...")
    response = requests.post(url, data=data, verify=False, timeout=10)
    
    print(f"Status Code: {response.status_code}")
    
    if response.status_code == 200:
        print("✅ AUTENTICACION EXITOSA")
        print("El token fue generado correctamente. Las credenciales son válidas.")
    else:
        print("❌ ERROR DE AUTENTICACION")
        try:
            error_data = response.json()
            print(f"Error: {error_data.get('error')}")
            print(f"Descripción: {error_data.get('error_description')}")
            
            desc = error_data.get('error_description', '').lower()
            if "aadsts700016" in desc:
                print("\n⚠️ DIAGNOSTICO: 'Application with identifier... was not found'.")
                print("   -> El CLIENT ID (Application ID) es incorrecto.")
            elif "aadsts7000215" in desc:
                print("\n⚠️ DIAGNOSTICO: 'Invalid client secret'.")
                print("   -> El CLIENT SECRET es incorrecto (verifique espacios o caracteres faltantes).")
            elif "aadsts90002" in desc:
                print("\n⚠️ DIAGNOSTICO: 'Tenant... not found'.")
                print("   -> El TENANT ID es incorrecto.")
        except:
            print(f"Respuesta Raw: {response.text}")

except Exception as e:
    print(f"❌ ERROR DE SCRIPT: {e}")
