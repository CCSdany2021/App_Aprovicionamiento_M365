
import os
import sys
import requests
import urllib3
from datetime import datetime

# Añadir path
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def check_env_vars():
    print("🔍 REVISIÓN DE VARIABLES DE ENTORNO (.env)")
    print("="*50)
    
    required_vars = [
        "TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", 
        "COLEGIO_DOMINIO", "LICENSE_STUDENT"
    ]
    
    missing = []
    for var in required_vars:
        val = os.getenv(var)
        if not val:
            print(f"❌ {var}: FALTANTE")
            missing.append(var)
        else:
            # Mask secret
            if "SECRET" in var:
                val = val[:4] + "..." + val[-4:]
            print(f"✅ {var}: {val}")
            
    if missing:
        print("\n❌ CRÍTICO: Faltan variables obligatorias en .env")
        return False
    return True

def check_graph_connectivity():
    print("\n🔍 PRUEBA DE CONEXIÓN MICROSOFT GRAPH")
    print("="*50)
    
    url = f"https://login.microsoftonline.com/{config.TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": config.CLIENT_ID,
        "client_secret": config.CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }
    
    try:
        print(f"Attempting to get token for Client ID: {config.CLIENT_ID}")
        response = requests.post(url, data=data, verify=False, timeout=10)
        
        if response.status_code == 200:
            print("✅ Conexión EXITOSA: Token obtenido correctamente.")
            token = response.json().get('access_token')
            
            # Prueba de lectura de usuarios (basic check)
            headers = {"Authorization": f"Bearer {token}"}
            user_url = "https://graph.microsoft.com/v1.0/users?$top=1"
            user_resp = requests.get(user_url, headers=headers, verify=False)
            
            if user_resp.status_code == 200:
                print("✅ Permisos de Lectura (User.Read.All): OK")
                print(f"   Usuario de prueba: {user_resp.json()['value'][0].get('userPrincipalName')}")
            else:
                print(f"⚠️  Token válido pero falla lectura de usuarios: {user_resp.status_code}")
                print(f"   Posible falta de permisos 'User.Read.All' o 'User.ReadWrite.All' en Azure AD.")
                
            return True
        elif response.status_code == 401:
            print("❌ ERROR 401: CREDENCIALES INVÁLIDAS")
            print("   El 'Client Secret' ha expirado o no coincide con el 'Client ID'.")
            print("   O el 'Tenant ID' es incorrecto.")
            return False
        else:
            print(f"❌ Error obteniendo token: {response.status_code}")
            print(response.text)
            return False
            
    except Exception as e:
        print(f"❌ Excepción de conexión: {e}")
        return False

def check_folders():
    print("\n🔍 VERIFICACIÓN DE ESTRUCTURA DE CARPETAS")
    print("="*50)
    
    folders = [
        "archivos", 
        "archivos/plantillas", 
        "resultados", 
        "resultados/logs", 
        "static", 
        "templates"
    ]
    
    for folder in folders:
        path = os.path.join(os.getcwd(), folder)
        if os.path.exists(path):
            print(f"✅ Carpeta existe: {folder}")
        else:
            try:
                os.makedirs(path)
                print(f"⚠️  Carpeta creada automáticamente: {folder}")
            except:
                print(f"❌ No se pudo crear carpeta: {folder}")

def run_diagnostics():
    print("\n🛠️  DIAGNÓSTICO GENERAL DEL SISTEMA")
    print("="*60)
    print(f"Fecha: {datetime.now()}")
    print("="*60)
    
    env_ok = check_env_vars()
    if env_ok:
        check_graph_connectivity()
    
    check_folders()
    print("\n🏁 Diagnóstico finalizado.")

if __name__ == "__main__":
    run_diagnostics()
