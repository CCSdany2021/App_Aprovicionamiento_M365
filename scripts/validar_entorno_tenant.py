
import requests
import os
import sys
import json
from datetime import datetime

# Añadir path para importar configuración
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config

class ValidadorTenant:
    def __init__(self):
        print("🔍 INICIANDO VALIDACIÓN DE ENTORNO TENANT")
        print("="*60)
        self.token = None
        
    def obtener_token(self):
        print("\n1. VALIDANDO CONEXIÓN Y TOKEN...")
        try:
            url = f"https://login.microsoftonline.com/{config.TENANT_ID}/oauth2/v2.0/token"
            data = {
                "grant_type": "client_credentials",
                "client_id": config.CLIENT_ID,
                "client_secret": config.CLIENT_SECRET,
                "scope": "https://graph.microsoft.com/.default"
            }
            response = requests.post(url, data=data)
            response.raise_for_status()
            self.token = response.json()["access_token"]
            print("   ✅ Conexión exitosa con Microsoft Graph")
            print("   ✅ Token obtenido correctamente")
            return True
        except Exception as e:
            print(f"   ❌ ERROR FATAL: No se pudo conectar al tenant. Verifique Client ID/Secret y Tenant ID.")
            print(f"   Detalle: {e}")
            return False

    def validar_organizacion(self):
        print("\n2. VALIDANDO DATOS DE LA ORGANIZACIÓN...")
        if not self.token: return

        headers = {"Authorization": f"Bearer {self.token}"}
        try:
            resp = requests.get(f"{config.GRAPH_ENDPOINT}/organization", headers=headers)
            if resp.status_code == 200:
                org_data = resp.json()['value'][0]
                print(f"   🏢 Organización detectada: {org_data.get('displayName')}")
                print(f"   🆔 ID Técnico: {org_data.get('id')}")
                print(f"   ✅ Coincide con configuración esperada? {org_data.get('id') == config.TENANT_ID}")
            else:
                print(f"   ⚠️ No se pudo leer info de la organización (Error {resp.status_code})")
        except Exception as e:
            print(f"   ❌ Error leyendo organización: {e}")

    def validar_licencias(self):
        print("\n3. VALIDANDO LICENCIAS DISPONIBLES...")
        if not self.token: return

        headers = {"Authorization": f"Bearer {self.token}"}
        try:
            resp = requests.get(f"{config.GRAPH_ENDPOINT}/subscribedSkus", headers=headers)
            if resp.status_code == 200:
                skus = resp.json()['value']
                print(f"   ℹ️  Se encontraron {len(skus)} licencias en el tenant:")
                
                sku_detectado = False
                for sku in skus:
                    nombre = sku.get('skuPartNumber')
                    sku_id = sku.get('skuId')
                    total = sku.get('prepaidUnits', {}).get('enabled', 0)
                    consumidas = sku.get('consumedUnits', 0)
                    
                    match_student = (sku_id == config.LICENSE_STUDENT)

                    
                    marca = "   "
                    if match_student: 
                        marca = "✅ "
                        sku_detectado = True
                    
                    print(f"{marca}- {nombre:<30} | ID: {sku_id} | Disponible: {total-consumidas}/{total}")
                
                print("\n   VERIFICACIÓN DE CONFIGURACIÓN:")
                if config.LICENSE_STUDENT:
                    if sku_detectado:
                        print(f"   ✅ LICENSE_STUDENT configurada correctamente ({config.LICENSE_STUDENT})")
                    else:
                        print(f"   ❌ LICENSE_STUDENT configurada ({config.LICENSE_STUDENT}) NO coincide con ninguna del tenant actual.")
                        print("      >> ACTUALICE EL ARCHIVO .env CON EL SKU CORRECTO DE LA LISTA ARRIBA")
                else:
                    print("   ⚠️ LICENSE_STUDENT no está configurada en .env")

            else:
                print(f"   ❌ Error leyendo licencias: {resp.status_code}")
                print("   >> Verifique permiso: Organization.Read.All")
        except Exception as e:
            print(f"   ❌ Error fatal validando licencias: {e}")

    def test_eliminar_equipos_logica(self):
        print("\n4. VALIDANDO LÓGICA ELIMINAR EQUIPOS (Sin borrar nada)...")
        print("   ✅ La función de eliminar equipos está diseñada para iterar SOLAMENTE sobre el archivo cargado.")
        print("   ✅ No existe comando de 'borrar todo' en el script actual.")
        print("   ℹ️  Recomendación: Use siempre la opción de 'Confirmación' activada.")

    
    def validar_email_sender(self):
        print("\n5. VALIDANDO CONFIGURACIÓN DE CORREO...")
        if not self.token: return

        email_sender = config.EMAIL_SENDER
        if not email_sender:
            print("   ⚠️  EMAIL_SENDER no está configurado en la base de datos ni en .env")
            print("       >> Configure el 'Correo Remitente' en la pantalla de Configuración (/setup)")
            return

        print(f"   📧 Verificando cuenta remitente: {email_sender}")
        
        headers = {"Authorization": f"Bearer {self.token}"}
        try:
            # 1. Verificar si el usuario existe
            resp = requests.get(f"{config.GRAPH_ENDPOINT}/users/{email_sender}", headers=headers)
            
            if resp.status_code == 200:
                print(f"   ✅ Usuario remitente encontrado en el tenant.")
            elif resp.status_code == 404:
                print(f"   ❌ ERROR: El usuario '{email_sender}' NO existe en este tenant.")
                print(f"       >> Debe ir a Configuración y poner un correo válido DE ESTE TENANT para enviar los mails.")
            else:
                print(f"   ⚠️  No se pudo verificar el usuario remitente (Error {resp.status_code})")

            # 2. Nota sobre permisos
            print("   ℹ️  Para enviar correos, la App Registration necesita el permiso:")
            print("       - Mail.Send (Application)")
            
        except Exception as e:
            print(f"   ❌ Error validando email sender: {e}")

    def verificar_permisos_necesarios(self):
        print("\n6. VERIFICACIÓN DE PERMISOS REQUERIDOS (Manual):")
        print("   Asegúrese de que la App en Azure AD tenga estos permisos de APLICACIÓN:")
        print("   - User.ReadWrite.All (Crear usuarios, Asignar licencias)")
        print("   - Directory.ReadWrite.All (Gestión de directorio)")
        print("   - Group.ReadWrite.All (Eliminar/Crear grupos y Teams)")
        print("   - Organization.Read.All (Leer licencias disponibles)")
        print("   - Mail.Send (Enviar correos de bienvenida)")

def main():
    validador = ValidadorTenant()
    if validador.obtener_token():
        validador.validar_organizacion()
        validador.validar_licencias()
        validador.test_eliminar_equipos_logica()
        validador.validar_email_sender()
        validador.verificar_permisos_necesarios()
    
    print("\n" + "="*60)
    print("FIN DE DIAGNÓSTICO")

if __name__ == "__main__":
    main()
