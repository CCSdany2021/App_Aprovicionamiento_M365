
import sys
import os
import requests
import urllib3

# Añadir directorio raíz al path
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config
from scripts.sincronizador_automatico_teams import SincronizadorAutomaticoTeams

urllib3.disable_warnings()

def diagnosticar_equipos():
    print("--- INICIO DIAGNÓSTICO EQUIPOS ---")
    try:
        # Inicializar sincronizador para reutilizar lógica de token
        sinc = SincronizadorAutomaticoTeams()
        if not sinc.obtener_token():
            print("❌ No se pudo obtener el token de acceso.")
            return

        print("✅ Token obtenido correctamente.")
        
        # Consultar los primeros 50 grupos para ver sus nombres
        print("\n🔍 Consultando muestra de grupos en el Tenant...")
        url = f"{config.GRAPH_ENDPOINT}/groups?$select=id,displayName&$top=50"
        headers = {'Authorization': f'Bearer {sinc.token}'}
        
        response = requests.get(url, headers=headers, verify=False)
        
        if response.status_code != 200:
            print(f"❌ Error consultando API: {response.status_code} - {response.text}")
            return

        grupos = response.json().get('value', [])
        print(f"Total grupos en la muestra: {len(grupos)}")
        
        if not grupos:
            print("⚠️ No se encontraron grupos en el Tenant.")
            return

        print("\n📋 Listado de nombres encontrados:")
        con_formato_correcto = 0
        for g in grupos:
            nombre = g.get('displayName', 'Sin Nombre')
            tiene_formato = ": " in nombre
            estado = "✅ CORRECTO" if tiene_formato else "⚠️ SIN FORMATO (Esperado ': ')"
            print(f"   - [{estado}] {nombre}")
            if tiene_formato:
                con_formato_correcto += 1
                
        print("\n" + "="*40)
        print(f"Resumen Muestra:")
        print(f"Total analizados: {len(grupos)}")
        print(f"Con formato 'CURSO: MATERIA': {con_formato_correcto}")
        print("Si el número de correctos es bajo, el script de sincronización ignorará los demás.")
        print("El formato esperado debe tener DOS PUNTOS Y ESPACIO. Ej: '901: Matemáticas'")
        print("="*40)

    except Exception as e:
        print(f"\n❌ Error Crítico en diagnóstico: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    diagnosticar_equipos()
