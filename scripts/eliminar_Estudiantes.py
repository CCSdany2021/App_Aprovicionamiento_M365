import pandas as pd
import requests
import urllib3
from datetime import datetime
import os
import sys

# Añadir la carpeta scripts al path para importar configuración
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class EliminadorEstudiantes:
    """Clase para eliminar estudiantes de prueba del tenant"""
    
    def __init__(self):
        # Validar configuración al inicializar
        config.validar_configuracion()
        self.token = None
        self.resultados = {
            "total": 0,
            "eliminados": 0,
            "no_encontrados": 0,
            "errores": 0,
            "detalles": []
        }
        
    def obtener_token(self) -> bool:
        """Obtiene token de acceso a Microsoft Graph API"""
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
            print("✅ Token obtenido correctamente")
            return True
        except requests.RequestException as e:
            print(f"❌ Error obteniendo token: {e}")
            return False

    def verificar_usuario_existe(self, codigo_estudiante: str) -> bool:
        """Verifica si el usuario existe en el tenant"""
        if not self.token:
            return False
            
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        user_email = f"{codigo_estudiante}@{config.COLEGIO_DOMINIO}"
        url = f"{config.GRAPH_ENDPOINT}/users/{user_email}"
        
        try:
            response = requests.get(url, headers=headers, verify=False)
            return response.status_code == 200
        except requests.RequestException:
            return False

    def eliminar_estudiante(self, codigo_estudiante: str) -> tuple[bool, str]:
        """Elimina un estudiante del tenant"""
        if not self.token:
            return False, "Token no disponible"
            
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        user_email = f"{codigo_estudiante}@{config.COLEGIO_DOMINIO}"
        url = f"{config.GRAPH_ENDPOINT}/users/{user_email}"
        
        try:
            # Primero verificar si existe
            if not self.verificar_usuario_existe(codigo_estudiante):
                return False, f"Usuario {codigo_estudiante} no encontrado"
            
            # Eliminar usuario
            response = requests.delete(url, headers=headers, verify=False)
            
            if response.status_code == 204:
                return True, f"Usuario {codigo_estudiante} eliminado exitosamente"
            else:
                return False, f"Error eliminando {codigo_estudiante}: {response.text}"
                
        except requests.RequestException as e:
            return False, f"Error de conexión eliminando {codigo_estudiante}: {e}"

    def cargar_lista_estudiantes(self, ruta_archivo: str = None) -> list:
        """Carga lista de estudiantes a eliminar desde archivo o rango"""
        estudiantes = []
        
        if ruta_archivo and os.path.exists(ruta_archivo):
            try:
                if ruta_archivo.endswith(".xlsx"):
                    df = pd.read_excel(ruta_archivo, dtype=str)
                elif ruta_archivo.endswith(".csv"):
                    df = pd.read_csv(ruta_archivo, dtype=str, encoding="utf-8")
                else:
                    raise ValueError("❌ Formato no soportado. Usa .xlsx o .csv")
                
                # Obtener códigos de estudiantes
                if 'CODIGO' in df.columns:
                    estudiantes = df['CODIGO'].astype(str).tolist()
                else:
                    print("❌ No se encontró columna 'CODIGO' en el archivo")
                    return []
                    
                print(f"✅ Cargados {len(estudiantes)} códigos desde archivo")
                
            except Exception as e:
                print(f"❌ Error leyendo archivo: {e}")
                return []
        else:
            # Generar rango de códigos de prueba (40302001-40302200)
            estudiantes = [str(40302000 + i + 1) for i in range(200)]
            print(f"✅ Generado rango de códigos de prueba: {len(estudiantes)} códigos")
        
        return estudiantes

    def eliminar_masivo_con_confirmacion(self, codigos_estudiantes: list, confirmacion: bool = True) -> dict:
        """Elimina estudiantes con confirmaciones de seguridad
        
        Args:
            codigos_estudiantes (list): Lista de códigos.
            confirmacion (bool, optional): Pedir confirmación. Defaults to True.
            
        Returns:
            dict: Resultados del proceso
        """
        self.resultados["total"] = len(codigos_estudiantes)
        
        if confirmacion:
            print(f"\n⚠️  ADVERTENCIA IMPORTANTE ⚠️")
            print("="*60)
            print("Este proceso eliminará PERMANENTEMENTE los usuarios del tenant.")
            print("NO se pueden recuperar una vez eliminados.")
            print(f"Se eliminarán {len(codigos_estudiantes)} usuarios potenciales.")
            print("="*60)
            
            # Primera confirmación
            respuesta1 = input("\n¿Está SEGURO de que desea continuar? (escriba 'SI ELIMINAR'): ")
            if respuesta1 != "SI ELIMINAR":
                print("❌ Operación cancelada por seguridad")
                return self.resultados
            
            # Segunda confirmación con nombre del colegio
            respuesta2 = input(f"\n¿Confirma eliminar usuarios del tenant '{config.COLEGIO_NOMBRE}'? (escriba el nombre del colegio): ")
            if respuesta2 != config.COLEGIO_NOMBRE:
                print("❌ Nombre del colegio no coincide. Operación cancelada")
                return self.resultados
            
            # Tercera confirmación
            respuesta3 = input("\nÚltima confirmación. ¿Proceder con la eliminación? (si/no): ").lower()
            if respuesta3 not in ['si', 's', 'yes', 'y']:
                print("❌ Operación cancelada")
                return self.resultados
        
        print(f"\n🗑️  Iniciando eliminación de usuarios...")
        print("="*50)
        
        # Procesar eliminaciones
        for index, codigo in enumerate(codigos_estudiantes, 1):
            print(f"\n🔍 Procesando {index}/{len(codigos_estudiantes)}: {codigo}")
            
            try:
                exito, mensaje = self.eliminar_estudiante(codigo)
                
                if exito:
                    self.resultados["eliminados"] += 1
                    print(f"✅ {mensaje}")
                elif "no encontrado" in mensaje.lower():
                    self.resultados["no_encontrados"] += 1
                    print(f"⚪ {mensaje}")
                else:
                    self.resultados["errores"] += 1
                    print(f"❌ {mensaje}")
                
                self.resultados["detalles"].append(f"{codigo}: {mensaje}")
                
            except Exception as e:
                error_msg = f"Error inesperado procesando {codigo}: {e}"
                print(f"❌ {error_msg}")
                self.resultados["errores"] += 1
                self.resultados["detalles"].append(f"{codigo}: {error_msg}")
        
        # Mostrar resumen final
        self.mostrar_resumen()
        self.guardar_log()
        
        return self.resultados

    def mostrar_resumen(self):
        """Muestra resumen de la operación de eliminación"""
        print("\n" + "="*60)
        print("🗑️  RESUMEN DE ELIMINACIÓN")
        print("="*60)
        print(f"🏫 Colegio: {config.COLEGIO_NOMBRE}")
        print(f"📅 Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"📊 Total procesados: {self.resultados['total']}")
        print(f"✅ Usuarios eliminados: {self.resultados['eliminados']}")
        print(f"⚪ Usuarios no encontrados: {self.resultados['no_encontrados']}")
        print(f"❌ Errores: {self.resultados['errores']}")
        print("="*60)

    def guardar_log(self):
        """Guarda log detallado de la eliminación"""
        try:
            os.makedirs(config.CARPETA_LOGS, exist_ok=True)
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            log_file = os.path.join(config.CARPETA_LOGS, f'eliminacion_estudiantes_{timestamp}.log')
            
            with open(log_file, 'w', encoding='utf-8') as f:
                f.write(f"ELIMINACIÓN DE ESTUDIANTES - {config.COLEGIO_NOMBRE}\n")
                f.write(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("="*50 + "\n")
                f.write(f"Total procesados: {self.resultados['total']}\n")
                f.write(f"Usuarios eliminados: {self.resultados['eliminados']}\n")
                f.write(f"Usuarios no encontrados: {self.resultados['no_encontrados']}\n")
                f.write(f"Errores: {self.resultados['errores']}\n\n")
                
                f.write("DETALLES POR USUARIO:\n")
                f.write("-"*30 + "\n")
                for detalle in self.resultados["detalles"]:
                    f.write(f"{detalle}\n")
            
            print(f"📝 Log detallado guardado en: {log_file}")
            
        except Exception as e:
            print(f"❌ Error guardando log: {e}")

def main():
    """Función principal"""
    print("🗑️  ELIMINADOR DE ESTUDIANTES DE PRUEBA")
    print(f"🏫 {config.COLEGIO_NOMBRE}")
    print("="*50)
    
    try:
        eliminador = EliminadorEstudiantes()
        
        # Obtener token de acceso
        if not eliminador.obtener_token():
            print("❌ No se pudo obtener token de acceso")
            return
        
        # Opciones de eliminación
        print("\nOpciones de eliminación:")
        print("1. Eliminar rango de códigos de prueba (40302001-40302200)")
        print("2. Eliminar desde archivo Excel/CSV")
        print("3. Cancelar")
        
        opcion = input("\nSeleccione una opción (1-3): ").strip()
        
        if opcion == "1":
            # Eliminar rango de prueba
            codigos = eliminador.cargar_lista_estudiantes()
            if codigos:
                eliminador.eliminar_masivo_con_confirmacion(codigos)
        
        elif opcion == "2":
            # Eliminar desde archivo
            ruta_archivo = input("Ruta del archivo (.xlsx o .csv): ").strip()
            codigos = eliminador.cargar_lista_estudiantes(ruta_archivo)
            if codigos:
                eliminador.eliminar_masivo_con_confirmacion(codigos)
        
        elif opcion == "3":
            print("❌ Operación cancelada")
        
        else:
            print("❌ Opción no válida")
            
    except KeyboardInterrupt:
        print("\n❌ Proceso interrumpido por el usuario")
    except Exception as e:
        print(f"❌ Error inesperado: {e}")

if __name__ == "__main__":
    main()