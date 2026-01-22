"script para actualizar los estudiantes"

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

class ActualizadorEstudiantes:
    """Clase para actualizar estudiantes existentes en Microsoft 365"""
    
    def __init__(self):
        # Validar configuración al inicializar
        config.validar_configuracion()
        self.token = None
        self.resultados = {
            "total": 0,
            "actualizados": 0,
            "errores": 0,
            "detalles_errores": []
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
            print("Token obtenido correctamente")
            return True
        except requests.RequestException as e:
            print(f"Error obteniendo token: {e}")
            return False

    def actualizar_estudiante(self, estudiante: dict) -> bool:
        """Actualiza un estudiante individual en Microsoft 365"""
        if not self.token:
            print("Token no disponible")
            return False
            
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }

        user_principal_name = f"{estudiante['CODIGO']}@{config.COLEGIO_DOMINIO}"

        # Lógica inteligente: Si no vienen nombres/apellidos, los consultamos
        if "NOMBRES" not in estudiante or "APELLIDOS" not in estudiante or not estudiante["NOMBRES"] or not estudiante["APELLIDOS"]:
            try:
                # Consultar usuario actual
                print(f"   ℹ️ Consultando datos actuales para {estudiante['CODIGO']}...")
                get_url = f"{config.GRAPH_ENDPOINT}/users/{user_principal_name}?$select=givenName,surname"
                get_resp = requests.get(get_url, headers=headers, verify=False)
                
                if get_resp.status_code == 200:
                    current_user = get_resp.json()
                    estudiante["NOMBRES"] = current_user.get("givenName", "")
                    estudiante["APELLIDOS"] = current_user.get("surname", "")
                else:
                    print(f"   ⚠️ No se pudieron obtener datos actuales para '{user_principal_name}': {get_resp.status_code} (User Not Found)")
                    # Fallback riesgoso o saltar
            except Exception as e:
                 print(f"   ⚠️ Error consultando datos actuales: {e}")

        # Datos a actualizar usando configuración
        datos_actualizacion = {
            "displayName": f"Estudiante {estudiante['CURSO']}: {estudiante['APELLIDOS']} {estudiante['NOMBRES']}",
            "jobTitle": estudiante["CURSO"],
            "department": config.DEFAULT_DEPARTMENT,
            "city": config.DEFAULT_CITY
        }
        
        # Solo actualizar nombres si venían en el archivo (opcional, pero para consistencia del displayName los incluimos si cambiaron)
        # En este caso, como reconstruimos el displayName, asumimos que esos son los nombres correctos.
        # Si venían del archivo, actualizamos los campos individuales también.
        if "NOMBRES" in estudiante and estudiante["NOMBRES"]:
             datos_actualizacion["givenName"] = estudiante["NOMBRES"]
        if "APELLIDOS" in estudiante and estudiante["APELLIDOS"]:
             datos_actualizacion["surname"] = estudiante["APELLIDOS"]


        try:
            url = f"{config.GRAPH_ENDPOINT}/users/{user_principal_name}"
            response = requests.patch(url, headers=headers, json=datos_actualizacion, verify=False)

            if response.status_code == 204:
                print(f"Estudiante actualizado: {estudiante['CODIGO']}")
                return True
            else:
                error_msg = f"Error actualizando {estudiante['CODIGO']}: {response.text}"
                print(f"{error_msg}")
                self.resultados["detalles_errores"].append(error_msg)
                return False

        except requests.RequestException as e:
            error_msg = f"Error de conexión para {estudiante['CODIGO']}: {e}"
            print(f"{error_msg}")
            self.resultados["detalles_errores"].append(error_msg)
            return False

    def cargar_archivo(self, ruta_archivo: str) -> pd.DataFrame:
        """Carga estudiantes desde archivo Excel o CSV"""
        try:
            if ruta_archivo.endswith(".xlsx"):
                df = pd.read_excel(ruta_archivo, dtype=str)
            elif ruta_archivo.endswith(".csv"):
                df = pd.read_csv(ruta_archivo, dtype=str, encoding="utf-8", sep=";")
            else:
                raise ValueError("Formato no soportado. Usa .xlsx o .csv")
            
            # Limpiar datos
            df.columns = df.columns.str.strip()
            # Convertir todo a string y quitar espacios de CADA VALOR
            df = df.astype(str).map(lambda x: x.strip())
            df = df.fillna("")
            
            print(f"Archivo cargado: {len(df)} estudiantes para actualizar")
            return df
            
        except Exception as e:
            raise Exception(f"Error leyendo archivo: {e}")

    def validar_datos(self, df: pd.DataFrame) -> bool:
        """Valida que el DataFrame tenga las columnas mínimas necesarias"""
        # Columnas mínimas para operar (aunque el template tenga más)
        columnas_requeridas = ["CODIGO", "CURSO"]
        
        # Normalizar columnas del DF
        df.columns = [c.strip().upper() for c in df.columns]
        
        columnas_faltantes = [col for col in columnas_requeridas if col not in df.columns]

        if columnas_faltantes:
            print(f"❌ Faltan columnas requeridas: {columnas_faltantes}")
            return False
        
        print("✅ Datos válidos (Columnas detectadas)")
        return True

    def ejecutar(self, ruta_archivo: str = None):
        """Procesa la actualización masiva (MODO GENERADOR) para streaming web"""
        yield {"status": "info", "message": "Iniciando proceso de actualización...", "progress": 0}

        try:
            # Usar archivo por defecto si no se especifica
            if not ruta_archivo:
                raise ValueError("Se debe especificar un archivo para procesar")
            
            yield {"status": "info", "message": f"Cargando archivo: {os.path.basename(ruta_archivo)}", "progress": 5}
            
            # Cargar y validar datos
            df = self.cargar_archivo(ruta_archivo)
            if not self.validar_datos(df):
                yield {"status": "error", "message": "El archivo no es válido (faltan columnas)"}
                return
            
            self.resultados["total"] = len(df)
            yield {"status": "info", "message": f"Se encontraron {len(df)} estudiantes para actualizar", "progress": 10}
            
            # Obtener token
            if not self.obtener_token():
                yield {"status": "error", "message": "No se pudo obtener token de M365"}
                return
            
            # Procesar actualizaciones
            total = len(df)
            
            for index, estudiante in df.iterrows():
                try:
                    # Progreso base 15% hasta 95%
                    progress = 15 + int((index / total) * 80)
                    
                    yield {"status": "process", "message": f"Procesando: {estudiante['CODIGO']}", "progress": progress}
                    
                    if self.actualizar_estudiante(estudiante):
                        self.resultados["actualizados"] += 1
                        yield {"status": "log", "message": f"✅ Estudiante actualizado: {estudiante['CODIGO']}"}
                    else:
                        self.resultados["errores"] += 1
                        yield {"status": "log", "message": f"❌ Error actualizando {estudiante['CODIGO']}"}
                        
                except Exception as e:
                    error_msg = f"Error procesando {estudiante.get('CODIGO', 'desconocido')}: {e}"
                    self.resultados["detalles_errores"].append(error_msg)
                    self.resultados["errores"] += 1
                    yield {"status": "log", "message": f"❌ {error_msg}"}
            
            # Finalizar
            yield {"status": "info", "message": "Guardando logs...", "progress": 98}
            self.guardar_log()
            
            yield {"status": "complete", "message": "Proceso finalizado", "progress": 100, "results": self.resultados}
            
        except Exception as e:
             yield {"status": "error", "message": f"Error general del proceso: {e}"}

    def procesar_actualizaciones(self, ruta_archivo: str = None, confirmacion: bool = True) -> dict:
        """Wrapper para mantener compatibilidad con modo consola consumiendo el generador"""
        # Si es modo web directo (sin stream) o consola
        gen = self.ejecutar(ruta_archivo)
        for update in gen:
            if update['status'] == 'log':
                print(update['message'])
            elif update['status'] == 'info':
                 print(f"INFO: {update['message']}")
            elif update['status'] == 'error':
                 print(f"ERROR: {update['message']}")
        return self.resultados

    def mostrar_resumen(self):
        """Muestra resumen de la operación"""
        print("\n" + "="*60)
        print("RESUMEN DE ACTUALIZACIONES")
        print("="*60)
        print(f"Colegio: {config.COLEGIO_NOMBRE}")
        print(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Total procesados: {self.resultados['total']}")
        print(f"Estudiantes actualizados: {self.resultados['actualizados']}")
        print(f"Errores: {self.resultados['errores']}")
        
        if self.resultados['errores'] > 0:
            print(f"\nDetalles de errores guardados en: {config.CARPETA_LOGS}")
        
        print("="*60)

    def guardar_log(self):
        """Guarda log del proceso"""
        try:
            # Crear carpeta de logs si no existe
            os.makedirs(config.CARPETA_LOGS, exist_ok=True)
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            log_file = os.path.join(config.CARPETA_LOGS, f'actualizacion_estudiantes_{timestamp}.log')
            
            with open(log_file, 'w', encoding='utf-8') as f:
                f.write(f"ACTUALIZACIÓN DE ESTUDIANTES - {config.COLEGIO_NOMBRE}\n")
                f.write(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("="*50 + "\n")
                f.write(f"Total procesados: {self.resultados['total']}\n")
                f.write(f"Estudiantes actualizados: {self.resultados['actualizados']}\n")
                f.write(f"Errores: {self.resultados['errores']}\n\n")
                
                if self.resultados['detalles_errores']:
                    f.write("DETALLES DE ERRORES:\n")
                    f.write("-"*30 + "\n")
                    for error in self.resultados['detalles_errores']:
                        f.write(f"- {error}\n")
            
            print(f"Log guardado en: {log_file}")
            
        except Exception as e:
            print(f"Error guardando log: {e}")

def main():
    """Función principal"""
    print("ACTUALIZADOR DE ESTUDIANTES MICROSOFT 365")
    print(f"Colegio: {config.COLEGIO_NOMBRE}")
    print("="*50)
    
    try:
        actualizador = ActualizadorEstudiantes()
        
        # Solicitar ruta del archivo
        ruta_archivo = input("Ruta del archivo: ").strip()
        
        if ruta_archivo:
             actualizador.procesar_actualizaciones(ruta_archivo)
        else:
             print("Debes especificar un archivo")
            
    except KeyboardInterrupt:
        print("\nProceso interrumpido por el usuario")
    except Exception as e:
        print(f"Error inesperado: {e}")

if __name__ == "__main__":
    main()