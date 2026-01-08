

import pandas as pd
import requests
import json
import urllib3
from datetime import datetime
import os
import sys

# Añadir la carpeta scripts al path para importar configuración
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config
from scripts.notificador_email import NotificadorEmail

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class CreadorEstudiantes:
    """Clase simplificada para crear estudiantes en Microsoft 365"""
    
    def __init__(self):
        # Validar configuración al inicializar
        config.validar_configuracion()
        self.token = None
        self.resultados = {
            "total": 0,
            "creados": 0,
            "licenciados": 0,
            "errores": 0,
            "correos_enviados": 0,
            "detalles_errores": []
        }
        self.notificador = NotificadorEmail()
        
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

    def crear_estudiante(self, estudiante: dict) -> bool:
        """Crea un estudiante individual en Microsoft 365"""
        if not self.token:
            print("❌ Token no disponible")
            return False
            
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }

        password_temporal = "TempPass2025!" # Podrías generar una dinámica si quieres
        
        # Datos del estudiante usando configuración
        user_data = {
            "accountEnabled": True,
            "displayName": f"Estudiante - {estudiante['CURSO']}: {estudiante['APELLIDOS']} {estudiante['NOMBRES']}",
            "mailNickname": estudiante["CODIGO"],
            "userPrincipalName": f"{estudiante['CODIGO']}@{config.COLEGIO_DOMINIO}",
            "passwordProfile": {
                "forceChangePasswordNextSignIn": True,
                "password": password_temporal
            },
            "givenName": estudiante["NOMBRES"],
            "surname": estudiante["APELLIDOS"],
            "jobTitle": estudiante["CURSO"],
            "department": config.DEFAULT_DEPARTMENT,
            "usageLocation": config.DEFAULT_USAGE_LOCATION,
            "city": config.DEFAULT_CITY
        }

        try:
            response = requests.post(
                f"{config.GRAPH_ENDPOINT}/users", 
                headers=headers, 
                json=user_data,
                verify=False
            )
            
            if response.status_code == 201:
                print(f"✅ Estudiante creado: {estudiante['CODIGO']}")
                return True
            else:
                error_msg = f"Error creando {estudiante['CODIGO']}: {response.text}"
                print(f"❌ {error_msg}")
                self.resultados["detalles_errores"].append(error_msg)
                return False
                
        except requests.RequestException as e:
            error_msg = f"Error de conexión creando {estudiante['CODIGO']}: {e}"
            print(f"❌ {error_msg}")
            self.resultados["detalles_errores"].append(error_msg)
            return False

    def asignar_licencia(self, codigo_estudiante: str) -> bool:
        """Asigna licencia A1 al estudiante"""
        if not self.token:
            return False
            
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }

        data = {
            "addLicenses": [{"skuId": config.LICENSE_STUDENT}],
            "removeLicenses": []
        }

        user_email = f"{codigo_estudiante}@{config.COLEGIO_DOMINIO}"
        url = f"{config.GRAPH_ENDPOINT}/users/{user_email}/assignLicense"
        
        try:
            response = requests.post(url, headers=headers, json=data, verify=False)
            if response.status_code == 200:
                print(f"✅ Licencia asignada a {codigo_estudiante}")
                return True
            else:
                print(f"❌ Error asignando licencia a {codigo_estudiante}: {response.text}")
                return False
                
        except requests.RequestException as e:
            print(f"❌ Error de conexión asignando licencia a {codigo_estudiante}: {e}")
            return False

    def cargar_archivo(self, ruta_archivo: str) -> pd.DataFrame:
        """Carga estudiantes desde archivo Excel o CSV"""
        try:
            if ruta_archivo.endswith(".xlsx"):
                df = pd.read_excel(ruta_archivo, dtype=str)
            elif ruta_archivo.endswith(".csv"):
                df = pd.read_csv(ruta_archivo, dtype=str, encoding="utf-8")
            else:
                raise ValueError("❌ Formato no soportado. Usa .xlsx o .csv")
            
            # Limpiar datos
            df.columns = df.columns.str.strip()
            df = df.fillna("")
            
            print(f"✅ Archivo cargado: {len(df)} estudiantes encontrados")
            return df
            
        except Exception as e:
            raise Exception(f"❌ Error leyendo archivo: {e}")
    
    def detectar_columnas(self, df: pd.DataFrame) -> dict:
        """Detecta las columnas necesarias de forma flexible"""
        mapeo = {
            "CODIGO": ["CODIGO", "Codigo", "Cod", "ID"],
            "GRADO": ["GRADO", "Grado"],
            "CURSO": ["CURSO", "Curso"],
            "APELLIDOS": ["APELLIDOS", "Apellidos", "Apellido"],
            "NOMBRES": ["NOMBRES", "Nombres", "Nombre"],
            "CORREO_PERSONAL": ["CORREO_PERSONAL", "CORREO_PEROSNAL", "Email_Personal", "Correo_Personal", "Correo", "Email"]
        }
        
        columnas_encontradas = {}
        for clave, opciones in mapeo.items():
            for opcion in opciones:
                # Búsqueda insensible a mayúsculas/minúsculas y espacios
                for col in df.columns:
                    if col.strip().upper() == opcion.upper():
                        columnas_encontradas[clave] = col
                        break
                if clave in columnas_encontradas:
                    break
        
        return columnas_encontradas

    def validar_datos(self, df: pd.DataFrame) -> bool:
        """Valida que el DataFrame tenga las columnas necesarias detectadas"""
        columnas_detectadas = self.detectar_columnas(df)
        columnas_requeridas = ["CODIGO", "GRADO", "CURSO", "APELLIDOS", "NOMBRES", "CORREO_PERSONAL"]
        
        faltantes = [col for col in columnas_requeridas if col not in columnas_detectadas]
        
        if faltantes:
            print(f"❌ No se pudieron detectar las columnas: {faltantes}")
            print(f"Columnas disponibles en el archivo: {list(df.columns)}")
            return False
        
        print(f"✅ Columnas detectadas correctamente: {list(columnas_detectadas.values())}")
        return True

    def ejecutar(self, ruta_archivo: str = None) -> dict:
        """Procesa la creación masiva de estudiantes (MODO GENERADOR)"""
        yield {"status": "info", "message": "Iniciando proceso de creación...", "progress": 0}
        
        try:
            # Usar archivo por defecto si no se especifica
            if not ruta_archivo:
                raise ValueError("Se debe especificar un archivo para procesar")

            
            yield {"status": "info", "message": f"Cargando archivo: {os.path.basename(ruta_archivo)}", "progress": 5}
            
            # Cargar y detectar columnas
            df = self.cargar_archivo(ruta_archivo)
            if not self.validar_datos(df):
                yield {"status": "error", "message": "El archivo no tiene las columnas necesarias"}
                return
            
            columnas = self.detectar_columnas(df)
            self.resultados["total"] = len(df)
            
            yield {"status": "info", "message": f"Se encontraron {len(df)} estudiantes para procesar", "progress": 10}
            
            # Obtener token
            if not self.obtener_token():
                yield {"status": "error", "message": "No se pudo obtener el token de Microsoft Graph"}
                return
            
            # Pasar token al notificador
            self.notificador.token = self.token
            
            # Procesar estudiantes
            total = len(df)
            
            for index, row in df.iterrows():
                try:
                    # Progreso base 15% hasta 95%
                    progress = 15 + int((index / total) * 80)
                    
                    # Mapear datos
                    estudiante = {clave: str(row[col_nombre]).strip() for clave, col_nombre in columnas.items()}
                    nombre_completo = f"{estudiante['NOMBRES']} {estudiante['APELLIDOS']}"
                    
                    yield {"status": "process", "message": f"Procesando: {estudiante['CODIGO']} - {nombre_completo}", "progress": progress}
                    
                    # Crear estudiante
                    if self.crear_estudiante(estudiante):
                        self.resultados["creados"] += 1
                        msg_log = f"✅ Usuario creado: {estudiante['CODIGO']}"
                        yield {"status": "log", "message": msg_log}
                        
                        # Asignar licencia
                        if self.asignar_licencia(estudiante['CODIGO']):
                            self.resultados["licenciados"] += 1
                            msg_lic = f"   🏷️ Licencia asignada"
                            yield {"status": "log", "message": msg_lic}
                            
                            # Enviar correo
                            upn = f"{estudiante['CODIGO']}@{config.COLEGIO_DOMINIO}"
                            correo_personal = estudiante['CORREO_PERSONAL']
                            
                            if self.notificador.enviar_credenciales(correo_personal, nombre_completo, upn, "TempPass2025!"):
                                self.resultados["correos_enviados"] += 1
                                yield {"status": "log", "message": f"   📧 Correo enviado a {correo_personal}"}
                            else:
                                yield {"status": "log", "message": f"   ⚠️ Correo NO enviado (Verificar .env)"}
                    else:
                        self.resultados["errores"] += 1
                        yield {"status": "log", "message": f"❌ Error creando {estudiante['CODIGO']}"}
                        
                except Exception as e:
                    error_msg = f"Error procesando {estudiante.get('CODIGO', 'desconocido')}: {e}"
                    self.resultados["detalles_errores"].append(error_msg)
                    self.resultados["errores"] += 1
                    yield {"status": "log", "message": f"❌ {error_msg}"}
            
            # Guardar log y finalizar
            yield {"status": "info", "message": "Guardando logs...", "progress": 98}
            self.guardar_log()
            
            yield {"status": "complete", "message": "Proceso finalizado", "progress": 100, "results": self.resultados}
            
        except Exception as e:
            yield {"status": "error", "message": f"Error general: {e}"}
            return

    def procesar_estudiantes(self, ruta_archivo: str = None, confirmacion: bool = True) -> dict:
        """Wrapper para mantener compatibilidad con modo consola si es necesario, consumiendo el generador"""
        gen = self.ejecutar(ruta_archivo)
        for update in gen:
            if update['status'] == 'log':
                print(update['message'])
            elif update['status'] == 'complete':
                return update['results']
            elif update['status'] == 'error':
                print(f"ERROR: {update['message']}")
        return self.resultados

    def guardar_log(self):
        """Guarda log del proceso"""
        try:
            # Crear carpeta de logs si no existe
            os.makedirs(config.CARPETA_LOGS, exist_ok=True)
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            log_file = os.path.join(config.CARPETA_LOGS, f'creacion_estudiantes_{timestamp}.log')
            
            with open(log_file, 'w', encoding='utf-8') as f:
                f.write(f"CREACIÓN DE ESTUDIANTES - {config.COLEGIO_NOMBRE}\n")
                f.write(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("="*50 + "\n")
                f.write(f"Total procesados: {self.resultados['total']}\n")
                f.write(f"Estudiantes creados: {self.resultados['creados']}\n")
                f.write(f"Licencias asignadas: {self.resultados['licenciados']}\n")
                f.write(f"Errores: {self.resultados['errores']}\n\n")
                
                if self.resultados['detalles_errores']:
                    f.write("DETALLES DE ERRORES:\n")
                    f.write("-"*30 + "\n")
                    for error in self.resultados['detalles_errores']:
                        f.write(f"- {error}\n")
            
            print(f"📝 Log guardado en: {log_file}")
            
        except Exception as e:
            print(f"❌ Error guardando log: {e}")

def main():
    """Función principal"""
    print("🎓 CREADOR DE ESTUDIANTES MICROSOFT 365")
    print(f"🏫 {config.COLEGIO_NOMBRE}")
    print("="*50)
    
    try:
        creador = CreadorEstudiantes()
        
        # Solicitar ruta del archivo
        ruta_archivo = input("📁 Ruta del archivo: ").strip()
        if ruta_archivo:
             creador.procesar_estudiantes(ruta_archivo)
        else:
             print("❌ Debes especificar un archivo")

            
    except KeyboardInterrupt:
        print("\n❌ Proceso interrumpido por el usuario")
    except Exception as e:
        print(f"❌ Error inesperado: {e}")

if __name__ == "__main__":
    main()