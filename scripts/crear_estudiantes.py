

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

        # Datos del estudiante usando configuración
        user_data = {
            "accountEnabled": True,
            "displayName": f"Estudiante - {estudiante['CURSO']}: {estudiante['NOMBRES']} {estudiante['APELLIDOS']}",
            "mailNickname": estudiante["CODIGO"],
            "userPrincipalName": f"{estudiante['CODIGO']}@{config.COLEGIO_DOMINIO}",
            "passwordProfile": {
                "forceChangePasswordNextSignIn": True,
                "password": "TempPass2024!"  # Contraseña temporal
            },
            "givenName": estudiante["NOMBRES"],
            "surname": estudiante["APELLIDOS"],
            "jobTitle": estudiante["CURSO"],
            "department": config.DEFAULT_DEPARTMENT,
            "usageLocation": config.DEFAULT_USAGE_LOCATION,
            "city": "Bogotá"
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
    
    def validar_datos(self, df: pd.DataFrame) -> bool:
        """Valida que el DataFrame tenga las columnas necesarias"""
        columnas_requeridas = ["CODIGO", "DOCUMENTO", "GRADO", "CURSO", "APELLIDOS", "NOMBRES"]
        columnas_faltantes = [col for col in columnas_requeridas if col not in df.columns]
        
        if columnas_faltantes:
            print(f"❌ Faltan columnas requeridas: {columnas_faltantes}")
            return False
        
        print("✅ Datos válidos")
        return True

    def procesar_estudiantes(self, ruta_archivo: str = None, confirmacion: bool = True) -> dict:
        """Procesa la creación masiva de estudiantes
        
        Args:
            ruta_archivo (str, optional): Ruta al archivo a procesar. Defaults to None.
            confirmacion (bool, optional): Si True, pide confirmación por consola. Si False, ejecuta directamente. Defaults to True.
            
        Returns:
            dict: Resultados del proceso
        """
        try:
            # Usar archivo por defecto si no se especifica
            if not ruta_archivo:
                ruta_archivo = config.ARCHIVO_NUEVOS
            
            print(f"🏫 Colegio: {config.COLEGIO_NOMBRE}")
            print(f"📁 Procesando archivo: {ruta_archivo}")
            print("="*50)
            
            # Cargar y validar datos
            df = self.cargar_archivo(ruta_archivo)
            if not self.validar_datos(df):
                return self.resultados
            
            self.resultados["total"] = len(df)
            
            # Mostrar vista previa
            print("\n📋 Vista previa de estudiantes:")
            print(df[["CODIGO", "DOCUMENTO", "GRADO", "CURSO", "APELLIDOS", "NOMBRES"]].head())
            
            if confirmacion:
                # Confirmación
                respuesta = input(f"\n¿Crear {len(df)} estudiantes en {config.COLEGIO_NOMBRE}? (si/no): ").lower()
                if respuesta not in ['si', 's', 'yes', 'y']:
                    print("❌ Operación cancelada")
                    return self.resultados
            
            # Obtener token
            if not self.obtener_token():
                return self.resultados
            
            # Procesar estudiantes
            print(f"\n🚀 Iniciando creación de {len(df)} estudiantes...")
            print("="*50)
            
            for index, estudiante in df.iterrows():
                try:
                    print(f"\n📝 Procesando {index + 1}/{len(df)}: {estudiante['CODIGO']}")
                    
                    # Crear estudiante
                    if self.crear_estudiante(estudiante):
                        self.resultados["creados"] += 1
                        
                        # Asignar licencia
                        if self.asignar_licencia(estudiante['CODIGO']):
                            self.resultados["licenciados"] += 1
                    else:
                        self.resultados["errores"] += 1
                        
                except Exception as e:
                    error_msg = f"Error procesando {estudiante.get('CODIGO', 'desconocido')}: {e}"
                    print(f"❌ {error_msg}")
                    self.resultados["detalles_errores"].append(error_msg)
                    self.resultados["errores"] += 1
            
            # Mostrar resumen
            self.mostrar_resumen()
            self.guardar_log()
            
            return self.resultados
            
        except Exception as e:
            print(f"❌ Error general: {e}")
            return self.resultados

    def mostrar_resumen(self):
        """Muestra resumen de la operación"""
        print("\n" + "="*60)
        print("📊 RESUMEN DEL PROCESO")
        print("="*60)
        print(f"🏫 Colegio: {config.COLEGIO_NOMBRE}")
        print(f"📅 Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"📊 Total procesados: {self.resultados['total']}")
        print(f"✅ Estudiantes creados: {self.resultados['creados']}")
        print(f"🎯 Licencias asignadas: {self.resultados['licenciados']}")
        print(f"❌ Errores: {self.resultados['errores']}")
        
        if self.resultados['errores'] > 0:
            print(f"\n📝 Detalles de errores guardados en: {config.CARPETA_LOGS}")
        
        print("="*60)

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
        
        # Usar archivo por defecto o solicitar ruta
        usar_default = input(f"\n¿Usar archivo por defecto '{config.ARCHIVO_NUEVOS}'? (si/no): ").lower()
        
        if usar_default in ['si', 's', 'yes', 'y']:
            creador.procesar_estudiantes()
        else:
            ruta_archivo = input("📁 Ruta del archivo: ").strip()
            creador.procesar_estudiantes(ruta_archivo)
            
    except KeyboardInterrupt:
        print("\n❌ Proceso interrumpido por el usuario")
    except Exception as e:
        print(f"❌ Error inesperado: {e}")

if __name__ == "__main__":
    main()