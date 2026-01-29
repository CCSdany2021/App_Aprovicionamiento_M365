"""
Módulo para ELIMINAR TEAMS de forma controlada
Reemplaza la lógica PowerShell con interfaz web segura

Flujo:
1. Carga Excel con Teams a eliminar (GroupId o DisplayName)
2. Busca cada Team en el tenant
3. Muestra lista de confirmación
4. Elimina teams uno a uno
5. Registra logs detallados de cada operación
"""

import pandas as pd
import requests
import urllib3
from datetime import datetime
import os
import sys

# Añadir la carpeta scripts al path
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class EliminadorTeams:
    """Clase para eliminar Teams del tenant de forma controlada"""
    
    def __init__(self):
        """Inicializa el eliminador de Teams"""
        try:
            config.validar_configuracion()
        except:
            pass
        
        self.token = None
        self.resultados = {
            "total": 0,
            "encontrados": 0,
            "eliminados": 0,
            "no_encontrados": 0,
            "errores": 0,
            "detalles": [],
            "equipos_a_eliminar": [],
            "equipos_eliminados": [],
            "equipos_errores": []
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

    def cargar_archivo(self, ruta_archivo: str) -> pd.DataFrame:
        """
        Carga los Teams desde archivo Excel o CSV
        
        Soporta columnas:
        - GroupId (ID directo del grupo)
        - DisplayName (nombre del equipo)
        - Cualquier combinación
        """
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
            
            print(f"✅ Archivo cargado: {len(df)} Teams encontrados")
            return df
            
        except Exception as e:
            raise Exception(f"❌ Error leyendo archivo: {e}")

    def detectar_columna_identificador(self, df: pd.DataFrame) -> str:
        """
        Detecta automáticamente la columna que contiene identificadores
        
        Busca: GroupId, Id, DisplayName, Name, Equipo, etc.
        """
        posibles_columnas = [
            'GroupId', 'Id', 'GROUP_ID', 'GROUPID',
            'DisplayName', 'Name', 'Equipo', 'Team',
            'id', 'groupid', 'displayname', 'name', 'equipo', 'team'
        ]
        
        for col in posibles_columnas:
            if col in df.columns:
                print(f"✅ Columna detectada: '{col}'")
                return col
        
        # Si no encuentra, usa la primera columna
        primera_columna = df.columns[0]
        print(f"⚠️  Usando primera columna: '{primera_columna}'")
        return primera_columna

    def buscar_team(self, identificador: str) -> dict or None:
        """
        Busca un Team por GroupId o DisplayName
        
        Returns:
            dict: {GroupId, DisplayName, Mail} o None si no encuentra
        """
        if not self.token:
            return None
        
        identificador = str(identificador).strip()
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        # Primero intentar como GroupId (ID directo)
        try:
            url = f"{config.GRAPH_ENDPOINT}/groups/{identificador}?$select=id,displayName,mail"
            response = requests.get(url, headers=headers, verify=False)
            
            if response.status_code == 200:
                data = response.json()
                return {
                    "GroupId": data.get("id"),
                    "DisplayName": data.get("displayName"),
                    "Mail": data.get("mail")
                }
        except:
            pass
        
        # Si no funciona como ID, buscar por DisplayName
        try:
            # Escapar comillas
            nombre_escapado = identificador.replace("'", "''")
            url = (
                f"{config.GRAPH_ENDPOINT}/groups?"
                f"$filter=displayName eq '{nombre_escapado}' "
                f"or mail eq '{nombre_escapado}'"
                f"&$select=id,displayName,mail"
            )
            response = requests.get(url, headers=headers, verify=False)
            
            if response.status_code == 200:
                data = response.json()
                if data.get('value'):
                    equipo = data['value'][0]
                    return {
                        "GroupId": equipo.get("id"),
                        "DisplayName": equipo.get("displayName"),
                        "Mail": equipo.get("mail")
                    }
        except:
            pass
        
        return None

    def obtener_lista_equipos_a_eliminar(self, ruta_archivo: str) -> list:
        """
        Carga archivo y retorna lista de Teams a eliminar con sus detalles
        
        Returns:
            list: [{GroupId, DisplayName, Mail, Identificador}, ...]
        """
        # Cargar archivo
        df = self.cargar_archivo(ruta_archivo)
        col_identificador = self.detectar_columna_identificador(df)
        
        # PREVIEW DE SEGURIDAD
        print(f"\n🔍 COLUMNA DETECTADA PARA BÚSQUEDA: '{col_identificador}'")
        first_3_values = df[col_identificador].head(3).tolist()
        print(f"📊 Primeros valores a buscar: {first_3_values}")
        print("   (Verifique que estos sean los IDs o Nombres correctos de los equipos)")

        # Obtener token
        if not self.obtener_token():
            raise Exception("No se pudo obtener token de acceso")
        
        equipos_a_eliminar = []
        
        print("\n🔍 Buscando Teams en el tenant (BÚSQUEDA EXACTA POR ARCHIVO)...")
        print("=" * 70)
        
        for idx, identificador in enumerate(df[col_identificador], 1):
            identificador = str(identificador).strip()
            if not identificador or identificador.lower() == 'nan':
                continue
            
            print(f"\n[{idx}] Buscando: {identificador}")
            
            # Buscar el Team
            team = self.buscar_team(identificador)
            
            if team:
                equipos_a_eliminar.append({
                    "Identificador": identificador,
                    "GroupId": team["GroupId"],
                    "DisplayName": team["DisplayName"],
                    "Mail": team["Mail"],
                    "Status": "Encontrado"
                })
                print(f"    ✅ Encontrado: {team['DisplayName']}")
                print(f"       Mail: {team['Mail']}")
                self.resultados["encontrados"] += 1
            else:
                equipos_a_eliminar.append({
                    "Identificador": identificador,
                    "GroupId": None,
                    "DisplayName": None,
                    "Mail": None,
                    "Status": "No encontrado"
                })
                print(f"    ⚠️  No encontrado")
                self.resultados["no_encontrados"] += 1
            
            self.resultados["equipos_a_eliminar"].append(equipos_a_eliminar[-1])
        
        self.resultados["total"] = len(df[col_identificador].dropna())
        
        print("\n" + "=" * 70)
        print(f"✅ Se encontraron {self.resultados['encontrados']} de {self.resultados['total']} Teams")
        print(f"⚠️  No encontrados: {self.resultados['no_encontrados']}")
        
        return equipos_a_eliminar

    def eliminar_team(self, group_id: str, display_name: str) -> tuple:
        """
        Elimina un Team del tenant
        
        Returns:
            (éxito: bool, mensaje: str)
        """
        if not self.token:
            return False, "Token no disponible"
        
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        url = f"{config.GRAPH_ENDPOINT}/groups/{group_id}"
        
        try:
            response = requests.delete(url, headers=headers, verify=False)
            
            if response.status_code == 204:
                return True, f"Team '{display_name}' eliminado correctamente"
            elif response.status_code == 404:
                return False, f"Team no encontrado"
            else:
                return False, f"Error {response.status_code}: {response.text[:100]}"
        
        except requests.RequestException as e:
            return False, f"Error de conexión: {str(e)}"

    def procesar_equipos(self, equipos: list, confirmacion: bool = True) -> dict:
        """
        Procesa la eliminación de Teams
        
        Args:
            equipos: Lista de dicts con datos de Teams a eliminar
            confirmacion: Si True, pide confirmación antes de eliminar
        """
        
        # Filtrar solo equipos encontrados
        equipos_encontrados = [e for e in equipos if e["Status"] == "Encontrado"]
        
        if not equipos_encontrados:
            print("❌ No hay Teams para eliminar")
            return self.resultados
        
        print("\n" + "=" * 70)
        print("📋 TEAMS A ELIMINAR:")
        print("=" * 70)
        for idx, equipo in enumerate(equipos_encontrados, 1):
            print(f"{idx}. {equipo['DisplayName']}")
            print(f"   Email: {equipo['Mail']}")
            print(f"   ID: {equipo['GroupId']}")
        
        print("\n" + "=" * 70)
        
        if confirmacion:
            # Confirmación de seguridad
            print("\n⚠️  ADVERTENCIA:")
            print("Esta operación eliminará PERMANENTEMENTE los Teams")
            print("NO se pueden recuperar una vez eliminados")
            
            respuesta = input(
                f"\n¿Está seguro de eliminar {len(equipos_encontrados)} Teams? "
                "(escriba 'ELIMINAR' para confirmar): "
            ).strip()
            
            if respuesta != "ELIMINAR":
                print("❌ Operación cancelada")
                return self.resultados
        
        # Eliminar Teams
        print("\n" + "=" * 70)
        print("🗑️  INICIANDO ELIMINACIÓN DE TEAMS...")
        print("=" * 70)
        
        for idx, equipo in enumerate(equipos_encontrados, 1):
            print(f"\n[{idx}/{len(equipos_encontrados)}] Eliminando: {equipo['DisplayName']}")
            
            exito, mensaje = self.eliminar_team(equipo["GroupId"], equipo["DisplayName"])
            
            if exito:
                print(f"✅ {mensaje}")
                self.resultados["eliminados"] += 1
                self.resultados["equipos_eliminados"].append(equipo)
            else:
                print(f"❌ {mensaje}")
                self.resultados["errores"] += 1
                self.resultados["equipos_errores"].append({
                    **equipo,
                    "error": mensaje
                })
        
        # Resumen final
        self.mostrar_resumen()
        self.guardar_log()
        
        return self.resultados

    def mostrar_resumen(self):
        """Muestra resumen de la operación"""
        print("\n" + "=" * 70)
        print("📊 RESUMEN DE ELIMINACIÓN DE TEAMS")
        print("=" * 70)
        print(f"Total en archivo: {self.resultados['total']}")
        print(f"Teams encontrados: {self.resultados['encontrados']}")
        print(f"Teams eliminados: {self.resultados['eliminados']}")
        print(f"Teams no encontrados: {self.resultados['no_encontrados']}")
        print(f"Errores: {self.resultados['errores']}")
        print("=" * 70)

    def guardar_log(self):
        """Guarda log detallado de la operación"""
        try:
            os.makedirs(config.CARPETA_LOGS, exist_ok=True)
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            log_file = os.path.join(config.CARPETA_LOGS, f'eliminacion_teams_{timestamp}.log')
            
            with open(log_file, 'w', encoding='utf-8') as f:
                f.write("ELIMINACIÓN DE TEAMS - " + config.COLEGIO_NOMBRE + "\n")
                f.write(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("=" * 70 + "\n\n")
                
                f.write("RESUMEN:\n")
                f.write(f"Total en archivo: {self.resultados['total']}\n")
                f.write(f"Teams encontrados: {self.resultados['encontrados']}\n")
                f.write(f"Teams eliminados: {self.resultados['eliminados']}\n")
                f.write(f"Teams no encontrados: {self.resultados['no_encontrados']}\n")
                f.write(f"Errores: {self.resultados['errores']}\n\n")
                
                f.write("TEAMS ELIMINADOS:\n")
                f.write("-" * 70 + "\n")
                for equipo in self.resultados['equipos_eliminados']:
                    f.write(f"✅ {equipo['DisplayName']}\n")
                    f.write(f"   Email: {equipo['Mail']}\n")
                    f.write(f"   ID: {equipo['GroupId']}\n\n")
                
                if self.resultados['equipos_errores']:
                    f.write("ERRORES:\n")
                    f.write("-" * 70 + "\n")
                    for equipo in self.resultados['equipos_errores']:
                        f.write(f"❌ {equipo['DisplayName'] or equipo['Identificador']}\n")
                        f.write(f"   Error: {equipo.get('error', 'Desconocido')}\n\n")
            
            print(f"\n📝 Log guardado en: {log_file}")
            
        except Exception as e:
            print(f"❌ Error guardando log: {e}")

    def procesar(self, ruta_archivo: str, confirmacion: bool = True) -> dict:
        """
        Proceso principal: carga, busca y elimina Teams
        
        Args:
            ruta_archivo: Ruta al archivo Excel/CSV
            confirmacion: Si pedir confirmación antes de eliminar
            
        Returns:
            dict: Resultados del proceso
        """
        print("\n" + "=" * 70)
        print("🏫 ELIMINADOR DE TEAMS - " + config.COLEGIO_NOMBRE)
        print("=" * 70)
        
        try:
            # 1. Obtener lista de Teams a eliminar
            equipos = self.obtener_lista_equipos_a_eliminar(ruta_archivo)
            
            # 2. Procesar eliminación
            self.procesar_equipos(equipos, confirmacion=confirmacion)
            
            return self.resultados
            
        except Exception as e:
            print(f"❌ Error general: {e}")
            self.resultados["errores"] += 1
            self.resultados["detalles"].append(str(e))
            return self.resultados


    def ejecutar(self, ruta_archivo: str, confirmacion: bool = False):
        """Procesa la eliminación de Teams (MODO GENERADOR) para streaming web"""
        yield {"status": "info", "message": "Iniciando proceso de eliminación de Teams...", "progress": 0}

        try:
             # 1. Obtener lista de Teams a eliminar
            yield {"status": "info", "message": f"Cargando y analizando archivo: {os.path.basename(ruta_archivo)}", "progress": 5}
            
            # Cargamos DF y detectamos columna (lógica interna de obtener_lista_equipos_a_eliminar adaptada)
            df = self.cargar_archivo(ruta_archivo)
            col_identificador = self.detectar_columna_identificador(df)
            
            # Obtener token
            if not self.obtener_token():
                yield {"status": "error", "message": "No se pudo obtener token de M365"}
                return

            equipos_a_eliminar = []
            yield {"status": "info", "message": "Buscando Teams en el tenant...", "progress": 10}
            
            total_busqueda = len(df[col_identificador].dropna())
            
            # Fase 1: Búsqueda
            for idx, identificador in enumerate(df[col_identificador], 1):
                identificador = str(identificador).strip()
                if not identificador:
                    continue
                
                # Progreso de búsqueda (10% a 40%)
                prog_busq = 10 + int((idx / total_busqueda) * 30)
                yield {"status": "process", "message": f"Buscando: {identificador}", "progress": prog_busq}
                
                team = self.buscar_team(identificador)
                
                if team:
                    equipos_a_eliminar.append({
                        "Identificador": identificador,
                        "GroupId": team["GroupId"],
                        "DisplayName": team["DisplayName"],
                        "Mail": team["Mail"],
                        "Status": "Encontrado"
                    })
                    self.resultados["encontrados"] += 1
                    yield {"status": "log", "message": f"✅ Encontrado: {team['DisplayName']}"}
                else:
                     self.resultados["no_encontrados"] += 1
                     yield {"status": "warning", "message": f"⚠️ No encontrado: {identificador}"}

            self.resultados["total"] = total_busqueda
            
            # Filtrar encontrados
            equipos_encontrados = [e for e in equipos_a_eliminar if e.get("Status") == "Encontrado"]
            
            if not equipos_encontrados:
                 yield {"status": "error", "message": "No se encontraron Teams válidos para eliminar"}
                 return

            total_eliminar = len(equipos_encontrados)
            yield {"status": "info", "message": f"Eliminando {total_eliminar} Teams encontrados...", "progress": 40}
            
            # Fase 2: Eliminación
            for idx, equipo in enumerate(equipos_encontrados, 1):
                # Progreso eliminación (40% a 95%)
                prog_elim = 40 + int((idx / total_eliminar) * 55)
                yield {"status": "process", "message": f"Eliminando: {equipo['DisplayName']}", "progress": prog_elim}
                
                exito, mensaje = self.eliminar_team(equipo["GroupId"], equipo["DisplayName"])
                
                if exito:
                    self.resultados["eliminados"] += 1
                    self.resultados["equipos_eliminados"].append(equipo)
                    yield {"status": "log", "message": f"✅ {mensaje}"}
                else:
                    self.resultados["errores"] += 1
                    self.resultados["equipos_errores"].append({**equipo, "error": mensaje})
                    yield {"status": "log", "message": f"❌ {mensaje}"}
            
            # Finalizar
            yield {"status": "info", "message": "Guardando logs...", "progress": 98}
            self.guardar_log()
            
            yield {"status": "complete", "message": "Proceso finalizado", "progress": 100, "results": self.resultados}

        except Exception as e:
            yield {"status": "error", "message": f"Error general: {e}"}

def main():
    """Función principal para pruebas"""
    print("🗑️  ELIMINADOR DE TEAMS MICROSOFT 365")
    print(f"🏫 {config.COLEGIO_NOMBRE}")
    print("=" * 70)
    
    try:
        eliminador = EliminadorTeams()
        
        # Solicitar ruta del archivo
        ruta_archivo = input("\n📁 Ruta del archivo (Excel/CSV): ").strip()
        
        if not os.path.exists(ruta_archivo):
            print(f"❌ Archivo no encontrado: {ruta_archivo}")
            return
        
        # Procesar
        eliminador.procesar(ruta_archivo, confirmacion=True)
        
    except KeyboardInterrupt:
        print("\n❌ Proceso interrumpido por el usuario")
    except Exception as e:
        print(f"❌ Error inesperado: {e}")

if __name__ == "__main__":
    main()