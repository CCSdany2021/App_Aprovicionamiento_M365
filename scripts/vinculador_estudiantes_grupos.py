import pandas as pd
import requests
import urllib3
from datetime import datetime
import os
import sys

sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


class VinculadorEstudiantesGrupos:
    """Vincula estudiantes a grupos de seguridad (como PowerShell #6)"""
    
    def __init__(self):
        try:
            config.validar_configuracion()
        except:
            pass
        
        self.token = None
        self.grupos_disponibles = []  # Todos los grupos de Azure AD
        self.grupos_cache = {}
        self.usuarios_cache = {}
        
        self.resultados = {
            "total_estudiantes": 0,
            "total_grupos": 0,
            "estudiantes_vinculados": 0,
            "estudiantes_ya_en_grupo": 0,
            "estudiantes_no_encontrados": 0,
            "grupos_no_encontrados": 0,
            "errores_vinculacion": 0,
            "errores": [],
            "detalles_grupos": [],
            "estudiantes_procesados": []
        }
    
    def obtener_token(self) -> bool:
        """Obtiene token"""
        url = f"https://login.microsoftonline.com/{config.TENANT_ID}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": config.CLIENT_ID,
            "client_secret": config.CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default"
        }
        
        try:
            response = requests.post(url, data=data, verify=False, timeout=10)
            response.raise_for_status()
            self.token = response.json()["access_token"]
            print("✅ Token obtenido")
            return True
        except Exception as e:
            print(f"❌ Error de autenticación: {e}")
            self.resultados["errores"].append(f"Error de token: {str(e)}")
            return False

    def obtener_todos_los_grupos(self) -> bool:
        """
        Obtiene TODOS los grupos de distribución de Azure AD
        Equivalente a: Get-DistributionGroup -Anr 'Estudiantes Curso -'
        """
        if not self.token:
            return False
        
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        print("\n🔍 Obteniendo grupos de seguridad desde Azure AD...")
        
        try:
            # Obtener grupos que contengan "Estudiantes Curso"
            url = (
                f"{config.GRAPH_ENDPOINT}/groups?"
                f"$filter=displayName eq 'Estudiantes Curso - 101' or "
                f"displayName eq 'Estudiantes Curso - 102' or "
                f"displayName eq 'Estudiantes Curso - 103' or "
                f"displayName eq 'Estudiantes Curso - 201' or "
                f"displayName eq 'Estudiantes Curso - 202' or "
                f"displayName eq 'Estudiantes Curso - 203' or "
                f"displayName eq 'Estudiantes Curso - 301' or "
                f"displayName eq 'Estudiantes Curso - 302' or "
                f"displayName eq 'Estudiantes Curso - 303' or "
                f"displayName eq 'Estudiantes Curso - 401' or "
                f"displayName eq 'Estudiantes Curso - 402' or "
                f"displayName eq 'Estudiantes Curso - 403' or "
                f"displayName eq 'Estudiantes Curso - 501' or "
                f"displayName eq 'Estudiantes Curso - 502' or "
                f"displayName eq 'Estudiantes Curso - 601' or "
                f"displayName eq 'Estudiantes Curso - 602' or "
                f"displayName eq 'Estudiantes Curso - 603' or "
                f"displayName eq 'Estudiantes Curso - 701' or "
                f"displayName eq 'Estudiantes Curso - 702' or "
                f"displayName eq 'Estudiantes Curso - 703' or "
                f"displayName eq 'Estudiantes Curso - 801' or "
                f"displayName eq 'Estudiantes Curso - 802' or "
                f"displayName eq 'Estudiantes Curso - 803' or "
                f"displayName eq 'Estudiantes Curso - 901' or "
                f"displayName eq 'Estudiantes Curso - 902' or "
                f"displayName eq 'Estudiantes Curso - 903' or "
                f"displayName eq 'Estudiantes Curso - 1001' or "
                f"displayName eq 'Estudiantes Curso - 1002' or "
                f"displayName eq 'Estudiantes Curso - 1003' or "
                f"displayName eq 'Estudiantes Curso - 1101' or "
                f"displayName eq 'Estudiantes Curso - 1102' or "
                f"displayName eq 'Estudiantes Curso - 1103' or "
                f"displayName eq 'Estudiantes Curso - TR01' or "
                f"displayName eq 'Estudiantes Curso - JR01'"
                f"&$select=id,displayName,mail"
            )
            
            # Mejor: Obtener TODOS los grupos y filtrar en Python
            url_simple = f"{config.GRAPH_ENDPOINT}/groups?$filter=startsWith(displayName, 'Estudiantes Curso -')&$select=id,displayName,mail"
            response = requests.get(url_simple, headers=headers, verify=False, timeout=15)
            
            if response.status_code == 200:
                data = response.json()
                self.grupos_disponibles = data.get('value', [])
                
                print(f"✅ {len(self.grupos_disponibles)} grupos encontrados")
                for grupo in self.grupos_disponibles[:10]:  # Mostrar primeros 10
                    print(f"   • {grupo.get('displayName')}")
                if len(self.grupos_disponibles) > 10:
                    print(f"   ... y {len(self.grupos_disponibles) - 10} más")
                
                self.resultados["total_grupos"] = len(self.grupos_disponibles)
                return True
            else:
                print(f"❌ Error obteniendo grupos: {response.status_code}")
                return False
        
        except Exception as e:
            print(f"❌ Error: {e}")
            self.resultados["errores"].append(f"Error obteniendo grupos: {str(e)}")
            return False

    def cargar_estudiantes(self, ruta_archivo: str) -> pd.DataFrame:
        """
        Carga Excel con estudiantes
        Estructura: CODIGO_ESTUDIANTE | CURSO
        """
        try:
            if ruta_archivo.endswith(".xlsx"):
                df = pd.read_excel(ruta_archivo, dtype=str)
            elif ruta_archivo.endswith(".csv"):
                df = pd.read_csv(ruta_archivo, dtype=str, encoding="utf-8")
            else:
                raise ValueError("Formato no soportado")
            
            df.columns = df.columns.str.strip()
            df = df.fillna("")
            
            print(f"✅ {len(df)} estudiantes cargados")
            self.resultados["total_estudiantes"] = len(df)
            return df
        
        except Exception as e:
            raise Exception(f"Error cargando archivo: {e}")

    def detectar_columnas(self, df: pd.DataFrame) -> tuple:
        """Detecta columnas de estudiante y curso"""
        posibles_estudiante = [
            'CODIGO_ESTUDIANTE', 'CODIGO', 'UserPrincipalName', 'UPN', 
            'Email', 'Mail', 'codigo_estudiante', 'codigo', 'upn'
        ]
        
        posibles_curso = [
            'CURSO', 'GRADO', 'Curso', 'Grado', 'curso', 'grado'
        ]
        
        col_estudiante = None
        col_curso = None
        
        for col in posibles_estudiante:
            if col in df.columns:
                col_estudiante = col
                break
        
        for col in posibles_curso:
            if col in df.columns:
                col_curso = col
                break
        
        if not col_estudiante or not col_curso:
            raise ValueError(
                f"Columnas no encontradas.\n"
                f"Se encontró: Estudiante={col_estudiante}, Curso={col_curso}\n"
                f"Disponibles: {list(df.columns)}"
            )
        
        print(f"✅ Columnas: {col_estudiante}, {col_curso}")
        return col_estudiante, col_curso

    def validar_datos(self, df: pd.DataFrame, col_est: str, col_curso: str) -> bool:
        """Valida datos"""
        print("\n🔍 Validando datos...")
        errores = []
        
        for idx, row in df.iterrows():
            est = str(row[col_est]).strip() if row[col_est] else ""
            curso = str(row[col_curso]).strip() if row[col_curso] else ""
            
            if not est or est == "nan":
                errores.append(f"Fila {idx+2}: Estudiante vacío")
            
            if not curso or curso == "nan":
                errores.append(f"Fila {idx+2}: Curso vacío")
        
        if errores:
            print(f"❌ {len(errores)} errores encontrados")
            for error in errores[:5]:
                print(f"   {error}")
            self.resultados["errores"].extend(errores)
            return False
        
        print("✅ Validación exitosa")
        return True

    def obtener_grupo_por_codigo(self, codigo_grupo: str) -> dict or None:
        """
        Busca grupo por código
        Ejemplo: Código "101" → busca "Estudiantes Curso - 101"
        """
        nombre_completo = f"Estudiantes Curso - {codigo_grupo}"
        
        for grupo in self.grupos_disponibles:
            if grupo.get("displayName") == nombre_completo:
                return grupo
        
        return None

    def obtener_user_id(self, upn: str) -> str or None:
        """Obtiene ID del usuario"""
        if not self.token or not upn:
            return None
        
        if upn in self.usuarios_cache:
            return self.usuarios_cache[upn]
        
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        try:
            url = f"{config.GRAPH_ENDPOINT}/users/{upn}?$select=id"
            response = requests.get(url, headers=headers, verify=False, timeout=10)
            
            if response.status_code == 200:
                user_id = response.json().get("id")
                self.usuarios_cache[upn] = user_id
                return user_id
        except:
            pass
        
        return None

    def agregar_a_grupo(self, group_id: str, user_id: str) -> tuple:
        """
        Agrega estudiante al grupo
        Equivalente a: Add-DistributionGroupMember
        """
        if not self.token:
            return False, "Token no disponible"
        
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        url = f"{config.GRAPH_ENDPOINT}/groups/{group_id}/members/$ref"
        body = {"@odata.id": f"{config.GRAPH_ENDPOINT}/directoryObjects/{user_id}"}
        
        try:
            response = requests.post(url, json=body, headers=headers, verify=False, timeout=10)
            
            if response.status_code == 204:
                return True, "Agregado"
            elif response.status_code == 400:
                # Ya está en el grupo
                return True, "Ya en grupo"
            else:
                return False, f"Error {response.status_code}"
        except Exception as e:
            return False, f"Error: {str(e)[:50]}"

    def procesar(self, df: pd.DataFrame, col_est: str, col_curso: str) -> dict:
        """
        Procesa vinculación
        EQUIVALENTE AL POWERSHELL #6:
        
        $DistributionGroups | foreach {
            $Curso = $_.Name
            $Filtered = $Estudiantes | Where-Object {($_.CURSO -eq $Curso)}
            $Filtered | foreach {
                Add-DistributionGroupMember -Identity $_ -Member $Estudiante
            }
        }
        """
        print("\n" + "="*70)
        print("🔄 VINCULANDO ESTUDIANTES A GRUPOS")
        print("="*70)
        
        # Agrupar estudiantes por curso
        estudiantes_por_curso = {}
        for idx, row in df.iterrows():
            est = str(row[col_est]).strip()
            curso = str(row[col_curso]).strip()
            
            if not est or est == "nan" or not curso or curso == "nan":
                continue
            
            if curso not in estudiantes_por_curso:
                estudiantes_por_curso[curso] = []
            
            estudiantes_por_curso[curso].append(est)
        
        print(f"📊 Estudiantes agrupados por {len(estudiantes_por_curso)} cursos")
        
        # Para cada grupo disponible
        count_grupos = 0
        for grupo in self.grupos_disponibles:
            nombre_grupo = grupo.get("displayName", "")
            group_id = grupo.get("id")
            
            # Extraer código del grupo (ej: "101" de "Estudiantes Curso - 101")
            codigo_grupo = nombre_grupo.replace("Estudiantes Curso - ", "").strip()
            
            print(f"\n[{count_grupos+1}] Procesando grupo: {nombre_grupo}")
            
            # Buscar estudiantes para este curso
            if codigo_grupo not in estudiantes_por_curso:
                print(f"    ℹ️  Sin estudiantes para este curso")
                self.resultados["detalles_grupos"].append({
                    "Grupo": nombre_grupo,
                    "Estudiantes": 0,
                    "Vinculados": 0,
                    "Errores": 0
                })
                count_grupos += 1
                continue
            
            estudiantes_curso = estudiantes_por_curso[codigo_grupo]
            count_estudiantes = 0
            count_errores = 0
            
            print(f"    👥 {len(estudiantes_curso)} estudiantes para vincular")
            
            # Para cada estudiante del curso
            for estudiante_upn in estudiantes_curso:
                # Obtener ID del estudiante
                user_id = self.obtener_user_id(estudiante_upn)
                
                if not user_id:
                    print(f"       ❌ {estudiante_upn}: No encontrado")
                    self.resultados["estudiantes_no_encontrados"] += 1
                    count_errores += 1
                    continue
                
                # Agregar al grupo
                exito, msg = self.agregar_a_grupo(group_id, user_id)
                
                if exito:
                    if "Ya en grupo" in msg:
                        self.resultados["estudiantes_ya_en_grupo"] += 1
                        print(f"       ⚠️  {estudiante_upn}: Ya estaba")
                    else:
                        self.resultados["estudiantes_vinculados"] += 1
                        count_estudiantes += 1
                        print(f"       ✅ {estudiante_upn}: Vinculado")
                else:
                    print(f"       ❌ {estudiante_upn}: {msg}")
                    self.resultados["errores_vinculacion"] += 1
                    count_errores += 1
                
                self.resultados["estudiantes_procesados"].append({
                    "Estudiante": estudiante_upn,
                    "Grupo": nombre_grupo,
                    "Resultado": msg
                })
            
            print(f"    Resumen: {count_estudiantes} vinculados, {count_errores} errores")
            
            self.resultados["detalles_grupos"].append({
                "Grupo": nombre_grupo,
                "Estudiantes": len(estudiantes_curso),
                "Vinculados": count_estudiantes,
                "Errores": count_errores
            })
            
            count_grupos += 1
        
        print("\n" + "="*70)
        return self.resultados

    def mostrar_resumen(self):
        """Muestra resumen final"""
        print("\n" + "="*70)
        print("📊 RESUMEN DE VINCULACIÓN")
        print("="*70)
        print(f"Total estudiantes cargados: {self.resultados['total_estudiantes']}")
        print(f"Total grupos disponibles: {self.resultados['total_grupos']}")
        print()
        print(f"✅ Estudiantes vinculados: {self.resultados['estudiantes_vinculados']}")
        print(f"⚠️  Ya estaban vinculados: {self.resultados['estudiantes_ya_en_grupo']}")
        print(f"❌ Estudiantes no encontrados: {self.resultados['estudiantes_no_encontrados']}")
        print(f"❌ Grupos no encontrados: {self.resultados['grupos_no_encontrados']}")
        print(f"❌ Errores de vinculación: {self.resultados['errores_vinculacion']}")
        print("="*70)

    def guardar_logs(self):
        """Guarda logs"""
        try:
            os.makedirs(config.CARPETA_LOGS, exist_ok=True)
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            log_file = os.path.join(config.CARPETA_LOGS, f'vinculacion_estudiantes_{timestamp}.log')
            
            with open(log_file, 'w', encoding='utf-8') as f:
                f.write("VINCULACIÓN DE ESTUDIANTES A GRUPOS DE SEGURIDAD\n")
                f.write(f"Colegio: {config.COLEGIO_NOMBRE}\n")
                f.write(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("="*70 + "\n\n")
                
                f.write("RESUMEN:\n")
                f.write(f"Total estudiantes: {self.resultados['total_estudiantes']}\n")
                f.write(f"Total grupos: {self.resultados['total_grupos']}\n")
                f.write(f"Vinculados exitosamente: {self.resultados['estudiantes_vinculados']}\n")
                f.write(f"Ya estaban vinculados: {self.resultados['estudiantes_ya_en_grupo']}\n")
                f.write(f"No encontrados: {self.resultados['estudiantes_no_encontrados']}\n")
                f.write(f"Errores: {self.resultados['errores_vinculacion']}\n\n")
                
                f.write("DETALLES POR GRUPO:\n")
                for detalle in self.resultados['detalles_grupos']:
                    f.write(f"\n{detalle['Grupo']}:\n")
                    f.write(f"  Estudiantes: {detalle['Estudiantes']}\n")
                    f.write(f"  Vinculados: {detalle['Vinculados']}\n")
                    f.write(f"  Errores: {detalle['Errores']}\n")
                
                if self.resultados['errores']:
                    f.write("\n\nERRORES:\n")
                    for error in self.resultados['errores']:
                        f.write(f"  • {error}\n")
            
            print(f"\n📝 Log guardado: {log_file}")
        
        except Exception as e:
            print(f"❌ Error guardando log: {e}")

    def ejecutar(self, ruta_archivo: str) -> dict:
        """Proceso principal"""
        print("\n" + "="*70)
        print("🏫 VINCULADOR DE ESTUDIANTES A GRUPOS")
        print("="*70)
        
        try:
            # 1. Token
            if not self.obtener_token():
                raise Exception("No se pudo obtener token")
            
            # 2. Obtener grupos de Azure AD
            if not self.obtener_todos_los_grupos():
                raise Exception("No se pudieron obtener grupos")
            
            # 3. Cargar estudiantes
            df = self.cargar_estudiantes(ruta_archivo)
            
            # 4. Detectar columnas
            col_est, col_curso = self.detectar_columnas(df)
            
            # 5. Validar
            if not self.validar_datos(df, col_est, col_curso):
                raise Exception("Validación fallida")
            
            # 6. Procesar
            self.procesar(df, col_est, col_curso)
            
            # 7. Resumen
            self.mostrar_resumen()
            
            # 8. Logs
            self.guardar_logs()
            
            return self.resultados
        
        except Exception as e:
            print(f"❌ Error: {e}")
            self.resultados["errores"].append(str(e))
            return self.resultados


def main():
    """Para pruebas"""
    print("🏫 VINCULADOR DE ESTUDIANTES A GRUPOS")
    
    try:
        vinculador = VinculadorEstudiantesGrupos()
        ruta = input("📁 Ruta del archivo Excel: ").strip()
        
        if not os.path.exists(ruta):
            print(f"❌ Archivo no encontrado")
            return
        
        vinculador.ejecutar(ruta)
    
    except Exception as e:
        print(f"❌ Error: {e}")


if __name__ == "__main__":
    main()
