import pandas as pd
import requests
import urllib3
from datetime import datetime
import os
import sys

sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


class GestorAprovisionamientoGruposSimplificado:
    """Gestor simplificado: Solo UPN + Curso_2026"""
    
    def __init__(self):
        """Inicializa el gestor"""
        try:
            config.validar_configuracion()
        except:
            pass
        
        self.token = None
        self.grupos_cache = {}  # Cache de grupos
        self.usuarios_cache = {}  # Cache de usuarios
        
        self.resultados = {
            "total": 0,
            "procesados": 0,
            "removidos_exitosos": 0,
            "removidos_fallidos": 0,
            "agregados_exitosos": 0,
            "agregados_fallidos": 0,
            "sin_cambios": 0,
            "sin_grupo_actual": 0,
            "usuario_no_encontrado": 0,
            "errores": [],
            "detalles": [],
            "estudiantes_procesados": []
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
            self.resultados["errores"].append(f"Error de autenticación: {str(e)}")
            return False

    def cargar_archivo(self, ruta_archivo: str) -> pd.DataFrame:
        """
        Carga el archivo Excel/CSV SIMPLIFICADO:
        UserPrincipalName | Curso_2026
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
            
            print(f"✅ Archivo cargado: {len(df)} estudiantes encontrados")
            self.resultados["total"] = len(df)
            return df
            
        except Exception as e:
            raise Exception(f"❌ Error leyendo archivo: {e}")

    def detectar_columnas(self, df: pd.DataFrame) -> tuple:
        """
        Detecta automáticamente las 2 columnas requeridas:
        1. UserPrincipalName (UPN)
        2. Curso_2026 (o similar)
        
        Retorna: (columna_upn, columna_curso_2026)
        """
        posibles_upn = [
            'UserPrincipalName', 'UPN', 'Email', 'Mail', 'Correo',
            'userprincipalname', 'upn', 'email', 'mail', 'correo'
        ]
        
        posibles_curso = [
            'Curso_2026', 'Curso_Nuevo', 'Curso', 'CursoNuevo',
            'curso_2026', 'curso_nuevo', 'curso', 'grado',
            'Grado_2026', 'NuevoCurso'
        ]
        
        col_upn = None
        col_curso = None
        
        # Buscar columna UPN
        for col in posibles_upn:
            if col in df.columns:
                col_upn = col
                break
        
        # Buscar columna Curso
        for col in posibles_curso:
            if col in df.columns:
                col_curso = col
                break
        
        if not col_upn or not col_curso:
            raise ValueError(
                f"❌ No se encontraron las columnas requeridas.\n"
                f"Se encontró: UPN={col_upn}, Curso_2026={col_curso}\n"
                f"Columnas disponibles: {list(df.columns)}\n"
                f"Requeridas: UserPrincipalName y Curso_2026"
            )
        
        print(f"✅ Columnas detectadas:")
        print(f"   UPN: {col_upn}")
        print(f"   Curso_2026: {col_curso}")
        
        return col_upn, col_curso

    def validar_datos(self, df: pd.DataFrame, col_upn: str, col_curso: str) -> bool:
        """
        Valida la integridad de los datos
        
        Reglas:
          • UPN no puede estar vacío
          • Curso_2026 no puede estar vacío
        """
        print("\n🔍 Validando datos...")
        errores = []
        
        for idx, row in df.iterrows():
            upn = str(row[col_upn]).strip() if row[col_upn] else ""
            curso = str(row[col_curso]).strip() if row[col_curso] else ""
            
            # Validar UPN
            if not upn or upn == "nan":
                errores.append(f"Fila {idx+2}: UPN vacío")
                continue
            
            # Validar Curso
            if not curso or curso == "nan":
                errores.append(f"Fila {idx+2} ({upn}): Curso_2026 vacío")
                continue
        
        if errores:
            print(f"❌ Se encontraron {len(errores)} errores:")
            for error in errores[:10]:
                print(f"   {error}")
            self.resultados["errores"].extend(errores)
            return False
        else:
            print("✅ Validación completada sin errores")
            return True

    def obtener_grupo_por_nombre(self, nombre_grupo: str) -> dict or None:
        """Busca un grupo de seguridad por nombre"""
        if not self.token:
            return None
        
        nombre_grupo = str(nombre_grupo).strip()
        if not nombre_grupo or nombre_grupo == "nan":
            return None
        
        # Buscar en cache
        if nombre_grupo in self.grupos_cache:
            return self.grupos_cache[nombre_grupo]
        
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        try:
            url = (
                f"{config.GRAPH_ENDPOINT}/groups?"
                f"$filter=displayName eq '{nombre_grupo}'"
                f"&$select=id,displayName,mail"
            )
            response = requests.get(url, headers=headers, verify=False, timeout=10)
            
            if response.status_code == 200:
                data = response.json()
                if data.get('value'):
                    grupo = data['value'][0]
                    resultado = {
                        "GroupId": grupo.get("id"),
                        "DisplayName": grupo.get("displayName"),
                        "Mail": grupo.get("mail")
                    }
                    self.grupos_cache[nombre_grupo] = resultado
                    return resultado
        except Exception as e:
            print(f"⚠️  Error buscando grupo '{nombre_grupo}': {e}")
        
        return None

    def obtener_grupos_usuario(self, user_id: str) -> list:
        """
        Obtiene todos los grupos de seguridad de un usuario
        
        Returns: Lista de GroupIds
        """
        if not self.token:
            return []
        
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        try:
            url = f"{config.GRAPH_ENDPOINT}/users/{user_id}/memberOf?$select=id,displayName"
            response = requests.get(url, headers=headers, verify=False, timeout=10)
            
            if response.status_code == 200:
                datos = response.json()
                grupos = []
                for item in datos.get('value', []):
                    # Solo incluir grupos de seguridad (no teams)
                    if item.get('@odata.type') == '#microsoft.graph.group':
                        grupos.append({
                            "id": item.get("id"),
                            "displayName": item.get("displayName")
                        })
                return grupos
        except Exception as e:
            pass
        
        return []

    def obtener_curso_actual_usuario(self, user_id: str, upn: str) -> str or None:
        """
        Obtiene el curso actual del estudiante buscando en sus grupos
        
        Lógica: Busca en los grupos del usuario si alguno contiene
        "Estudiantes Curso - XXX" y extrae el número del curso
        """
        grupos = self.obtener_grupos_usuario(user_id)
        
        for grupo in grupos:
            nombre = grupo.get("displayName", "")
            # Buscar patrón "Estudiantes Curso - XXX"
            if "Estudiantes Curso -" in nombre:
                # Extraer el curso
                curso = nombre.replace("Estudiantes Curso -", "").strip()
                if curso:
                    return curso
        
        return None

    def obtener_user_id(self, upn: str) -> str or None:
        """Obtiene el ID de un usuario por su UPN"""
        if not self.token:
            return None
        
        # Buscar en cache
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
                # Guardar en cache
                self.usuarios_cache[upn] = user_id
                return user_id
        except:
            pass
        
        return None

    def remover_de_grupo(self, group_id: str, user_id: str) -> tuple:
        """Remueve usuario de grupo. Returns: (éxito, mensaje)"""
        if not self.token:
            return False, "Token no disponible"
        
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        url = f"{config.GRAPH_ENDPOINT}/groups/{group_id}/members/{user_id}/$ref"
        
        try:
            response = requests.delete(url, headers=headers, verify=False, timeout=10)
            
            if response.status_code == 204:
                return True, "Removido exitosamente"
            elif response.status_code == 404:
                return False, "Grupo o usuario no encontrado"
            else:
                return False, f"Error {response.status_code}"
        except Exception as e:
            return False, f"Error: {str(e)[:50]}"

    def agregar_a_grupo(self, group_id: str, user_id: str) -> tuple:
        """Agrega usuario a grupo. Returns: (éxito, mensaje)"""
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
                return True, "Agregado exitosamente"
            elif response.status_code == 400:
                # Ya está en el grupo
                return True, "Ya estaba en el grupo"
            elif response.status_code == 404:
                return False, "Grupo o usuario no encontrado"
            else:
                return False, f"Error {response.status_code}"
        except Exception as e:
            return False, f"Error: {str(e)[:50]}"

    def procesar_estudiantes(self, df: pd.DataFrame, col_upn: str, col_curso: str) -> dict:
        """
        Procesa todos los estudiantes
        
        Lógica SIMPLIFICADA:
          1. Obtener curso actual del estudiante desde Azure AD
          2. Comparar con curso_2026
          3. Si diferentes:
             - Remover del grupo actual
             - Agregar al nuevo grupo
          4. Si iguales:
             - Sin cambio (ignorar)
          5. Si sin grupo actual:
             - Solo agregar (nuevo ingreso)
        """
        print("\n" + "="*70)
        print("🔄 PROCESANDO ESTUDIANTES")
        print("="*70)
        
        for idx, row in df.iterrows():
            upn = str(row[col_upn]).strip()
            curso_nuevo = str(row[col_curso]).strip()
            
            # Validar datos básicos
            if not upn or upn == "nan":
                continue
            if not curso_nuevo or curso_nuevo == "nan":
                continue
            
            print(f"\n[{idx+1}] Procesando: {upn}")
            
            # Obtener ID del usuario
            user_id = self.obtener_user_id(upn)
            if not user_id:
                print(f"    ❌ Usuario no encontrado en Azure AD")
                self.resultados["usuario_no_encontrado"] += 1
                self.resultados["errores"].append(f"{upn}: Usuario no encontrado")
                continue
            
            # Obtener curso actual del usuario
            curso_actual = self.obtener_curso_actual_usuario(user_id, upn)
            
            print(f"    📌 Curso actual: {curso_actual if curso_actual else '(Sin grupo de curso)'}")
            print(f"    📌 Curso nuevo: {curso_nuevo}")
            
            # CASO 1: Sin grupo actual (nuevo ingreso)
            if not curso_actual:
                print(f"    ➕ NUEVO INGRESO → Agregando a Curso {curso_nuevo}")
                
                nombre_grupo = f"Estudiantes Curso - {curso_nuevo}"
                grupo = self.obtener_grupo_por_nombre(nombre_grupo)
                
                if grupo:
                    exito, msg = self.agregar_a_grupo(grupo["GroupId"], user_id)
                    if exito:
                        print(f"    ✅ {msg}")
                        self.resultados["agregados_exitosos"] += 1
                    else:
                        print(f"    ❌ {msg}")
                        self.resultados["agregados_fallidos"] += 1
                else:
                    print(f"    ❌ Grupo no encontrado: {nombre_grupo}")
                    self.resultados["agregados_fallidos"] += 1
            
            # CASO 2: Curso actual = Curso nuevo (sin cambio)
            elif curso_actual == curso_nuevo:
                print(f"    ✅ Sin cambio (mantiene Curso {curso_actual})")
                self.resultados["sin_cambios"] += 1
            
            # CASO 3: Curso actual ≠ Curso nuevo (cambio de curso)
            else:
                print(f"    🔄 CAMBIO DE CURSO → De {curso_actual} a {curso_nuevo}")
                
                # Remover del grupo actual
                nombre_grupo_actual = f"Estudiantes Curso - {curso_actual}"
                grupo_actual = self.obtener_grupo_por_nombre(nombre_grupo_actual)
                
                if grupo_actual:
                    exito_rem, msg_rem = self.remover_de_grupo(grupo_actual["GroupId"], user_id)
                    if exito_rem:
                        print(f"    ✅ Removido: {msg_rem}")
                        self.resultados["removidos_exitosos"] += 1
                    else:
                        print(f"    ⚠️  Error remover: {msg_rem}")
                        self.resultados["removidos_fallidos"] += 1
                else:
                    print(f"    ⚠️  Grupo actual no encontrado: {nombre_grupo_actual}")
                
                # Agregar al nuevo grupo
                nombre_grupo_nuevo = f"Estudiantes Curso - {curso_nuevo}"
                grupo_nuevo = self.obtener_grupo_por_nombre(nombre_grupo_nuevo)
                
                if grupo_nuevo:
                    exito_agr, msg_agr = self.agregar_a_grupo(grupo_nuevo["GroupId"], user_id)
                    if exito_agr:
                        print(f"    ✅ Agregado: {msg_agr}")
                        self.resultados["agregados_exitosos"] += 1
                    else:
                        print(f"    ⚠️  Error agregar: {msg_agr}")
                        self.resultados["agregados_fallidos"] += 1
                else:
                    print(f"    ❌ Grupo nuevo no encontrado: {nombre_grupo_nuevo}")
                    self.resultados["agregados_fallidos"] += 1
            
            self.resultados["procesados"] += 1
            self.resultados["estudiantes_procesados"].append({
                "UPN": upn,
                "Curso_Actual": curso_actual,
                "Curso_Nuevo": curso_nuevo,
                "UserID": user_id
            })
        
        print("\n" + "="*70)
        return self.resultados

    def mostrar_resumen(self):
        """Muestra resumen de la operación"""
        print("\n" + "="*70)
        print("📊 RESUMEN DE APROVISIONAMIENTO")
        print("="*70)
        print(f"Total cargado: {self.resultados['total']}")
        print(f"Total procesado: {self.resultados['procesados']}")
        print()
        print(f"Cambios de curso:")
        print(f"  • Removidos exitosos: {self.resultados['removidos_exitosos']}")
        print(f"  • Removidos fallidos: {self.resultados['removidos_fallidos']}")
        print()
        print(f"Adiciones:")
        print(f"  • Agregados exitosos: {self.resultados['agregados_exitosos']}")
        print(f"  • Agregados fallidos: {self.resultados['agregados_fallidos']}")
        print()
        print(f"Sin cambios: {self.resultados['sin_cambios']}")
        print(f"Usuarios no encontrados: {self.resultados['usuario_no_encontrado']}")
        print()
        if self.resultados["errores"]:
            print(f"❌ Errores: {len(self.resultados['errores'])}")
        print("="*70)

    def guardar_logs(self):
        """Guarda logs detallados"""
        try:
            os.makedirs(config.CARPETA_LOGS, exist_ok=True)
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            log_file = os.path.join(config.CARPETA_LOGS, f'aprovisionamiento_grupos_{timestamp}.log')
            
            with open(log_file, 'w', encoding='utf-8') as f:
                f.write("APROVISIONAMIENTO SIMPLIFICADO - ESTUDIANTES A GRUPOS DE SEGURIDAD\n")
                f.write(f"Colegio: {config.COLEGIO_NOMBRE}\n")
                f.write(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("="*70 + "\n\n")
                
                f.write("RESUMEN:\n")
                f.write(f"Total cargado: {self.resultados['total']}\n")
                f.write(f"Total procesado: {self.resultados['procesados']}\n")
                f.write(f"Removidos exitosos: {self.resultados['removidos_exitosos']}\n")
                f.write(f"Agregados exitosos: {self.resultados['agregados_exitosos']}\n")
                f.write(f"Sin cambios: {self.resultados['sin_cambios']}\n")
                f.write(f"Usuarios no encontrados: {self.resultados['usuario_no_encontrado']}\n")
                f.write(f"Errores: {len(self.resultados['errores'])}\n\n")
                
                if self.resultados['errores']:
                    f.write("ERRORES:\n")
                    for error in self.resultados['errores']:
                        f.write(f"  • {error}\n")
                    f.write("\n")
                
                f.write("ESTUDIANTES PROCESADOS:\n")
                for est in self.resultados['estudiantes_procesados'][:100]:
                    f.write(f"  {est['UPN']}: {est['Curso_Actual']} → {est['Curso_Nuevo']}\n")
            
            print(f"\n📝 Log guardado en: {log_file}")
            
        except Exception as e:
            print(f"❌ Error guardando log: {e}")

    def procesar(self, ruta_archivo: str) -> dict:
        """Proceso principal"""
        print("\n" + "="*70)
        print("🏫 APROVISIONAMIENTO SIMPLIFICADO - " + config.COLEGIO_NOMBRE)
        print("="*70)
        
        try:
            # 1. Cargar archivo
            df = self.cargar_archivo(ruta_archivo)
            
            # 2. Detectar columnas
            col_upn, col_curso = self.detectar_columnas(df)
            
            # 3. Validar datos
            if not self.validar_datos(df, col_upn, col_curso):
                raise Exception("Validación de datos fallida")
            
            # 4. Obtener token
            if not self.obtener_token():
                raise Exception("No se pudo obtener token de acceso")
            
            # 5. Procesar estudiantes
            self.procesar_estudiantes(df, col_upn, col_curso)
            
            # 6. Mostrar resumen
            self.mostrar_resumen()
            
            # 7. Guardar logs
            self.guardar_logs()
            
            return self.resultados
            
        except Exception as e:
            print(f"❌ Error: {e}")
            self.resultados["errores"].append(str(e))
            return self.resultados


def main():
    """Función principal para pruebas"""
    print("🏫 APROVISIONAMIENTO SIMPLIFICADO - ESTUDIANTES A GRUPOS")
    print("="*70)
    
    try:
        gestor = GestorAprovisionamientoGruposSimplificado()
        
        ruta_archivo = input("\n📁 Ruta del archivo (Excel/CSV): ").strip()
        
        if not os.path.exists(ruta_archivo):
            print(f"❌ Archivo no encontrado: {ruta_archivo}")
            return
        
        gestor.procesar(ruta_archivo)
        
    except KeyboardInterrupt:
        print("\n❌ Interrumpido por usuario")
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    main()
