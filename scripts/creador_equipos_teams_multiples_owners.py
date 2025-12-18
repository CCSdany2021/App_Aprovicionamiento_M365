import pandas as pd
import requests
import urllib3
from datetime import datetime
import os
import sys
import time

sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


class CreadorEquiposTeamsMultipleOwners:
    """Crea Teams con múltiples owners automáticamente"""
    
    def __init__(self):
        try:
            config.validar_configuracion()
        except:
            pass
        
        self.token = None
        self.team_fuente_id = self.obtener_team_fuente_id_desde_env()
        self.usuarios_cache = {}
        self.teams_existentes = {}
        
        self.resultados = {
            "total": 0,
            "creados_exitosamente": 0,
            "equipos_ya_existentes": 0,
            "equipos_omitidos_duplicado": 0,
            "docentes_no_encontrados": 0,
            "team_fuente_no_disponible": 0,
            "errores_clonacion": 0,
            "total_owners_agregados": 0,
            "errores_agregando_owners": 0,
            "errores": [],
            "detalles": [],
            "equipos_procesados": [],
            "equipos_saltados": []
        }
    
    def obtener_team_fuente_id_desde_env(self) -> str:
        """Obtiene ID del Team Fuente desde .env"""
        team_fuente_id = os.getenv('TEAM_FUENTE_ID')
        
        if team_fuente_id:
            print(f"✅ ID Team Fuente obtenido de variable de entorno")
            return team_fuente_id
        
        try:
            env_paths = [
                '.env',
                os.path.join(os.path.dirname(__file__), '..', '.env'),
                os.path.join(os.path.dirname(__file__), '..', '..', '.env')
            ]
            
            for env_file in env_paths:
                if os.path.exists(env_file):
                    with open(env_file, 'r') as f:
                        for line in f:
                            if line.startswith('TEAM_FUENTE_ID='):
                                team_fuente_id = line.split('=')[1].strip()
                                if team_fuente_id:
                                    print(f"✅ ID Team Fuente obtenido de {env_file}")
                                    return team_fuente_id
        except Exception as e:
            print(f"⚠️ Error leyendo .env: {e}")
        
        print("❌ ADVERTENCIA: ID del Team Fuente NO configurado")
        return None
    
    def obtener_token(self) -> bool:
        """Obtiene token de acceso"""
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

    def cargar_archivo(self, ruta_archivo: str) -> pd.DataFrame:
        """Carga Excel - Detecta automáticamente la hoja"""
        try:
            if ruta_archivo.endswith(".xlsx"):
                try:
                    df = pd.read_excel(ruta_archivo, sheet_name="Grupos de Estudio", dtype=str)
                    print("✅ Hoja 'Grupos de Estudio' encontrada")
                except:
                    print("⚠️ Hoja 'Grupos de Estudio' no encontrada. Detectando hojas disponibles...")
                    xl_file = pd.ExcelFile(ruta_archivo)
                    hojas_disponibles = xl_file.sheet_names
                    
                    print(f"📋 Hojas disponibles: {hojas_disponibles}")
                    
                    if not hojas_disponibles:
                        raise ValueError("El Excel no tiene hojas")
                    
                    df = None
                    for hoja in hojas_disponibles:
                        temp_df = pd.read_excel(ruta_archivo, sheet_name=hoja, dtype=str)
                        if len(temp_df) > 0:
                            df = temp_df
                            print(f"✅ Usando hoja: '{hoja}' ({len(df)} filas)")
                            break
                    
                    if df is None:
                        raise ValueError("Todas las hojas están vacías")
            
            elif ruta_archivo.endswith(".csv"):
                df = pd.read_csv(ruta_archivo, dtype=str, encoding="utf-8")
            else:
                raise ValueError("Formato no soportado")
            
            df.columns = df.columns.str.strip()
            df = df.fillna("")
            
            print(f"✅ {len(df)} equipos cargados")
            self.resultados["total"] = len(df)
            return df
        
        except Exception as e:
            raise Exception(f"Error cargando archivo: {e}")

    def detectar_columnas(self, df: pd.DataFrame) -> dict:
        """Detecta TODAS las columnas necesarias y opcionales"""
        columnas_encontradas = {}
        
        mapeos = {
            'Equipo': ['Equipo', 'equipo', 'EQUIPO', 'displayName', 'Team', 'Nombre Equipo'],
            'Docente': ['Docente', 'docente', 'DOCENTE', 'Owner', 'owner', 'Usuario', 'Teacher', 'Profesor'],
            'Grupo': ['Grupo', 'grupo', 'GRUPO', 'Course', 'Section', 'Sección'],
            'Asignatura': ['Asignatura', 'asignatura', 'ASIGNATURA', 'Subject', 'Materia'],
            'Grado': ['Grado', 'grado', 'GRADO', 'Grade', 'Nivel'],
            'CoordinadorSeccion': ['CoordinadorSeccion', 'Coordinador Sección', 'CoordinadordeSeccion', 'SectionCoordinator', 'Coordinador'],
            'CuentaAcademica': ['CuentaAcademica', 'Cuenta Académica', 'CuentaAcademica', 'AcademicAccount'],
            'Owner3': ['Owner3', 'owner3', 'OWNER3', 'AdditionalOwner1'],
            'Owner4': ['Owner4', 'owner4', 'OWNER4', 'AdditionalOwner2']
        }
        
        for key, posibles in mapeos.items():
            for col in posibles:
                if col in df.columns:
                    columnas_encontradas[key] = col
                    break
        
        if 'Equipo' not in columnas_encontradas or 'Docente' not in columnas_encontradas:
            raise ValueError(
                f"Columnas requeridas no encontradas.\n"
                f"Se requieren: 'Equipo' y 'Docente'"
            )
        
        print(f"✅ Columnas detectadas:")
        print(f"   Obligatorias: Equipo, Docente")
        if any(k in columnas_encontradas for k in ['CoordinadorSeccion', 'CuentaAcademica', 'Owner3', 'Owner4']):
            print(f"   Owners adicionales: {', '.join([k for k in columnas_encontradas if k in ['CoordinadorSeccion', 'CuentaAcademica', 'Owner3', 'Owner4']])}")
        
        return columnas_encontradas

    def obtener_todos_teams_existentes(self) -> dict:
        """ANTI-DUPLICADOS: Obtiene TODOS los Teams existentes"""
        if not self.token:
            return {}
        
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        print("\n🔍 Escaneando Teams existentes (ANTI-DUPLICADOS)...")
        
        try:
            teams = {}
            url = f"{config.GRAPH_ENDPOINT}/groups?$select=id,displayName&$top=999"
            
            while url:
                response = requests.get(url, headers=headers, verify=False, timeout=15)
                
                if response.status_code == 200:
                    data = response.json()
                    for item in data.get('value', []):
                        display_name = item.get('displayName', '').strip()
                        team_id = item.get('id')
                        if display_name:
                            teams[display_name] = team_id
                    
                    url = data.get('@odata.nextLink')
                else:
                    break
            
            self.teams_existentes = teams
            print(f"✅ {len(teams)} Teams existentes en el tenant")
            return teams
        
        except Exception as e:
            print(f"⚠️ Error obteniendo Teams: {e}")
            return {}

    def validar_datos(self, df: pd.DataFrame, col_eq: str, col_doc: str) -> bool:
        """Valida datos básicos"""
        print("\n🔍 Validando datos...")
        errores = []
        
        for idx, row in df.iterrows():
            eq = str(row[col_eq]).strip() if row[col_eq] else ""
            doc = str(row[col_doc]).strip() if row[col_doc] else ""
            
            if not eq or eq == "nan":
                errores.append(f"Fila {idx+2}: Equipo vacío")
            
            if not doc or doc == "nan":
                errores.append(f"Fila {idx+2}: Docente vacío")
        
        if errores:
            print(f"❌ {len(errores)} errores encontrados")
            for error in errores[:5]:
                print(f"   {error}")
            self.resultados["errores"].extend(errores)
            return False
        
        print("✅ Validación exitosa")
        return True

    def obtener_user_id(self, upn: str) -> str or None:
        """Obtiene ID del usuario (docente u owner)"""
        if not self.token or not upn:
            return None
        
        upn = upn.strip().lower()
        
        if upn in self.usuarios_cache:
            return self.usuarios_cache[upn]
        
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        try:
            url = f"{config.GRAPH_ENDPOINT}/users/{upn}?$select=id,displayName"
            response = requests.get(url, headers=headers, verify=False, timeout=10)
            
            if response.status_code == 200:
                user_id = response.json().get("id")
                self.usuarios_cache[upn] = user_id
                return user_id
        except Exception as e:
            pass
        
        return None

    def equipo_existe(self, nombre_equipo: str) -> bool:
        """ANTI-DUPLICADOS: Verifica si el Team YA EXISTE"""
        nombre_equipo = nombre_equipo.strip()
        return nombre_equipo in self.teams_existentes

    def clonar_team(self, display_name: str, description: str, owner_principal_upn: str) -> tuple:
        """Clona Team "Fuente" (sin agregar owners aún)"""
        if not self.token or not self.team_fuente_id:
            return False, None, "Token o Team Fuente no disponible"
        
        if self.equipo_existe(display_name):
            print(f"    ⚠️ {display_name}: YA EXISTE (saltando)")
            self.resultados["equipos_omitidos_duplicado"] += 1
            self.resultados["equipos_saltados"].append({
                "Equipo": display_name,
                "Razon": "Ya existe en el tenant"
            })
            return True, None, "Ya existe (saltado)"
        
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        owner_id = self.obtener_user_id(owner_principal_upn)
        if not owner_id:
            return False, None, f"Docente no encontrado: {owner_principal_upn}"
        
        mail_nickname = display_name.replace(" ", "").replace("-", "")[:25]
        
        body = {
            "displayName": display_name,
            "description": description,
            "mailNickname": mail_nickname,
            "partsToClone": "apps,tabs,settings,channels,members"
        }
        
        url = f"{config.GRAPH_ENDPOINT}/teams/{self.team_fuente_id}/clone"
        
        try:
            response = requests.post(url, json=body, headers=headers, verify=False, timeout=30)
            
            if response.status_code == 202:
                print(f"    ✅ Clonado: {display_name}")
                time.sleep(2)
                team_id = self.obtener_team_id_por_nombre(display_name)
                self.teams_existentes[display_name] = "cloning"
                return True, team_id, "Clonado"
            
            elif response.status_code == 400:
                error_detail = response.json().get('error', {}).get('message', '')
                if "already exists" in error_detail.lower():
                    return True, None, "Rechazado (ya existe)"
                return False, None, f"Error: {error_detail[:50]}"
            
            else:
                return False, None, f"Error {response.status_code}"
        
        except Exception as e:
            return False, None, f"Error: {str(e)[:50]}"

    def obtener_team_id_por_nombre(self, display_name: str) -> str or None:
        """Obtiene el ID del Team por su nombre"""
        if not self.token:
            return None
        
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        try:
            if display_name in self.teams_existentes:
                return self.teams_existentes[display_name]
            
            url = f"{config.GRAPH_ENDPOINT}/groups?$filter=displayName eq '{display_name}'&$select=id"
            response = requests.get(url, headers=headers, verify=False, timeout=10)
            
            if response.status_code == 200:
                data = response.json()
                if data.get('value'):
                    team_id = data['value'][0].get('id')
                    return team_id
        except:
            pass
        
        return None

    def actualizar_rol_a_owner(self, team_id: str, user_id: str, email: str) -> bool:
        """Actualiza el rol de un usuario a OWNER si ya es miembro"""
        if not self.token or not team_id or not user_id:
            return False
        
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        body = {
            "roles": ["owner"]
        }
        
        url = f"{config.GRAPH_ENDPOINT}/teams/{team_id}/members/{user_id}"
        
        try:
            response = requests.patch(url, json=body, headers=headers, verify=False, timeout=10)
            
            if response.status_code in [200, 204]:
                print(f"       ✅ ACTUALIZADO A OWNER: {email}")
                return True
            else:
                print(f"       ❌ Error actualizando rol {email}: Status {response.status_code}")
                return False
        
        except Exception as e:
            print(f"       ❌ Error actualizando rol {email}: {str(e)[:50]}")
            return False

    def agregar_owner_individual(self, team_id: str, email: str) -> bool:
        """Agrega UN SOLO owner al Team o actualiza su rol si ya es miembro"""
        if not self.token or not team_id or not email:
            return False
        
        email = email.strip().lower()
        
        if not email or email == "nan" or email == "":
            return False
        
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        user_id = self.obtener_user_id(email)
        if not user_id:
            print(f"       ⚠️ NO ENCONTRADO: {email}")
            self.resultados["errores_agregando_owners"] += 1
            self.resultados["errores"].append(f"Owner no agregado: {email} - Usuario no encontrado")
            return False
        
        body = {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "roles": ["owner"],
            "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{user_id}')"
        }
        
        url = f"{config.GRAPH_ENDPOINT}/teams/{team_id}/members"
        
        try:
            response = requests.post(url, json=body, headers=headers, verify=False, timeout=10)
            
            if response.status_code in [200, 201]:
                print(f"       ✅ OWNER AGREGADO: {email}")
                return True
            
            elif response.status_code == 409:
                # ✅ NUEVO: Ya es miembro, actualizar a owner
                print(f"       ⚠️ {email}: Ya era miembro, actualizando a OWNER...")
                if self.actualizar_rol_a_owner(team_id, user_id, email):
                    return True
                else:
                    print(f"       ❌ No se pudo actualizar: {email}")
                    return False
            
            else:
                print(f"       ❌ Error agregando {email}: Status {response.status_code}")
                return False
        
        except Exception as e:
            print(f"       ❌ Error agregando {email}: {str(e)[:50]}")
            return False

    def es_valor_valido(self, valor) -> bool:
        """Verifica si un valor es válido y no está vacío"""
        if valor is None:
            return False
        valor_str = str(valor).strip()
        if not valor_str or valor_str == "nan" or valor_str == "":
            return False
        return True

    def procesar(self, df: pd.DataFrame, columnas: dict) -> dict:
        """Procesa clonación Y AGREGACIÓN DE OWNERS - VERSIÓN FINAL CORREGIDA"""
        print("\n" + "="*70)
        print("🔄 CLONANDO TEAMS CON MÚLTIPLES OWNERS")
        print("="*70)
        
        col_eq = columnas.get('Equipo')
        col_doc = columnas.get('Docente')
        col_grupo = columnas.get('Grupo')
        col_asignatura = columnas.get('Asignatura')
        col_grado = columnas.get('Grado')
        col_coord_sec = columnas.get('CoordinadorSeccion')
        col_cuenta_acad = columnas.get('CuentaAcademica')
        col_owner3 = columnas.get('Owner3')
        col_owner4 = columnas.get('Owner4')
        
        for idx, row in df.iterrows():
            eq = str(row[col_eq]).strip() if row[col_eq] else ""
            doc = str(row[col_doc]).strip() if row[col_doc] else ""
            grupo = str(row[col_grupo]).strip() if col_grupo and row[col_grupo] else ""
            asignatura = str(row[col_asignatura]).strip() if col_asignatura and row[col_asignatura] else ""
            grado = str(row[col_grado]).strip() if col_grado and row[col_grado] else ""
            
            if not eq or eq == "nan" or not doc or doc == "nan":
                continue
            
            print(f"\n[{idx+1}] Procesando: {eq}")
            
            description = f"{asignatura} - {grado} {grupo}".strip()
            
            # PASO 1: CLONAR
            exito_clonacion, team_id, msg_clonacion = self.clonar_team(eq, description, doc)
            
            if not exito_clonacion:
                print(f"    ❌ Error clonando: {msg_clonacion}")
                if "Docente no encontrado" in msg_clonacion:
                    self.resultados["docentes_no_encontrados"] += 1
                else:
                    self.resultados["errores_clonacion"] += 1
                self.resultados["errores"].append(f"{eq}: {msg_clonacion}")
            
            elif "Ya existe" in msg_clonacion or "Rechazado" in msg_clonacion:
                self.resultados["equipos_ya_existentes"] += 1
            
            else:
                # PASO 2: AGREGAR MÚLTIPLES OWNERS - CORREGIDO
                if team_id:
                    print(f"    🔐 Agregando owners...")
                    
                    # DOCENTE (siempre se agrega)
                    if self.agregar_owner_individual(team_id, doc):
                        self.resultados["total_owners_agregados"] += 1
                    
                    # COORDINADOR DE SECCIÓN - ✅ VALIDACIÓN CORRECTA
                    if col_coord_sec is not None:
                        if self.es_valor_valido(row[col_coord_sec]):
                            coord = str(row[col_coord_sec]).strip()
                            if self.agregar_owner_individual(team_id, coord):
                                self.resultados["total_owners_agregados"] += 1
                    
                    # CUENTA ACADÉMICA - ✅ VALIDACIÓN CORRECTA
                    if col_cuenta_acad is not None:
                        if self.es_valor_valido(row[col_cuenta_acad]):
                            cuenta = str(row[col_cuenta_acad]).strip()
                            if self.agregar_owner_individual(team_id, cuenta):
                                self.resultados["total_owners_agregados"] += 1
                    
                    # OWNER 3 - ✅ VALIDACIÓN CORRECTA
                    if col_owner3 is not None:
                        if self.es_valor_valido(row[col_owner3]):
                            owner3 = str(row[col_owner3]).strip()
                            if self.agregar_owner_individual(team_id, owner3):
                                self.resultados["total_owners_agregados"] += 1
                    
                    # OWNER 4 - ✅ VALIDACIÓN CORRECTA
                    if col_owner4 is not None:
                        if self.es_valor_valido(row[col_owner4]):
                            owner4 = str(row[col_owner4]).strip()
                            if self.agregar_owner_individual(team_id, owner4):
                                self.resultados["total_owners_agregados"] += 1
                    
                    self.resultados["creados_exitosamente"] += 1
                else:
                    self.resultados["creados_exitosamente"] += 1
            
            self.resultados["equipos_procesados"].append({
                "Equipo": eq,
                "Docente": doc,
                "Resultado": msg_clonacion
            })
            
            time.sleep(1)
        
        print("\n" + "="*70)
        return self.resultados

    def mostrar_resumen(self):
        """Muestra resumen final"""
        print("\n" + "="*70)
        print("📊 RESUMEN: TEAMS CON MÚLTIPLES OWNERS")
        print("="*70)
        print(f"Total equipos procesados: {self.resultados['total']}")
        print()
        print(f"✅ Clonados NUEVOS: {self.resultados['creados_exitosamente']}")
        print(f"⚠️  Ya existían: {self.resultados['equipos_ya_existentes']}")
        print()
        print(f"🔐 OWNERS AGREGADOS:")
        print(f"   Total owners: {self.resultados['total_owners_agregados']}")
        print(f"   Errores: {self.resultados['errores_agregando_owners']}")
        print()
        print(f"❌ Docentes no encontrados: {self.resultados['docentes_no_encontrados']}")
        print(f"❌ Errores de clonación: {self.resultados['errores_clonacion']}")
        print()
        print("🛡️ ANTI-DUPLICADOS: ACTIVADO")
        print("="*70)

    def guardar_logs(self):
        """Guarda logs detallados"""
        try:
            os.makedirs(config.CARPETA_LOGS, exist_ok=True)
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            log_file = os.path.join(config.CARPETA_LOGS, f'teams_con_owners_{timestamp}.log')
            
            with open(log_file, 'w', encoding='utf-8') as f:
                f.write("CLONACIÓN DE TEAMS CON MÚLTIPLES OWNERS\n")
                f.write(f"Colegio: {config.COLEGIO_NOMBRE}\n")
                f.write(f"Team Fuente ID: {self.team_fuente_id}\n")
                f.write(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("="*70 + "\n\n")
                
                f.write("RESUMEN:\n")
                f.write(f"Total procesados: {self.resultados['total']}\n")
                f.write(f"Clonados nuevos: {self.resultados['creados_exitosamente']}\n")
                f.write(f"Ya existían: {self.resultados['equipos_ya_existentes']}\n\n")
                
                f.write(f"OWNERS AGREGADOS:\n")
                f.write(f"Total owners: {self.resultados['total_owners_agregados']}\n")
                f.write(f"Errores: {self.resultados['errores_agregando_owners']}\n\n")
                
                if self.resultados['errores']:
                    f.write("ERRORES:\n")
                    for error in self.resultados['errores']:
                        f.write(f"  • {error}\n")
            
            print(f"\n📝 Log guardado: {log_file}")
        
        except Exception as e:
            print(f"❌ Error guardando log: {e}")

    def ejecutar(self, ruta_archivo: str) -> dict:
        """Proceso principal"""
        print("\n" + "="*70)
        print("🏫 CREADOR DE TEAMS CON MÚLTIPLES OWNERS")
        print("="*70)
        
        try:
            if not self.team_fuente_id:
                raise Exception(
                    "❌ ID del Team Fuente NO configurado\n"
                    "Por favor, agrega a .env:\n"
                    "TEAM_FUENTE_ID=eb1887ba-4fed-4f74-bc55-a0a8fdd7c4f0"
                )
            
            if not self.obtener_token():
                raise Exception("No se pudo obtener token")
            
            self.obtener_todos_teams_existentes()
            
            df = self.cargar_archivo(ruta_archivo)
            
            columnas = self.detectar_columnas(df)
            
            if not self.validar_datos(df, columnas['Equipo'], columnas['Docente']):
                raise Exception("Validación fallida")
            
            self.procesar(df, columnas)
            
            self.mostrar_resumen()
            
            self.guardar_logs()
            
            return self.resultados
        
        except Exception as e:
            print(f"❌ Error: {e}")
            self.resultados["errores"].append(str(e))
            return self.resultados












