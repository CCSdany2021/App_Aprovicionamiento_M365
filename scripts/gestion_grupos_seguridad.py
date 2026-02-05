
import requests
import pandas as pd
import json
import os
import sys
from datetime import datetime

# Añadir path para importar configuración
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config

class GestorGruposSeguridad:
    def __init__(self):
        try:
            config.validar_configuracion()
            self.output_folder = config.CARPETA_RESULTADOS
            self.logs_folder = config.CARPETA_LOGS
            self.token = None
            self.headers = {}
        except Exception as e:
            print(f"❌ Error de configuración: {e}")
            sys.exit(1)

    def _obtener_token(self):
        """Obtiene token usando Client Credentials (Service Principal)"""
        print("🔐 Autenticando con Microsoft Graph...")
        url = f"https://login.microsoftonline.com/{config.TENANT_ID}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": config.CLIENT_ID,
            "client_secret": config.CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default"
        }
        try:
            resp = requests.post(url, data=data)
            resp.raise_for_status()
            self.token = resp.json()["access_token"]
            self.headers = {
                "Authorization": f"Bearer {self.token}",
                "Content-Type": "application/json"
            }
            print("✅ Autenticación exitosa")
            return True
        except Exception as e:
            print(f"❌ Error obteniendo token: {e}")
            return False

    def _api_get_all(self, url):
        """Helper para paginación automática de Microsoft Graph"""
        items = []
        while url:
            try:
                if not self.token: self._obtener_token()
                resp = requests.get(url, headers=self.headers)
                
                # Manejo de token expirado
                if resp.status_code == 401:
                    print("🔄 Token expirado, renovando...")
                    if self._obtener_token():
                        resp = requests.get(url, headers=self.headers)
                    else:
                        raise Exception("No se pudo renovar el token")

                if resp.status_code != 200:
                    print(f"⚠️ Error API: {resp.text}")
                    break

                data = resp.json()
                items.extend(data.get('value', []))
                url = data.get('@odata.nextLink')
                
                # Feedback visual simple
                print(f"\r📥 Descargando... {len(items)} registros", end="", flush=True)
                
            except Exception as e:
                print(f"\n❌ Error en petición: {e}")
                break
        print("") # Nueva línea
        return items

    def listar_grupos_seguridad(self):
        """1. Lista todos los grupos de seguridad y genera Excel"""
        if not self.token and not self._obtener_token(): return

        print("\n🔍 Buscando Grupos de Seguridad...")
        
        # Filtramos solo securityEnabled = true
        url = f"{config.GRAPH_ENDPOINT}/groups?$filter=securityEnabled eq true&$select=id,displayName,mail,description,groupTypes,mailEnabled,securityEnabled&$top=999"
        
        grupos = self._api_get_all(url)
        
        if not grupos:
            print("⚠️ No se encontraron grupos.")
            return

        print(f"✅ Se encontraron {len(grupos)} grupos de seguridad.")
        
        # Procesar datos
        data = []
        for g in grupos:
            tipo = "Seguridad"
            if g.get('mailEnabled'): tipo += " + Correo"
            if "Unified" in g.get('groupTypes', []): tipo = "Microsoft 365"
            
            data.append({
                "Nombre Grupo": g.get('displayName'),
                "Correo (Email)": g.get('mail'),
                "Tipo": tipo,
                "ID Grupo": g.get('id'),
                "Descripción": g.get('description'),
                "Habilitado Correo": "SI" if g.get('mailEnabled') else "NO"
            })
            
        df = pd.DataFrame(data)
        
        # Guardar
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        archivo = os.path.join(self.output_folder, f"Inventario_Grupos_Seguridad_{timestamp}.xlsx")
        df.to_excel(archivo, index=False)
        print(f"📄 Archivo guardado: {archivo}")
        return df

    def exportar_miembros_grupo(self):
        """2. Exporta miembros de un grupo específico"""
        nombre_buscar = input("\n📝 Ingrese el NOMBRE (o parte) del grupo a buscar: ").strip()
        if not nombre_buscar: return

        if not self.token and not self._obtener_token(): return

        # Buscar grupo primero
        print(f"🔍 Buscando '{nombre_buscar}'...")
        url = f"{config.GRAPH_ENDPOINT}/groups?$filter=startswith(displayName, '{nombre_buscar}') and securityEnabled eq true&$select=id,displayName,mail"
        
        # Nota: startswith es case-sensitive en algunos queries de Graph, pero intentamos simple primero.
        # Fallback a búsqueda local si no encuentra preciso.
        
        try:
            resp = requests.get(url, headers=self.headers)
            candidates = resp.json().get('value', [])
        except:
            candidates = []

        if not candidates:
            print("❌ No se encontró el grupo.")
            return

        # Selección de grupo
        print("\nGrupos encontrados:")
        for idx, g in enumerate(candidates):
            print(f"  {idx+1}. {g['displayName']} ({g.get('mail') or 'Sin correo'})")
        
        try:
            sel = int(input("\n👉 Seleccione el número (0 para cancelar): "))
            if sel == 0: return
            grupo = candidates[sel-1]
        except:
            print("❌ Selección inválida")
            return

        print(f"\n📥 Descargando miembros de: {grupo['displayName']}...")
        url_miembros = f"{config.GRAPH_ENDPOINT}/groups/{grupo['id']}/members?$select=id,displayName,userPrincipalName,mail,jobTitle,department"
        
        miembros = self._api_get_all(url_miembros)
        
        data = []
        for m in miembros:
            if m.get('@odata.type') == '#microsoft.graph.user':
                data.append({
                    "Nombre": m.get('displayName'),
                    "Usuario (UPN)": m.get('userPrincipalName'),
                    "Correo": m.get('mail'),
                    "Departamento": m.get('department'),
                    "Cargo": m.get('jobTitle'),
                    "Tipo": "Usuario"
                })
            else:
                data.append({
                    "Nombre": m.get('displayName'),
                    "Tipo": "Grupo/Otro",
                    "ID": m.get('id')
                })
        
        if not data:
            print("⚠️ El grupo está vacío.")
        else:
            df = pd.DataFrame(data)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            sanitized_name = grupo['displayName'].replace(" ", "_").replace("/", "-")
            archivo = os.path.join(self.output_folder, f"Miembros_{sanitized_name}_{timestamp}.xlsx")
            df.to_excel(archivo, index=False)
            print(f"📄 Miembros exportados a: {archivo}")

    def descargar_plantilla(self):
        """3. Genera plantilla Excel para creación masiva"""
        data = [{
            "NombreGrupo": "Profesores 2026",
            "MailNickname": "profesores2026",
            "Descripcion": "Grupo de seguridad para profesores",
            "RequiereCorreo": "NO"
        }, {
            "NombreGrupo": "Coordinación - Notificaciones",
            "MailNickname": "coordinacion_notif",
            "Descripcion": "Grupo habilitado con correo",
            "RequiereCorreo": "SI"
        }]
        
        df = pd.DataFrame(data)
        archivo = os.path.join(self.output_folder, "Plantilla_Creacion_Grupos.xlsx")
        df.to_excel(archivo, index=False)
        print(f"\n✅ Plantilla generada en: {archivo}")
        print("ℹ️  Columnas requeridas: NombreGrupo, MailNickname, Descripcion, RequiereCorreo")

    def crear_grupos_masivo(self):
        """4. Crea grupos desde Excel"""
        ruta = input("\n📂 Arrastre el archivo Excel de plantilla aquí: ").strip().replace('"', '')
        if not os.path.exists(ruta):
            print("❌ Archivo no encontrado.")
            return

        if not self.token and not self._obtener_token(): return

        try:
            df = pd.read_excel(ruta)
            df = df.fillna("")
        except Exception as e:
            print(f"❌ Error leyendo Excel: {e}")
            return

        print(f"\n🔄 Procesando {len(df)} grupos...")
        
        exitos = 0
        errores = 0

        for idx, row in df.iterrows():
            nombre = str(row.get("NombreGrupo")).strip()
            nickname = str(row.get("MailNickname")).strip()
            desc = str(row.get("Descripcion", "")).strip()
            mail_enabled = str(row.get("RequiereCorreo", "NO")).upper() == "SI"

            if not nombre or not nickname:
                print(f"⚠️  Fila {idx+2}: Falta Nombre o Nickname. Saltando.")
                continue

            # Verificar existencia previa (simple por Nickname)
            # Nota: Esto es opcional pero recomendado para no duplicar errores
            print(f"procesando: {nombre}...", end="")

            body = {
                "displayName": nombre,
                "mailNickname": nickname,
                "description": desc,
                "securityEnabled": True,
                "mailEnabled": mail_enabled,
                "groupTypes": []  # Security Group puro
            }

            try:
                resp = requests.post(
                    f"{config.GRAPH_ENDPOINT}/groups",
                    headers=self.headers,
                    json=body
                )
                
                if resp.status_code == 201:
                    print(" ✅ Creado")
                    exitos += 1
                elif resp.status_code == 400 and "Request_BadRequest" in resp.text:
                    # Intento común: MailEnabled=True desde Graph a veces falla si no está configurado Exchange
                    print(" ❌ Falló (Posible restricción de Exchange). Intente RequiereCorreo=NO")
                    errores += 1
                    # Log detallado
                    with open(os.path.join(self.logs_folder, "errores_creacion.log"), "a", encoding="utf-8") as f:
                        f.write(f"{nombre}: {resp.text}\n")
                else:
                     print(f" ❌ Error {resp.status_code}")
                     errores += 1
                     with open(os.path.join(self.logs_folder, "errores_creacion.log"), "a", encoding="utf-8") as f:
                        f.write(f"{nombre}: {resp.text}\n")

            except Exception as e:
                print(f" ❌ Excepción: {e}")
                errores += 1

        print(f"\n📊 Resultado Final: {exitos} creados, {errores} fallidos.")
        if errores > 0:
            print("⚠️  Revise 'errores_creacion.log' en la carpeta de logs.")

    def reporte_grupos_vacios(self):
        """5. Reporte de grupos sin miembros (Auditoría)"""
        if not self.token and not self._obtener_token(): return
        print("\n🔍 Analizando grupos (esto puede tardar)...")

        # 1. Traer todos
        url = f"{config.GRAPH_ENDPOINT}/groups?$filter=securityEnabled eq true&$select=id,displayName,mail"
        grupos = self._api_get_all(url)
        
        vacios = []
        print(f"\nVerificando {len(grupos)} grupos...")
        
        for i, g in enumerate(grupos):
            sys.stdout.write(f"\rAnalizando {i+1}/{len(grupos)}")
            sys.stdout.flush()
            
            # Chequeo rápido: pedir solo 1 miembro
            url_check = f"{config.GRAPH_ENDPOINT}/groups/{g['id']}/members?$top=1&$select=id"
            try:
                r = requests.get(url_check, headers=self.headers)
                if r.status_code == 200:
                    data = r.json()
                    if not data.get('value'):
                        vacios.append(g)
            except:
                pass
        
        print(f"\n\n⚠️  Se encontraron {len(vacios)} GRUPOS VACÍOS:")
        for v in vacios[:10]:
            print(f"  - {v['displayName']}")
        if len(vacios) > 10: print(f"  ... y {len(vacios)-10} más.")

        # Guardar
        if vacios:
            df = pd.DataFrame(vacios)
            archivo = os.path.join(self.output_folder, f"Reporte_Grupos_Vacios_{datetime.now().strftime('%Y%m%d')}.xlsx")
            df.to_excel(archivo, index=False)
            print(f"📄 Reporte completo guardado en: {archivo}")

    def menu(self):
        while True:
            print("\n" + "="*50)
            print("🛡️  GESTIÓN DE GRUPOS DE SEGURIDAD (M365) 🛡️")
            print("="*50)
            print("1. 📄 Listar TODOS los grupos (Inventario a Excel)")
            print("2. 👥 Exportar Miembros de un grupo")
            print("3. 📥 Descargar Plantilla de Creación")
            print("4. ✨ Crear Grupos Masivamente (desde Excel)")
            print("5. 🧹 Reporte Grupos Vacíos (Limpieza)")
            print("6. 🚪 Salir")
            
            opcion = input("\n👉 Seleccione una opción: ")
            
            if opcion == '1': self.listar_grupos_seguridad()
            elif opcion == '2': self.exportar_miembros_grupo()
            elif opcion == '3': self.descargar_plantilla()
            elif opcion == '4': self.crear_grupos_masivo()
            elif opcion == '5': self.reporte_grupos_vacios()
            elif opcion == '6': break
            else: print("❌ Opción inválida")

if __name__ == "__main__":
    app = GestorGruposSeguridad()
    app.menu()
