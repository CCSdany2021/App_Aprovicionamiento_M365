
import requests
import json
import pandas as pd
import os
import sys
from datetime import datetime

# Ajustar path para importar configuración
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from scripts.configuracion import config

class GestorGruposSeguridadV2:
    def __init__(self):
        try:
            config.validar_configuracion()
            self.token = None
            self.output_folder = config.CARPETA_RESULTADOS
            self.logs_folder = config.CARPETA_LOGS
            
        except Exception as e:
            print(f"Error config: {e}")
            raise

    def ejecutar(self, accion, **kwargs):
        """Método principal para streaming de eventos (Generator)"""
        
        yield {"status": "info", "message": f"Iniciando gestor de grupos... Acción: {accion}", "progress": 5}
        
        try:
            if not self._obtener_token():
                yield {"status": "error", "message": "No se pudo autenticar con Azure AD"}
                return
                
            if accion == 'reporte_vacios':
                yield from self._generar_reporte_vacios()
            elif accion == 'inventario_completo':
                yield from self._generar_inventario_completo()
            elif accion == 'crear_masivo':
                filepath = kwargs.get('filepath')
                yield from self._crear_grupos_masivo(filepath)
            elif accion == 'exportar_miembros':
                group_id_o_nombre = kwargs.get('grupo_target')
                yield from self._exportar_miembros(group_id_o_nombre)
            elif accion == '_agregar_miembros':
                mode = kwargs.pop('mode', None)
                yield from self._agregar_miembros(mode, **kwargs)
            else:
                 yield {"status": "error", "message": f"Acción desconocida: {accion}"}

        except Exception as e:
             yield {"status": "error", "message": f"Error crítico: {str(e)}"}


    def _obtener_token(self):
        url = f"https://login.microsoftonline.com/{config.TENANT_ID}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": config.CLIENT_ID,
            "client_secret": config.CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default"
        }
        try:
            r = requests.post(url, data=data)
            r.raise_for_status()
            self.token = r.json()["access_token"]
            self.headers = {"Authorization": f"Bearer {self.token}", "Content-Type": "application/json"}
            return True
        except Exception as e:
            print(f"Error auth: {e}")
            return False

    def _api_get_all(self, url, yield_progress=False):
        """Helper para paginación"""
        items = []
        while url:
            try:
                r = requests.get(url, headers=self.headers)
                if r.status_code == 401:
                    self._obtener_token() # Intento renovación simple
                    r = requests.get(url, headers=self.headers)
                
                if r.status_code != 200:
                    break
                    
                data = r.json()
                batch = data.get('value', [])
                items.extend(batch)
                
                if yield_progress:
                    # Solo update log ligero, no yield real aquí para no romper flujo
                    pass

                url = data.get('@odata.nextLink')
            except:
                break
        return items

    def _generar_reporte_vacios(self):
        yield {"status": "info", "message": "Consultando todos los grupos de seguridad...", "progress": 10}
        
        url = f"{config.GRAPH_ENDPOINT}/groups?$filter=securityEnabled eq true&$select=id,displayName,mail"
        grupos = self._api_get_all(url)
        
        total = len(grupos)
        yield {"status": "info", "message": f"Analizando {total} grupos...", "progress": 20}
        
        vacios = []
        for i, g in enumerate(grupos):
            # Checkeo rápido top 1
            url_check = f"{config.GRAPH_ENDPOINT}/groups/{g['id']}/members?$top=1&$select=id"
            try:
                r = requests.get(url_check, headers=self.headers)
                if r.status_code == 200:
                     if not r.json().get('value'):
                         vacios.append(g)
            except:
                pass
            
            # Progress cada 10 items o 10%
            if i % 10 == 0:
                 prog = 20 + int((i / total) * 70)
                 yield {"status": "log", "message": f"Revisando: {g.get('displayName')}...", "progress": prog}

        if vacios:
            df = pd.DataFrame(vacios)
            archivo = f"Reporte_Grupos_Vacios_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            path = os.path.join(self.output_folder, archivo)
            df.to_excel(path, index=False)
            
            yield {"status": "complete", "message": f"Se encontraron {len(vacios)} grupos vacíos.", "results": {"archivo": archivo}}
        else:
            yield {"status": "complete", "message": "No se encontraron grupos vacíos."}

    def _generar_inventario_completo(self):
        yield {"status": "info", "message": "Descargando inventario completo...", "progress": 10}
        
        url = f"{config.GRAPH_ENDPOINT}/groups?$filter=securityEnabled eq true&$select=id,displayName,mail,description,groupTypes,mailEnabled"
        grupos = self._api_get_all(url)
        
        yield {"status": "info", "message": f"Procesando {len(grupos)} registros...", "progress": 80}
        
        data = []
        for g in grupos:
            tipo = "Seguridad"
            if g.get('mailEnabled'): tipo += " + Correo"
            data.append({
                "Nombre": g.get('displayName'),
                "Email": g.get('mail'),
                "Tipo": tipo,
                "ID": g.get('id'),
                "Descripción": g.get('description'),
                "Habilitado Correo": "SI" if g.get('mailEnabled') else "NO"
            })
            
        df = pd.DataFrame(data)
        archivo = f"Inventario_Total_Grupos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path = os.path.join(self.output_folder, archivo)
        df.to_excel(path, index=False)
        
        yield {"status": "complete", "message": "Inventario generado exitosamente", "results": {"archivo": archivo}}

    def _crear_grupos_masivo(self, filepath):
        yield {"status": "info", "message": "Leyendo archivo de plantilla...", "progress": 10}
        
        try:
            df = pd.read_excel(filepath).fillna("")
        except Exception as e:
            yield {"status": "error", "message": f"Error leyendo Excel: {e}"}
            return

        total = len(df)
        exitos = 0
        errores = 0
        
        yield {"status": "info", "message": f"Iniciando creación de {total} grupos...", "progress": 15}
        
        for idx, row in df.iterrows():
            nombre = str(row.get("NombreGrupo", "")).strip()
            nickname = str(row.get("MailNickname", "")).strip()
            # Fallback generacion nickname
            if not nickname and nombre:
                 nickname = nombre.replace(" ", "").lower()[:20]
                 
            if not nombre: continue
            
            desc = str(row.get("Descripcion", "")).strip()
            mail_enabled = str(row.get("RequiereCorreo", "NO")).upper() == "SI"
            
            body = {
                "displayName": nombre,
                "mailNickname": nickname,
                "description": desc,
                "securityEnabled": True,
                "mailEnabled": mail_enabled,
                "groupTypes": []
            }
            
            try:
                r = requests.post(f"{config.GRAPH_ENDPOINT}/groups", headers=self.headers, json=body)
                if r.status_code in [200, 201]:
                    exitos += 1
                    yield {"status": "log", "message": f"✅ Creado: {nombre}"}
                else:
                    errores += 1
                    yield {"status": "log", "message": f"❌ Error {nombre}: {r.text}"}
                    
            except Exception as e:
                errores += 1
                yield {"status": "log", "message": f"❌ Excepción {nombre}: {e}"}
            
            # Progress
            prog = 15 + int(((idx+1)/total)*80)
            yield {"status": "info", "message": f"Procesando {idx+1}/{total}...", "progress": prog}

        yield {"status": "complete", "message": f"Proceso finalizado. Éxitos: {exitos}, Errores: {errores}"}

    def _exportar_miembros(self, nombre_buscar):
        yield {"status": "info", "message": f"Buscando grupo: {nombre_buscar}...", "progress": 10}
        
        # 1. Buscar ID
        url = f"{config.GRAPH_ENDPOINT}/groups?$filter=startswith(displayName, '{nombre_buscar}')&$select=id,displayName"
        res = self._api_get_all(url)
        
        if not res:
             yield {"status": "error", "message": "No se encontró ningún grupo con ese nombre."}
             return
             
        # Tomamos el primero por defecto (simplificación para web)
        grupo_target = res[0]
        yield {"status": "info", "message": f"Encontrado: {grupo_target['displayName']}. Descargando miembros...", "progress": 30}
        
        url_m = f"{config.GRAPH_ENDPOINT}/groups/{grupo_target['id']}/members?$select=id,displayName,userPrincipalName,mail,mailNickname,department,jobTitle"
        miembros = self._api_get_all(url_m)
        
        if not miembros:
            yield {"status": "complete", "message": "El grupo está vacío (0 miembros)."}
            return
            
        yield {"status": "info", "message": f"Exportando {len(miembros)} miembros...", "progress": 80}
        
        data = []
        for m in miembros:
             if m.get('@odata.type') == '#microsoft.graph.user':
                # Calcular Alias (mailNickname o parte del UPN)
                alias = m.get('mailNickname')
                if not alias and m.get('userPrincipalName'):
                     alias = m.get('userPrincipalName').split('@')[0]
                     
                data.append({
                    "Nombre": m.get('displayName'),
                    "UPN": m.get('userPrincipalName'),
                    "Alias": alias,
                    "Departamento": m.get('department'),
                    "Cargo": m.get('jobTitle')
                })
        
        df = pd.DataFrame(data)
        sanitized = grupo_target['displayName'].replace(" ", "_")
        archivo = f"Miembros_{sanitized}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        path = os.path.join(self.output_folder, archivo)
        df.to_excel(path, index=False)
        
        yield {"status": "complete", "message": "Exportación finalizada.", "results": {"archivo": archivo}}


    def _buscar_usuario_por_email_o_upn(self, identificador):
        """Busca ID de usuario por mail o UPN"""
        identificador = identificador.strip().replace("'", "")
        # Intento UPN
        url = f"{config.GRAPH_ENDPOINT}/users/{identificador}?$select=id,displayName"
        try:
            r = requests.get(url, headers=self.headers)
            if r.status_code == 200:
                return r.json().get('id')
        except: pass
        
        # Intento mail
        url = f"{config.GRAPH_ENDPOINT}/users?$filter=mail eq '{identificador}'&$select=id"
        try:
            r = requests.get(url, headers=self.headers)
            if r.status_code == 200:
                val = r.json().get('value')
                if val: return val[0].get('id')
        except: pass
        
        return None

    def _buscar_grupo_exacto(self, nombre):
        """Busca ID de grupo por nombre exacto (o startswith)"""
        url = f"{config.GRAPH_ENDPOINT}/groups?$filter=displayName eq '{nombre}'&$select=id"
        try:
            r = requests.get(url, headers=self.headers)
            if r.status_code == 200 and r.json().get('value'):
                return r.json()['value'][0]['id']
        except: pass
        return None

    def _agregar_miembros(self, mode, **kwargs):
        """Agrega miembros. mode='individual' o 'masivo'"""
        yield {"status": "info", "message": f"Iniciando adición de miembros ({mode})...", "progress": 5}
        
        tareas = [] # Tuplas (email_usuario, nombre_grupo)
        
        if mode == 'individual':
            tareas.append((kwargs.get('usuario'), kwargs.get('grupo')))
        elif mode == 'masivo':
            filepath = kwargs.get('filepath')
            try:
                df = pd.read_excel(filepath)
                # Normalizar columnas
                # Esperamos: Usuario, NombreGrupo
                # Si no, buscar primera y segunda columna
                if 'Usuario' not in df.columns or 'NombreGrupo' not in df.columns:
                     cols = df.columns.tolist()
                     if len(cols) >= 2:
                         df.rename(columns={cols[0]: 'Usuario', cols[1]: 'NombreGrupo'}, inplace=True)
                
                for _, row in df.iterrows():
                    tareas.append((row['Usuario'], row['NombreGrupo']))
                    
            except Exception as e:
                yield {"status": "error", "message": f"Error leyendo archivo: {e}"}
                return

        total = len(tareas)
        exitos = 0
        fallos = 0
        
        yield {"status": "info", "message": f"Procesando {total} solicitudes...", "progress": 10}
        
        cache_grupos = {} # Nombre -> ID
        
        for idx, (usuario, nombre_grupo) in enumerate(tareas):
            usuario = str(usuario).strip()
            nombre_grupo = str(nombre_grupo).strip()
            
            if not usuario or not nombre_grupo: continue
            
            # 1. Resolver Grupo
            group_id = cache_grupos.get(nombre_grupo)
            if not group_id:
                group_id = self._buscar_grupo_exacto(nombre_grupo)
                if group_id: 
                    cache_grupos[nombre_grupo] = group_id
                else:
                    yield {"status": "log", "message": f"❌ Grupo no encontrado: {nombre_grupo}"}
                    fallos += 1
                    continue
            
            # 2. Resolver Usuario
            user_id = self._buscar_usuario_por_email_o_upn(usuario)
            if not user_id:
                yield {"status": "log", "message": f"❌ Usuario no encontrado: {usuario}"}
                fallos += 1
                continue
                
            # 3. Agregar
            url = f"{config.GRAPH_ENDPOINT}/groups/{group_id}/members/$ref"
            body = {"@odata.id": f"{config.GRAPH_ENDPOINT}/directoryObjects/{user_id}"}
            
            try:
                r = requests.post(url, headers=self.headers, json=body)
                if r.status_code == 204:
                    yield {"status": "log", "message": f"✅ {usuario} -> {nombre_grupo}"}
                    exitos += 1
                elif "already exist" in r.text or "One or more added object references already exist" in r.text:
                    yield {"status": "log", "message": f"⚠️ Ya existe: {usuario} en {nombre_grupo}"}
                    # Contamos como éxito técnico (el estado final deseado se cumple)
                    exitos += 1
                else:
                    # Intentar parsear el error para mostrar algo útil
                    try:
                        err_json = r.json()
                        err_msg = err_json.get('error', {}).get('message', r.text)
                    except:
                        err_msg = r.text
                    
                    yield {"status": "log", "message": f"❌ Error API ({r.status_code}): {err_msg}"}
                    fallos += 1
            except Exception as e:
                yield {"status": "log", "message": f"❌ Excepción: {e}"}
                fallos += 1
            
            # Progress
            if total > 1:
                prog = 10 + int(((idx+1)/total)*85)
                yield {"status": "info", "message": f"Procesando... ({exitos} OK / {fallos} Error)", "progress": prog}
                
        yield {"status": "complete", "message": f"Proceso finalizado. Agregados/Verificados: {exitos}, Errores: {fallos}"}
