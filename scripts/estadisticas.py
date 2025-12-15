import os
import re
from datetime import datetime, timedelta
from collections import defaultdict
from scripts.configuracion import config

class AnalizadorEstadisticas:
    """Analiza logs para generar estadísticas y métricas"""
    
    def __init__(self):
        self.carpeta_logs = config.CARPETA_LOGS
        
    def obtener_estadisticas_generales(self):
        """Obtiene estadísticas generales de todos los logs"""
        stats = {
            'total_operaciones': 0,
            'estudiantes_creados': 0,
            'estudiantes_actualizados': 0,
            'estudiantes_eliminados': 0,
            'teams_procesados': 0,
            'miembros_eliminados': 0,
            'owners_eliminados': 0,
            'total_errores': 0,
            'tasa_exito': 0,
            'operaciones_por_tipo': defaultdict(int),
            'actividad_reciente': [],
            'operaciones_por_dia': defaultdict(int)
        }
        
        if not os.path.exists(self.carpeta_logs):
            return stats
            
        archivos_log = sorted(
            [f for f in os.listdir(self.carpeta_logs) if f.endswith('.log')],
            reverse=True
        )
        
        for archivo in archivos_log[:50]:  # Últimos 50 logs
            ruta = os.path.join(self.carpeta_logs, archivo)
            tipo_operacion, datos = self._analizar_log(ruta, archivo)
            
            if tipo_operacion:
                stats['total_operaciones'] += 1
                stats['operaciones_por_tipo'][tipo_operacion] += 1
                
                # Extraer fecha del nombre del archivo
                fecha_match = re.search(r'(\d{8})_', archivo)
                if fecha_match:
                    fecha_str = fecha_match.group(1)
                    stats['operaciones_por_dia'][fecha_str] += 1
                
                # Acumular datos específicos
                if tipo_operacion == 'crear_estudiantes':
                    stats['estudiantes_creados'] += datos.get('creados', 0)
                elif tipo_operacion == 'actualizar_estudiantes':
                    stats['estudiantes_actualizados'] += datos.get('actualizados', 0)
                elif tipo_operacion == 'eliminar_estudiantes':
                    stats['estudiantes_eliminados'] += datos.get('eliminados', 0)
                elif tipo_operacion == 'vaciar_equipos':
                    stats['teams_procesados'] += datos.get('equipos', 0)
                    stats['miembros_eliminados'] += datos.get('miembros', 0)
                    stats['owners_eliminados'] += datos.get('owners', 0)
                
                stats['total_errores'] += datos.get('errores', 0)
                
                # Actividad reciente (últimas 10)
                if len(stats['actividad_reciente']) < 10:
                    stats['actividad_reciente'].append({
                        'tipo': tipo_operacion,
                        'fecha': datos.get('fecha', 'N/A'),
                        'exito': datos.get('errores', 0) == 0,
                        'detalles': datos.get('resumen', '')
                    })
        
        # Calcular tasa de éxito
        if stats['total_operaciones'] > 0:
            operaciones_exitosas = stats['total_operaciones'] - (stats['total_errores'] / max(stats['total_operaciones'], 1))
            stats['tasa_exito'] = round((operaciones_exitosas / stats['total_operaciones']) * 100, 1)
        
        return stats
    
    def _analizar_log(self, ruta_archivo, nombre_archivo):
        """Analiza un archivo de log individual"""
        try:
            with open(ruta_archivo, 'r', encoding='utf-8') as f:
                contenido = f.read()
            
            datos = {
                'fecha': self._extraer_fecha(contenido, nombre_archivo),
                'errores': 0,
                'resumen': ''
            }
            
            # Determinar tipo de operación
            if 'crear_estudiantes' in nombre_archivo or 'CREACIÓN DE ESTUDIANTES' in contenido:
                tipo = 'crear_estudiantes'
                datos['creados'] = self._extraer_numero(contenido, r'Estudiantes Creados:\s*(\d+)')
                datos['errores'] = self._extraer_numero(contenido, r'Errores:\s*(\d+)')
                datos['resumen'] = f"{datos['creados']} estudiantes creados"
                
            elif 'actualizacion_estudiantes' in nombre_archivo or 'ACTUALIZACIÓN DE ESTUDIANTES' in contenido:
                tipo = 'actualizar_estudiantes'
                datos['actualizados'] = self._extraer_numero(contenido, r'Estudiantes Actualizados:\s*(\d+)')
                datos['errores'] = self._extraer_numero(contenido, r'Errores:\s*(\d+)')
                datos['resumen'] = f"{datos['actualizados']} estudiantes actualizados"
                
            elif 'eliminacion_estudiantes' in nombre_archivo or 'ELIMINACIÓN DE ESTUDIANTES' in contenido:
                tipo = 'eliminar_estudiantes'
                datos['eliminados'] = self._extraer_numero(contenido, r'Estudiantes Eliminados:\s*(\d+)')
                datos['errores'] = self._extraer_numero(contenido, r'Errores:\s*(\d+)')
                datos['resumen'] = f"{datos['eliminados']} estudiantes eliminados"
                
            elif 'vaciado_equipos' in nombre_archivo or 'VACIADO DE EQUIPOS' in contenido:
                tipo = 'vaciar_equipos'
                datos['equipos'] = self._extraer_numero(contenido, r'Equipos Procesados:\s*(\d+)')
                datos['miembros'] = self._extraer_numero(contenido, r'Miembros Eliminados:\s*(\d+)')
                datos['owners'] = self._extraer_numero(contenido, r'Owners Eliminados:\s*(\d+)')
                datos['errores'] = self._extraer_numero(contenido, r'Errores:\s*(\d+)')
                datos['resumen'] = f"{datos['equipos']} teams procesados"
            else:
                return None, {}
            
            return tipo, datos
            
        except Exception as e:
            print(f"Error analizando {nombre_archivo}: {e}")
            return None, {}
    
    def _extraer_numero(self, texto, patron):
        """Extrae un número usando regex"""
        match = re.search(patron, texto)
        return int(match.group(1)) if match else 0
    
    def _extraer_fecha(self, contenido, nombre_archivo):
        """Extrae la fecha del contenido o nombre del archivo"""
        # Intentar extraer del contenido
        match = re.search(r'Fecha:\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})', contenido)
        if match:
            return match.group(1)
        
        # Intentar extraer del nombre del archivo
        match = re.search(r'(\d{8})_(\d{6})', nombre_archivo)
        if match:
            fecha = match.group(1)
            hora = match.group(2)
            return f"{fecha[:4]}-{fecha[4:6]}-{fecha[6:8]} {hora[:2]}:{hora[2:4]}:{hora[4:6]}"
        
        return 'N/A'
    
    def obtener_datos_grafico_lineas(self, dias=30):
        """Obtiene datos para gráfico de líneas (operaciones por día)"""
        stats = self.obtener_estadisticas_generales()
        
        # Generar últimos N días
        hoy = datetime.now()
        fechas = [(hoy - timedelta(days=i)).strftime('%Y%m%d') for i in range(dias-1, -1, -1)]
        
        datos = {
            'labels': [datetime.strptime(f, '%Y%m%d').strftime('%d/%m') for f in fechas],
            'datasets': [{
                'label': 'Operaciones',
                'data': [stats['operaciones_por_dia'].get(f, 0) for f in fechas]
            }]
        }
        
        return datos
    
    def obtener_datos_grafico_barras(self):
        """Obtiene datos para gráfico de barras (operaciones por tipo)"""
        stats = self.obtener_estadisticas_generales()
        
        tipos_nombres = {
            'crear_estudiantes': 'Crear',
            'actualizar_estudiantes': 'Actualizar',
            'eliminar_estudiantes': 'Eliminar',
            'vaciar_equipos': 'Teams'
        }
        
        datos = {
            'labels': [tipos_nombres.get(k, k) for k in stats['operaciones_por_tipo'].keys()],
            'datasets': [{
                'label': 'Cantidad',
                'data': list(stats['operaciones_por_tipo'].values())
            }]
        }
        
        return datos
    
    def obtener_datos_grafico_dona(self):
        """Obtiene datos para gráfico de dona (éxito vs errores)"""
        stats = self.obtener_estadisticas_generales()
        
        total = stats['total_operaciones']
        errores = stats['total_errores']
        exitos = max(0, total - errores)
        
        datos = {
            'labels': ['Exitosas', 'Con Errores'],
            'datasets': [{
                'data': [exitos, errores]
            }]
        }
        
        return datos
