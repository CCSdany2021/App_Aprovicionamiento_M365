import os
import sys
import secrets
import json
import pandas as pd
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, flash, session, send_from_directory, Response, stream_with_context, jsonify, send_file
from functools import wraps
from werkzeug.utils import secure_filename


# Añadir carpeta scripts al path
sys.path.append(os.path.join(os.path.dirname(__file__), 'scripts'))

from scripts.crear_estudiantes import CreadorEstudiantes
from scripts.actualizacion_estudiantes import ActualizadorEstudiantes
from scripts.eliminar_Estudiantes import EliminadorEstudiantes
from scripts.vaciar_equipos import VaciadorEquipos
from scripts.eliminar_equipos_teams import EliminadorTeams
from scripts.estadisticas import AnalizadorEstadisticas
from scripts.configuracion import config
from scripts.gestor_aprovisionamiento_grupos_simplificado import GestorAprovisionamientoGruposSimplificado
from scripts.vinculador_estudiantes_grupos import VinculadorEstudiantesGrupos
from scripts.creador_equipos_teams_multiples_owners import CreadorEquiposTeamsMultipleOwners
from scripts.sincronizador_automatico_teams import SincronizadorAutomaticoTeams
from scripts.sincronizador_politicas_teams import SincronizadorPoliticasTeams
from scripts.creador_reglas_reenvio import CreadorReglasReenvio

    
app = Flask(__name__)
app.secret_key = os.getenv('FLASK_SECRET_KEY', secrets.token_hex(16))
app.config['UPLOAD_FOLDER'] = 'archivos_subidos'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024 # 16MB max

# Asegurar carpetas
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(config.CARPETA_RESULTADOS, exist_ok=True)
os.makedirs(config.CARPETA_LOGS, exist_ok=True)

# Decorador de login requerido
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'logged_in' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

@app.route('/login', methods=['GET', 'POST'])
def login():
    """Manejo de inicio de sesión"""
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        
        if username == config.ADMIN_USER and password == config.ADMIN_PASSWORD:
            session['logged_in'] = True
            session['user'] = username
            flash('Bienvenido al sistema', 'success')
            return redirect(url_for('index'))
        else:
            flash('Credenciales incorrectas', 'error')
            
    return render_template('login.html')

@app.route('/logout')
def logout():
    """Cerrar sesión"""
    session.clear()
    flash('Sesión cerrada correctamente', 'success')
    return redirect(url_for('login'))

@app.route('/')
@login_required
def index():
    return render_template('index.html')

@app.route('/dashboard')
@login_required
def dashboard():
    """Muestra el dashboard de estadísticas"""
    analizador = AnalizadorEstadisticas()
    stats = analizador.obtener_estadisticas_generales()
    return render_template('dashboard.html', stats=stats)

@app.route('/api/dashboard/charts')
@login_required
def dashboard_charts():
    """API para obtener datos de gráficos"""
    analizador = AnalizadorEstadisticas()
    datos = {
        'lineas': analizador.obtener_datos_grafico_lineas(),
        'barras': analizador.obtener_datos_grafico_barras(),
        'dona': analizador.obtener_datos_grafico_dona()
    }
    return jsonify(datos)

@app.route('/upload/<accion>', methods=['GET', 'POST'])
@login_required
def upload(accion):
    if accion not in ['crear', 'actualizar', 'eliminar', 'desvincular', 'aprovisionar_grupos', 'vincular_grupos', 'eliminar_teams', 'crear_teams_con_owners', 'aprovisionar_estudiantes_teams', 'asignar_politicas', 'crear_reglas_reenvio']:
        flash('Acción no válida', 'error')
        return redirect(url_for('index'))
        
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No se seleccionó ningún archivo', 'error')
            return redirect(request.url)
            
        file = request.files['file']
        
        if file.filename == '':
            flash('No se seleccionó ningún archivo', 'error')
            return redirect(request.url)
            
        filename_lower = file.filename.lower()
        if file and (filename_lower.endswith('.xlsx') or filename_lower.endswith('.csv')):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            # Ejecutar proceso
            if accion in ['crear_teams_con_owners', 'aprovisionar_estudiantes_teams']:
                titulos = {
                    'crear_teams_con_owners': 'Creación de Equipos con Owners',
                    'aprovisionar_estudiantes_teams': 'Vinculación de Estudiantes a Teams'
                }
                return render_template('progress.html', 
                                     titulo=titulos.get(accion), 
                                     endpoint=url_for('stream_process', accion=accion, filename=filename),
                                     proxima_ruta=url_for('index'))
            
            tipo_usuario = request.form.get('tipo_usuario', 'estudiante')
            resultados = procesar_accion(accion, filepath, tipo_usuario=tipo_usuario)
            
            return render_template('results.html', resultados=resultados, accion=accion)
        else:
            flash('Formato no permitido. Use .xlsx o .csv', 'error')
            
    titulos = {
        'crear': 'Crear Estudiantes',
        'actualizar': 'Actualizar Estudiantes',
        'eliminar': 'Eliminar Estudiantes',
        'configuracion': 'Configuración del Sistema',
        'desvincular': 'Desvincular Estudiantes',
        'eliminar_teams': 'Eliminar Equipos Teams',
        'aprovisionar_grupos': 'Aprovisionar Grupos de Seguridad',
        'vincular_grupos': 'Vincular Estudiantes a Grupos',
        'crear_teams_con_owners': 'Crear Teams con Múltiples Owners',
        'aprovisionar_estudiantes_teams': 'Vincular Estudiantes a Teams',
        'asignar_politicas': 'Asignar Políticas Teams (Manual)',
        'crear_reglas_reenvio': 'Crear Reglas de Reenvío CCS'
    }
            
    return render_template('upload.html', accion=accion, titulo=titulos.get(accion, 'Acción desconocida'))

def procesar_accion(accion, filepath, **kwargs):
    """Procesa la acción seleccionada"""
    resultados = {}
    
    if accion == 'crear':
        creador = CreadorEstudiantes()
        resultados = creador.procesar_estudiantes(filepath, confirmacion=False)
        
    elif accion == 'actualizar':
        actualizador = ActualizadorEstudiantes()
        resultados = actualizador.procesar_actualizaciones(filepath, confirmacion=False)
        
    elif accion == 'eliminar':
        eliminador = EliminadorEstudiantes()
        # Para eliminar, primero cargamos la lista
        codigos = eliminador.cargar_lista_estudiantes(filepath)
        resultados = eliminador.eliminar_masivo_con_confirmacion(codigos, confirmacion=False)
        
    elif accion == 'desvincular':
        vaciador = VaciadorEquipos()
        resultados = vaciador.procesar(filepath, confirmacion=False)
        
    elif accion == 'aprovisionar_grupos':        
        gestor = GestorAprovisionamientoGruposSimplificado() 
        resultados = gestor.procesar(filepath)  
        
    elif accion == 'vincular_grupos':
        vinculador = VinculadorEstudiantesGrupos()
        resultados = vinculador.ejecutar(filepath)
    
    # ✅ CAMBIO IMPORTANTE: Agregar elif para crear_teams_con_owners
    elif accion == 'crear_teams_con_owners':
        creador = CreadorEquiposTeamsMultipleOwners()
        resultados = creador.ejecutar(filepath)
    
    elif accion == 'eliminar_teams':
        eliminador = EliminadorTeams()
        resultados = eliminador.procesar(filepath, confirmacion=False)

    elif accion == 'aprovisionar_estudiantes_teams':
        vinculador = VinculadorEstudiantesTeams()
        resultados = vinculador.ejecutar(filepath)
    
    elif accion == 'sincronizacion_automatica':
        sincronizador = SincronizadorAutomaticoTeams(departamento_filtro="Estudiantes 2026")
        # Para llamadas síncronas, agotamos el generador y tomamos el último resultado
        res_list = list(sincronizador.ejecutar())
        if res_list:
            resultados = res_list[-1].get("results", {})
        
    elif accion == 'asignar_politicas':
        asignador = SincronizadorPoliticasTeams()
        # Para llamadas síncronas (opcional, aunque se prefiere SSE)
        res_list = list(asignador.ejecutar(filepath))
        if res_list:
            resultados = res_list[-1].get("results", {})
            
    elif accion == 'crear_reglas_reenvio':
        creador = CreadorReglasReenvio()
        res_list = list(creador.ejecutar(filepath))
        if res_list:
            resultados = res_list[-1].get("results", {})
            
    return resultados

@app.route('/stream_process', methods=['GET'])
@login_required
def stream_process():
    """Endpoint para streaming de eventos SSE"""
    accion = request.args.get('accion')
    filename = request.args.get('filename') or request.args.get('amp;filename', '')
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename) if filename else ''
    
    print(f"DEBUG: Stream solicitado - Acción: {accion}, Filename: {filename}, Path: {filepath}")
    
    def generate():
        objeto_proceso = None
        
        if accion == 'crear_teams_con_owners':
            objeto_proceso = CreadorEquiposTeamsMultipleOwners()
            gen = objeto_proceso.ejecutar(filepath)
        elif accion == 'aprovisionar_estudiantes_teams':
            objeto_proceso = VinculadorEstudiantesTeams()
            gen = objeto_proceso.ejecutar(filepath)
        elif accion == 'sync_teams':
            objeto_proceso = SincronizadorAutomaticoTeams(departamento_filtro="Estudiantes 2026")
            gen = objeto_proceso.ejecutar()
        elif accion == 'sync_policies':
            objeto_proceso = SincronizadorPoliticasTeams(departamento_filtro="Estudiantes 2026")
            gen = objeto_proceso.ejecutar()
        elif accion == 'asignar_politicas':
            objeto_proceso = SincronizadorPoliticasTeams()
            gen = objeto_proceso.ejecutar(filepath)
        elif accion == 'crear_reglas_reenvio':
            objeto_proceso = CreadorReglasReenvio()
            gen = objeto_proceso.ejecutar(filepath)
        else:
            yield f"data: {json.dumps({'status': 'error', 'message': 'Acción no soportada para streaming'})}\n\n"
            return

        for update in gen:
            yield f"data: {json.dumps(update)}\n\n"

    return Response(stream_with_context(generate()), mimetype='text/event-stream')

@app.route('/sync_teams', methods=['POST'])
@login_required
def sync_teams():
    """Muestra la interfaz de progreso para sincronización"""
    return render_template('progress.html', 
                          titulo="Sincronización Automática de Equipos",
                          endpoint=url_for('stream_process', accion='sync_teams'),
                          proxima_ruta=url_for('index'))

@app.route('/sync_policies', methods=['POST'])
@login_required
def sync_policies():
    """Muestra la interfaz de progreso para sincronización de políticas"""
    return render_template('progress.html', 
                          titulo="Sincronización Automática de Políticas Teams",
                          endpoint=url_for('stream_process', accion='sync_policies'),
                          proxima_ruta=url_for('index'))

@app.route('/ccs_dashboard')
@login_required
def ccs_dashboard():
    """Dashboard exclusivo para reglas de reenvío CCS"""
    return render_template('ccs_dashboard.html')

@app.route('/descargar_plantilla_reenvio')
@login_required
def descargar_plantilla_reenvio():
    """Descarga la plantilla CSV de reenvío"""
    ruta = os.path.join(os.getcwd(), 'archivos', 'plantillas', 'Reglas_Reenvio_Plantilla.xlsx')
    return send_file(ruta, as_attachment=True)

@app.route('/logs')
@login_required
def logs():
    log_files = sorted(os.listdir(config.CARPETA_LOGS), reverse=True)
    return render_template('logs.html', logs=log_files)

@app.route('/ver_log/<filename>')
@login_required
def ver_log(filename):
    try:
        filepath = os.path.join(config.CARPETA_LOGS, filename)
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        return render_template('view_log.html', content=content, filename=filename)
    except Exception as e:
        flash(f'Error leyendo log: {e}', 'error')
        return redirect(url_for('logs'))

@app.route('/descargar_log/<filename>')
@login_required
def descargar_log(filename):
    """Descarga un archivo de log"""
    try:
        filepath = os.path.join(config.CARPETA_LOGS, filename)
        if os.path.exists(filepath):
            return send_file(filepath, as_attachment=True, download_name=filename)
        else:
            flash('Archivo de log no encontrado', 'error')
            return redirect(url_for('logs'))
    except Exception as e:
        flash(f'Error descargando log: {e}', 'error')
        return redirect(url_for('logs'))


@app.route('/descargar_inventario')
@login_required
def descargar_inventario():
    """Genera y descarga el inventario de equipos"""
    try:
        vaciador = VaciadorEquipos()
        ruta_archivo = vaciador.generar_inventario(config.CARPETA_RESULTADOS)
        
        if ruta_archivo and os.path.exists(ruta_archivo):
            return send_file(ruta_archivo, as_attachment=True)
        else:
            flash("Error generando el inventario o no se encontraron equipos.", "error")
            return redirect(url_for('index'))
    except Exception as e:
        flash(f"Error crítico: {str(e)}", "error")
        return redirect(url_for('index'))

@app.route('/descargar_plantilla/<tipo>')
@login_required
def descargar_plantilla(tipo):
    """Descarga plantillas para los procesos"""
    try:
        flash("Tipo de plantilla no válido", "error")
        return redirect(url_for('index'))
    except Exception as e:
        flash(f"Error generando plantilla: {e}", "error")
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, port=5000)