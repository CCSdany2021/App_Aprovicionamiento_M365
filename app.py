from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify
import os
import sys
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

    
app = Flask(__name__)
app.secret_key = 'supersecretkey_calasanz' # Cambiar en producción
app.config['UPLOAD_FOLDER'] = 'archivos_subidos'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024 # 16MB max

# Asegurar carpetas
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(config.CARPETA_RESULTADOS, exist_ok=True)
os.makedirs(config.CARPETA_LOGS, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/dashboard')
def dashboard():
    """Muestra el dashboard de estadísticas"""
    analizador = AnalizadorEstadisticas()
    stats = analizador.obtener_estadisticas_generales()
    return render_template('dashboard.html', stats=stats)

@app.route('/api/dashboard/charts')
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
def upload(accion):
    # ✅ CAMBIO IMPORTANTE: Agregar 'crear_teams_con_owners' a la lista válida
    if accion not in ['crear', 'actualizar', 'eliminar', 'desvincular', 'aprovisionar_grupos', 'vincular_grupos', 'eliminar_teams', 'crear_teams_con_owners']:
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
            
        if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.csv')):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            # Ejecutar proceso
            resultados = procesar_accion(accion, filepath)
            
            return render_template('results.html', resultados=resultados, accion=accion)
        else:
            flash('Formato no permitido. Use .xlsx o .csv', 'error')
            
    # ✅ CAMBIO IMPORTANTE: Diccionario de títulos
    titulos = {
        'crear': 'Crear Nuevos Estudiantes',
        'actualizar': 'Actualizar Estudiantes',
        'eliminar': 'Eliminar Estudiantes',           
        'configuracion': 'Configuración del Sistema',
        'desvincular': 'Vaciar Equipos (Teams)',
        'eliminar_teams': 'Eliminar Teams del Tenant',
        'aprovisionar_grupos': 'Aprovisionar Estudiantes a Grupos',
        'vincular_grupos': 'Vincular Estudiantes a Grupos',
        'crear_teams_con_owners': 'Crear Equipos de Teams con Owners'  # ✅ NUEVA ACCIÓN
    }
            
    return render_template('upload.html', accion=accion, titulo=titulos.get(accion, 'Acción desconocida'))

def procesar_accion(accion, filepath):
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
           
    return resultados

@app.route('/logs')
def logs():
    log_files = sorted(os.listdir(config.CARPETA_LOGS), reverse=True)
    return render_template('logs.html', logs=log_files)

@app.route('/ver_log/<filename>')
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

if __name__ == '__main__':
    app.run(debug=True, port=5000)