# Guía de Exportación e Instalación Manual - Aprovisionamiento M365

Este documento explica cómo llevar la aplicación a otro servidor o máquina sin usar GitHub.

## 1. Preparación para Exportar
Antes de comprimir la carpeta, asegúrate de limpiar los archivos temporales para que el paquete sea ligero y seguro:
- **Borrar carpeta `venv`**: Es la carpeta del entorno virtual (ocupa mucho espacio y no sirve en otra máquina).
- **Borrar carpetas `__pycache__`**: Son archivos temporales de Python.
- **Borrar carpeta `.git`**: (Opcional) Si no quieres llevar el historial de versiones.
- **Archivo `.env`**: Asegúrate de NO incluir tu archivo `.env` personal si contiene claves sensibles. El usuario usará la plantilla.

## 2. Archivos Críticos que deben ir en el ZIP
- `app.py` (Servidor principal)
- `requirements.txt` (Lista de librerías)
- `.env.template` (La guía para el nuevo usuario)
- Carpetas `scripts/`, `static/` y `templates/` (Todo el código y diseño)

## 3. Instalación en la nueva máquina

### Paso A: Entorno
1.  **Instalar Python**: Descargar la versión 3.12 o superior de [python.org](https://www.python.org/).
2.  **Descomprimir**: Poner los archivos en una carpeta fija (ej: `C:\GestionM365`).

### Paso B: Configuración
1.  Renombrar el archivo `.env.template` a `.env`.
2.  Abrir el `.env` con un bloc de notas y cambiar los valores:
    - **TENANT_ID**, **CLIENT_ID**, **CLIENT_SECRET** del nuevo cliente.
    - **DEFAULT_CITY** (ej: Cali, Medellín, etc.).
    - **COLEGIO_DOMINIO**, etc.

### Paso C: Instalación de Librerías
Abrir una terminal (PowerShell o CMD) en la carpeta del proyecto y ejecutar:
```bash
# Crear entorno virtual
python -m venv venv

# Activar entorno
.\venv\Scripts\activate

# Instalar librerías
pip install -r requirements.txt
```

### Paso D: Ejecución
Para iniciar la aplicación, siempre con el entorno activado:
```bash
python app.py
```
La aplicación estará disponible en `http://127.0.0.1:5000`.
