"""
Módulo de autenticación para la aplicación
"""
from functools import wraps
from flask import session, redirect, url_for, flash
from werkzeug.security import check_password_hash, generate_password_hash
import os
from dotenv import load_dotenv

load_dotenv()

def login_required(f):
    """Decorador para proteger rutas que requieren autenticación"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'logged_in' not in session or not session['logged_in']:
            flash('Por favor inicie sesión para acceder a esta página', 'warning')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def verificar_credenciales(username, password):
    """Verifica las credenciales del usuario"""
    admin_username = os.getenv('ADMIN_USERNAME', 'admin')
    admin_password = os.getenv('ADMIN_PASSWORD', 'Admin2024!')
    
    # Verificar username y password
    if username == admin_username and password == admin_password:
        return True
    return False

def generar_password_hash(password):
    """Genera un hash seguro de una contraseña"""
    return generate_password_hash(password, method='pbkdf2:sha256')
