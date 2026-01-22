# scripts/configuracion.py
"""
Configuración centralizada para la gestión de estudiantes M365 COLEGIOS 
"""

import os
from dotenv import load_dotenv
from werkzeug.security import generate_password_hash

# Cargar variables de entorno
from scripts.gestor_configuracion import GestorConfiguracion

# Cargar variables de entorno
load_dotenv()

class ConfiguracionM365:
    """Clase para manejar toda la configuración del proyecto"""
    
    def __init__(self):
        # 1. Cargar desde DB si existe
        self.gestor = GestorConfiguracion()
        db_config = self.gestor.obtener_configuracion()
        
        if db_config:
            self.TENANT_ID = db_config['tenant_id']
            self.CLIENT_ID = db_config['client_id']
            self.CLIENT_SECRET = db_config['client_secret']
            self.COLEGIO_NOMBRE = db_config['colegio_nombre']
            self.COLEGIO_DOMINIO = db_config['colegio_dominio']
            self.PERIODO_ACTUAL = db_config['periodo_actual']
            self.ADMIN_USER = db_config['admin_user']
            self.ADMIN_PASSWORD_HASH = db_config['admin_password_hash']
            self.ADMIN_PASSWORD_HASH = db_config['admin_password_hash']
            self.EMAIL_SENDER = db_config['email_sender']
            self.TEAM_FUENTE_ID = db_config['team_fuente_id']
            self.DEFAULT_CITY = db_config.get('default_city', 'Bogotá')
        else:
            # 2. Fallback al .env
            self.TENANT_ID = os.getenv('TENANT_ID')
            self.CLIENT_ID = os.getenv('CLIENT_ID')
            self.CLIENT_SECRET = os.getenv('CLIENT_SECRET')
            self.COLEGIO_NOMBRE = os.getenv('COLEGIO_NOMBRE')
            self.COLEGIO_DOMINIO = os.getenv('COLEGIO_DOMINIO')
            self.PERIODO_ACTUAL = os.getenv('DEFAULT_DEPARTMENT', 'Estudiantes 2026').split()[-1]
            self.ADMIN_USER = os.getenv('ADMIN_USER', 'admin')
            self.ADMIN_PASSWORD = os.getenv('ADMIN_PASSWORD')
            self.ADMIN_PASSWORD_HASH = generate_password_hash(self.ADMIN_PASSWORD) if self.ADMIN_PASSWORD else None
            self.EMAIL_SENDER = os.getenv('EMAIL_SENDER')
            self.TEAM_FUENTE_ID = os.getenv('TEAM_FUENTE_ID')

        # Configuración Microsoft 365 Invariante
        self.AUTHORITY = os.getenv('AUTHORITY', f"https://login.microsoftonline.com/{self.TENANT_ID}")
        self.GRAPH_ENDPOINT = os.getenv('GRAPH_ENDPOINT', 'https://graph.microsoft.com/v1.0')
        
        # Configuración del colegio (ya cargada arriba)
        self.COLEGIO_CODIGO = os.getenv('COLEGIO_CODIGO', 'CCS')
        
        # Configuración por defecto para usuarios
        self.DEFAULT_PASSWORD_POLICY = os.getenv('DEFAULT_PASSWORD_POLICY', 'DisablePasswordExpiration')
        self.DEFAULT_USAGE_LOCATION = os.getenv('DEFAULT_USAGE_LOCATION', 'CO')
        self.DEFAULT_DEPARTMENT = os.getenv('DEFAULT_DEPARTMENT', 'Estudiantes')
        self.DEFAULT_JOB_TITLE = os.getenv('DEFAULT_JOB_TITLE', 'Estudiante')
        self.DEFAULT_CITY = getattr(self, 'DEFAULT_CITY', os.getenv('DEFAULT_CITY', 'Bogotá'))
        
        # Licencias
        self.LICENSE_STUDENT = os.getenv('LICENSE_STUDENT')
        self.LICENSE_FACULTY = os.getenv('LICENSE_FACULTY')
        
        # Configuración de Email (Microsoft Graph API)
        self.EMAIL_SENDER = os.getenv('EMAIL_SENDER')
        

        
        # Carpetas
        self.CARPETA_RESULTADOS = os.getenv('CARPETA_RESULTADOS', 'resultados')
        self.CARPETA_LOGS = os.getenv('CARPETA_LOGS', 'resultados/logs')
        
        # Credenciales Admin
        self.ADMIN_USER = os.getenv('ADMIN_USER', 'admin')
        self.ADMIN_PASSWORD = os.getenv('ADMIN_PASSWORD')

        # Logging
        self.LOG_LEVEL = os.getenv('LOG_LEVEL', 'INFO')
        self.LOG_FORMAT = os.getenv('LOG_FORMAT', '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    
    def validar_configuracion(self):
        """Valida que todas las configuraciones necesarias estén presentes"""
        errores = []
        
        if not self.TENANT_ID:
            errores.append("TENANT_ID no configurado")
        if not self.CLIENT_ID:
            errores.append("CLIENT_ID no configurado")
        if not self.CLIENT_SECRET:
            errores.append("CLIENT_SECRET no configurado")
        if not self.COLEGIO_DOMINIO:
            errores.append("COLEGIO_DOMINIO no configurado")
            
        # Validación opcional para SMTP si se desea que sea obligatorio
        # if not self.SMTP_USER or not self.SMTP_PASSWORD:
        #     errores.append("Configuración SMTP incompleta")
            
        if errores:
            raise ValueError(f"Errores de configuración: {', '.join(errores)}")
        
        return True
    
    def mostrar_configuracion(self):
        """Muestra la configuración actual (sin mostrar secretos)"""
        print(f"🏫 Colegio: {self.COLEGIO_NOMBRE}")
        print(f"🌐 Dominio: {self.COLEGIO_DOMINIO}")
        print(f"🏷️  Código: {self.COLEGIO_CODIGO}")
        print(f"🔐 Tenant ID: {self.TENANT_ID[:8]}...")
        print(f"📱 Client ID: {self.CLIENT_ID[:8]}...")
        print(f"✅ Configuración válida")

# Instancia global de configuración
config = ConfiguracionM365()

if __name__ == "__main__":
    try:
        config.validar_configuracion()
        config.mostrar_configuracion()
        print("✅ Configuración cargada correctamente")
    except Exception as e:
        print(f"❌ Error en configuración: {e}")
        print("\n💡 Revisa tu archivo .env")