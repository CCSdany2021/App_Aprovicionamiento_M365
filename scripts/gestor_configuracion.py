import sqlite3
import os
import json
from cryptography.fernet import Fernet
from werkzeug.security import generate_password_hash, check_password_hash

class GestorConfiguracion:
    """Maneja la persistencia y cifrado de la configuración del Tenant"""
    
    def __init__(self, db_path='archivos/configuracion.db'):
        self.db_path = db_path
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        self._inicializar_db()
        
        # Llave para cifrado
        self.key_path = 'archivos/.key'
        self._cargar_o_generar_llave()
        self.cipher = Fernet(self.key)

    def _inicializar_db(self):
        """Crea la tabla de configuración y realiza migraciones si es necesario"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS configuracion (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    tenant_id TEXT,
                    client_id TEXT,
                    client_secret_enc BLOB,
                    colegio_nombre TEXT,
                    colegio_dominio TEXT,
                    periodo_actual TEXT DEFAULT '2026',
                    admin_user TEXT DEFAULT 'admin',
                admin_password_hash TEXT,
                    email_sender TEXT,
                    team_fuente_id TEXT
                )
            ''')
            
            # Migración: Verificar si faltan columnas nuevas
            cursor.execute("PRAGMA table_info(configuracion)")
            columnas = [info[1] for info in cursor.fetchall()]
            
            if 'admin_user' not in columnas:
                cursor.execute("ALTER TABLE configuracion ADD COLUMN admin_user TEXT DEFAULT 'admin'")
            if 'admin_password_hash' not in columnas:
                cursor.execute("ALTER TABLE configuracion ADD COLUMN admin_password_hash TEXT")
            if 'email_sender' not in columnas:
                cursor.execute("ALTER TABLE configuracion ADD COLUMN email_sender TEXT")
            if 'team_fuente_id' not in columnas:
                cursor.execute("ALTER TABLE configuracion ADD COLUMN team_fuente_id TEXT")
                
            conn.commit()

    def _cargar_o_generar_llave(self):
        """Carga la llave de cifrado o genera una nueva"""
        if os.path.exists(self.key_path):
            with open(self.key_path, 'rb') as f:
                self.key = f.read()
        else:
            self.key = Fernet.generate_key()
            with open(self.key_path, 'wb') as f:
                f.write(self.key)

    def guardar_configuracion(self, data):
        """Guarda o actualiza la configuración"""
        secret_enc = self.cipher.encrypt(data['client_secret'].encode())
        pass_hash = generate_password_hash(data['admin_password']) if 'admin_password' in data else None
        
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            # Borramos anterior (solo permitimos una configuración activa por ahora)
            cursor.execute('DELETE FROM configuracion')
            cursor.execute('''
                INSERT INTO configuracion 
                (tenant_id, client_id, client_secret_enc, colegio_nombre, colegio_dominio, periodo_actual, admin_user, admin_password_hash, email_sender, team_fuente_id)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                data['tenant_id'], 
                data['client_id'], 
                secret_enc, 
                data['colegio_nombre'], 
                data['colegio_dominio'],
                data.get('periodo_actual', '2026'),
                data.get('admin_user', 'admin'),
                pass_hash if pass_hash else data.get('admin_password_hash'),
                data.get('email_sender'),
                data.get('team_fuente_id')
            ))
            conn.commit()

    def obtener_configuracion(self):
        """Obtiene la configuración actual decodificada"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                cursor.execute('SELECT * FROM configuracion ORDER BY id DESC LIMIT 1')
                row = cursor.fetchone()
                
                if row:
                    data = dict(row)
                    data['client_secret'] = self.cipher.decrypt(row['client_secret_enc']).decode()
                    return data
                return None
        except Exception as e:
            print(f"Error cargando config DB: {e}")
            return None

if __name__ == "__main__":
    # Test rápido
    gestor = GestorConfiguracion()
    print("Base de datos inicializada")
