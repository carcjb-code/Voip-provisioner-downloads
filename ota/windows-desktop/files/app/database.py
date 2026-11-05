import sqlite3
import os
import logging

logger = logging.getLogger(__name__)

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'provisioner.db')

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS perfiles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL,
            host TEXT NOT NULL,
            usuario TEXT DEFAULT 'root',
            password TEXT DEFAULT '',
            puerto_ssh INTEGER DEFAULT 22,
            path_tftpboot TEXT DEFAULT '/tftpboot'
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS modelos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            marca TEXT NOT NULL,
            modelo TEXT NOT NULL,
            plantilla TEXT NOT NULL
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS backup_config (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            perfil_nombre TEXT UNIQUE NOT NULL,
            hora TEXT NOT NULL,
            dias TEXT NOT NULL,
            destino TEXT DEFAULT '/var/backups/voip',
            frecuencia TEXT DEFAULT 'diario'
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS extensiones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            perfil_nombre TEXT DEFAULT '',
            interno TEXT NOT NULL,
            nombre_mostrar TEXT DEFAULT '',
            nombre_saliente TEXT DEFAULT '',
            mac TEXT DEFAULT '',
            ip TEXT DEFAULT '',
            sn TEXT DEFAULT '',
            modelo TEXT DEFAULT '',
            edificio TEXT DEFAULT '',
            piso TEXT DEFAULT '',
            dependencia TEXT DEFAULT '',
            password TEXT DEFAULT ''
        )
    ''')

    conn.commit()
    conn.close()

    migrar_db()

def migrar_db():
    conn = get_db()
    cursor = conn.cursor()
    
    try:
        cursor.execute("PRAGMA table_info(extensiones);")
        columnas_existentes = [row['name'] for row in cursor.fetchall()]
        
        columnas_requeridas = {
            'nombre_mostrar': 'TEXT DEFAULT ""',
            'nombre_saliente': 'TEXT DEFAULT ""',
            'perfil_nombre': 'TEXT DEFAULT ""',
            'edificio': 'TEXT DEFAULT ""',
            'piso': 'TEXT DEFAULT ""',
            'dependencia': 'TEXT DEFAULT ""',
            'sn': 'TEXT DEFAULT ""',
            'modelo': 'TEXT DEFAULT ""',
            'ip': 'TEXT DEFAULT ""',
            'mac': 'TEXT DEFAULT ""',
            'password': 'TEXT DEFAULT ""'
        }
        
        for col, col_def in columnas_requeridas.items():
            if col not in columnas_existentes:
                cursor.execute(f"ALTER TABLE extensiones ADD COLUMN {col} {col_def};")
                logger.info(f"Columna '{col}' agregada a la tabla extensiones.")
                
        conn.commit()
    except Exception as e:
        logger.error(f"Error migrando base de datos: {e}")
    finally:
        conn.close()
