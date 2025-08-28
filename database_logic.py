import sqlite3
from datetime import date
from typing import Tuple

DB_FILE = 'transporte_operaciones.db'

def setup_database():
    """Crea las tablas de la base de datos si no existen."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    # Tabla para usuarios (personal)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            role TEXT,
            badge TEXT UNIQUE NOT NULL,
            source TEXT NOT NULL
        )
    ''')
    
    # Tabla para el historial de operaciones/rotaciones
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS operations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            role TEXT,
            badge TEXT,
            start_date TEXT NOT NULL,
            end_date TEXT NOT NULL
        )
    ''')
    conn.commit()
    conn.close()

def add_user(name: str, role: str, badge: str, source: str) -> Tuple[bool, str]:
    """Añade un nuevo usuario a la base de datos."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO users (name, role, badge, source) VALUES (?, ?, ?, ?)", (name, role, badge, source))
        conn.commit()
        return True, f"Usuario {name} añadido exitosamente."
    except sqlite3.IntegrityError:
        return False, f"Error: El badge '{badge}' ya existe en la base de datos."
    except sqlite3.Error as e:
        return False, f"Error de base de datos: {e}"
    finally:
        conn.close()

def add_users_bulk(users: list, source: str) -> int:
    """Añade una lista de usuarios a la base de datos, evitando duplicados por 'badge'."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    cursor.execute("SELECT badge FROM users WHERE source = ?", (source,))
    existing_badges = {row[0] for row in cursor.fetchall()}
    
    new_users = [user for user in users if str(user.get('badge')) not in existing_badges]
    
    if not new_users:
        conn.close()
        return 0

    user_data = [(user['name'], user['role'], str(user['badge']), source) for user in new_users]
    
    try:
        cursor.executemany("INSERT INTO users (name, role, badge, source) VALUES (?, ?, ?, ?)", user_data)
        conn.commit()
        added_count = cursor.rowcount
    except sqlite3.Error as e:
        print(f"Error en la base de datos al añadir usuarios en bloque: {e}")
        added_count = 0
    finally:
        conn.close()
        
    return added_count

def get_all_users(source: str) -> list:
    """Obtiene todos los usuarios de la base de datos para una fuente específica."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT id, name, role, badge FROM users WHERE source = ? ORDER BY name", (source,))
    users = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return users

def update_user(user_id: int, name: str, role: str, badge: str, source: str) -> Tuple[bool, str]:
    """Actualiza los datos de un usuario existente."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    try:
        # Verificar si el nuevo badge ya está en uso por OTRO usuario de la misma fuente
        cursor.execute("SELECT id FROM users WHERE badge = ? AND source = ? AND id != ?", (badge, source, user_id))
        if cursor.fetchone():
            return False, f"Error: El badge '{badge}' ya está asignado a otro usuario."
        
        cursor.execute("UPDATE users SET name = ?, role = ?, badge = ? WHERE id = ?", (name, role, badge, user_id))
        conn.commit()
        if cursor.rowcount > 0:
            return True, f"Usuario {name} actualizado exitosamente."
        else:
            return False, "Error: No se encontró el usuario para actualizar."
    except sqlite3.Error as e:
        return False, f"Error de base de datos: {e}"
    finally:
        conn.close()

def delete_user(user_id: int) -> Tuple[bool, str]:
    """Elimina un usuario de la base de datos."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM users WHERE id = ?", (user_id,))
        conn.commit()
        if cursor.rowcount > 0:
            return True, "Usuario eliminado exitosamente."
        else:
            return False, "Error: No se encontró el usuario para eliminar."
    except sqlite3.Error as e:
        return False, f"Error de base de datos: {e}"
    finally:
        conn.close()

def add_operation(username: str, role: str, badge: str, start_date: date, end_date: date):
    """Añade un registro de operación (rotación) a la base de datos."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("INSERT INTO operations (username, role, badge, start_date, end_date) VALUES (?, ?, ?, ?, ?)",
                   (username, role, badge, start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d')))
    conn.commit()
    conn.close()

def get_all_operations() -> list:
    """Obtiene todos los registros de operaciones de la base de datos."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM operations ORDER BY start_date DESC")
    operations = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return operations