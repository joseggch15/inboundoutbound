import sqlite3
from datetime import date

DB_NAME = "transporte_operaciones.db"

def setup_database():
    """
    Prepara la base de datos. Crea las tablas 'operations' y 'users' 
    y se asegura de que la tabla 'users' tenga la columna 'source'.
    """
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS operations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            role TEXT,
            badge TEXT,
            start_date DATE NOT NULL,
            end_date DATE NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            role TEXT,
            badge TEXT,
            source TEXT,
            UNIQUE(name, source),
            UNIQUE(badge, source)
        )
    """)
    conn.commit()
    conn.close()

# --- Funciones para CRUD de Usuarios ---

def add_user(name: str, role: str, badge: str, source: str) -> tuple:
    """Añade un nuevo usuario a la tabla 'users' con su fuente (source)."""
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO users (name, role, badge, source) VALUES (?, ?, ?, ?)", (name, role, badge, source))
        conn.commit()
        conn.close()
        return True, "Usuario añadido exitosamente."
    except sqlite3.IntegrityError as e:
        return False, f"Error: El nombre o el badge ya existen para {source}. ({e})"

def add_users_bulk(users: list, source: str) -> int:
    """
    Añade una lista de usuarios a la base de datos en una sola transacción,
    asociándolos a una fuente (source).
    """
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    user_data = [(user['NAME'], user['ROLE'], str(user['BADGE']), source) for user in users]
    
    cursor.executemany("INSERT OR IGNORE INTO users (name, role, badge, source) VALUES (?, ?, ?, ?)", user_data)
    
    added_rows = cursor.rowcount
    conn.commit()
    conn.close()
    return added_rows

def get_all_users(source: str) -> list:
    """Obtiene todos los usuarios de una fuente específica, ordenados por nombre."""
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT id, name, role, badge FROM users WHERE source = ? ORDER BY name", (source,))
    records = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return records

def update_user(user_id: int, name: str, role: str, badge: str, source: str) -> tuple:
    """Actualiza un usuario existente."""
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE users SET name = ?, role = ?, badge = ? WHERE id = ? AND source = ?",
            (name, role, badge, user_id, source)
        )
        conn.commit()
        conn.close()
        return True, "Usuario actualizado exitosamente."
    except sqlite3.IntegrityError as e:
        return False, f"Error: El nombre o el badge ya existen para {source}. ({e})"

def delete_user(user_id: int) -> tuple:
    """Elimina un usuario de la tabla 'users'."""
    try:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM users WHERE id = ?", (user_id,))
        conn.commit()
        conn.close()
        return True, "Usuario eliminado exitosamente."
    except Exception as e:
        return False, f"Error al eliminar el usuario: {e}"

# --- Funciones para Operaciones (Plan Staff) ---

def add_operation(username: str, role: str, badge: str, start_date: date, end_date: date):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO operations (username, role, badge, start_date, end_date) VALUES (?, ?, ?, ?, ?)",
        (username, role, badge, start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d'))
    )
    conn.commit()
    conn.close()

def get_all_operations() -> list:
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT id, username, role, badge, start_date, end_date FROM operations ORDER BY id DESC")
    records = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return records
