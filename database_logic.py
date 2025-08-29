import sqlite3
from datetime import date, timedelta
from typing import Tuple, List, Dict, Optional

DB_FILE = 'transporte_operaciones.db'


def setup_database():
    """Create database tables if they do not exist."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    # -------------------------
    # Users (staff)
    # -------------------------
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            role TEXT,
            badge TEXT UNIQUE NOT NULL,
            source TEXT NOT NULL
        )
    ''')

    # -------------------------
    # Operations/rotations history (rangos informativos)
    # -------------------------
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

    # -------------------------
    # NUEVO: schedules (estado día a día)
    # -------------------------
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS schedules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            badge TEXT NOT NULL,
            date TEXT NOT NULL,              -- 'YYYY-MM-DD'
            status TEXT NOT NULL,            -- 'ON', 'ON NS', 'OFF'
            shift_type TEXT,                 -- 'Day Shift' | 'Night Shift' | NULL
            source TEXT NOT NULL,
            UNIQUE (badge, date, source)
    )''')

    # -------------------------
    # NUEVO: audit_log
    # -------------------------
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,          -- quien realizó la acción
            source TEXT NOT NULL,            -- RGM | Newmont
            action_type TEXT NOT NULL,       -- INICIO_SESION | MODIFICACION_TURNO | EXPORTACION_DATOS | IMPORTACION_DATOS | ...
            detail TEXT,
            ts TEXT NOT NULL DEFAULT (datetime('now'))
    )''')

    conn.commit()
    conn.close()


# -----------------------------------------------------------------------------
# Users CRUD
# -----------------------------------------------------------------------------
def add_user(name: str, role: str, badge: str, source: str) -> Tuple[bool, str]:
    """Add a new user to the database."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    try:
        cursor.execute(
            "INSERT INTO users (name, role, badge, source) VALUES (?, ?, ?, ?)",
            (name, role, badge, source)
        )
        conn.commit()
        return True, f"User {name} added successfully."
    except sqlite3.IntegrityError:
        return False, f"Error: The badge '{badge}' already exists in the database."
    except sqlite3.Error as e:
        return False, f"Database error: {e}"
    finally:
        conn.close()


def add_users_bulk(users: list, source: str) -> int:
    """Add a list of users to the database, avoiding duplicates by 'badge'."""
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
        cursor.executemany(
            "INSERT INTO users (name, role, badge, source) VALUES (?, ?, ?, ?)",
            user_data
        )
        conn.commit()
        added_count = cursor.rowcount
    except sqlite3.Error as e:
        print(f"Database error when adding users in bulk: {e}")
        added_count = 0
    finally:
        conn.close()

    return added_count


def get_all_users(source: str) -> list:
    """Get all users from the database for a specific source."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id, name, role, badge FROM users WHERE source = ? ORDER BY name",
        (source,)
    )
    users = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return users


def update_user(user_id: int, name: str, role: str, badge: str, source: str) -> Tuple[bool, str]:
    """Update an existing user's data."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    try:
        # Check if the new badge is already in use by ANOTHER user from the same source
        cursor.execute(
            "SELECT id FROM users WHERE badge = ? AND source = ? AND id != ?",
            (badge, source, user_id)
        )
        if cursor.fetchone():
            return False, f"Error: The badge '{badge}' is already assigned to another user."

        cursor.execute(
            "UPDATE users SET name = ?, role = ?, badge = ? WHERE id = ?",
            (name, role, badge, user_id)
        )
        conn.commit()
        if cursor.rowcount > 0:
            return True, f"User {name} updated successfully."
        else:
            return False, "Error: User not found for update."
    except sqlite3.Error as e:
        return False, f"Database error: {e}"
    finally:
        conn.close()


def delete_user(user_id: int) -> Tuple[bool, str]:
    """Delete a user from the database."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM users WHERE id = ?", (user_id,))
        conn.commit()
        if cursor.rowcount > 0:
            return True, "User deleted successfully."
        else:
            return False, "Error: User not found for deletion."
    except sqlite3.Error as e:
        return False, f"Database error: {e}"
    finally:
        conn.close()


# -----------------------------------------------------------------------------
# Operations (rangos) + Auditoría + Schedules (día a día)
# -----------------------------------------------------------------------------
def add_operation(username: str, role: str, badge: str, start_date: date, end_date: date):
    """Add an operation (rotation range) record to the database."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO operations (username, role, badge, start_date, end_date) VALUES (?, ?, ?, ?, ?)",
        (username, role, badge, start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d'))
    )
    conn.commit()
    conn.close()


def get_all_operations() -> list:
    """Get all operation records from the database."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM operations ORDER BY start_date DESC")
    operations = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return operations


def log_event(username: str, source: str, action_type: str, detail: Optional[str] = None) -> None:
    """Append an event to the audit_log table."""
    try:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO audit_log (username, source, action_type, detail) VALUES (?, ?, ?, ?)",
            (username, source, action_type, detail)
        )
        conn.commit()
    except Exception as e:
        # No romper el flujo si el log falla
        print(f"[audit_log] Could not log event: {e}")
    finally:
        try:
            conn.close()
        except Exception:
            pass


def upsert_schedule_single(badge: str, d: date, status: str,
                           shift_type: Optional[str], source: str) -> None:
    """
    Inserta/actualiza un único día en 'schedules'.
    Mapea 'ON' + Night a 'ON NS'.
    """
    text_status = status
    if status == "ON" and (shift_type or "").lower().startswith("night"):
        text_status = "ON NS"

    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO schedules (badge, date, status, shift_type, source)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(badge, date, source) DO UPDATE
            SET status=excluded.status, shift_type=excluded.shift_type
    """, (badge, d.strftime('%Y-%m-%d'), text_status, shift_type, source))
    conn.commit()
    conn.close()


def upsert_schedule_range(badge: str, start_date: date, end_date: date,
                          status: str, shift_type: Optional[str], source: str) -> int:
    """Inserta/actualiza estado día a día en schedules."""
    count = 0
    d = start_date
    while d <= end_date:
        upsert_schedule_single(badge, d, status, shift_type, source)
        d += timedelta(days=1)
        count += 1
    return count


def clear_schedule_range(badge: str, start_date: date, end_date: date, source: str) -> int:
    """Elimina registros de 'schedules' para el rango indicado (útil al 'limpiar' días)."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        DELETE FROM schedules
          WHERE badge = ? AND source = ?
            AND date BETWEEN ? AND ?
    """, (badge, source, start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d')))
    deleted = cursor.rowcount
    conn.commit()
    conn.close()
    return deleted


def get_schedules_for_source(source: str) -> List[Dict]:
    """Devuelve todos los schedules para un 'source'."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM schedules WHERE source=? ORDER BY date", (source,))
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


def get_audit_log(limit: int = 500, source: Optional[str] = None) -> List[Dict]:
    """Devuelve eventos de auditoría (más recientes primero)."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    if source:
        cursor.execute(
            "SELECT * FROM audit_log WHERE source = ? ORDER BY ts DESC LIMIT ?",
            (source, limit)
        )
    else:
        cursor.execute(
            "SELECT * FROM audit_log ORDER BY ts DESC LIMIT ?",
            (limit,)
        )
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows
