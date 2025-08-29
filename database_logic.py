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
    # schedules (estado día a día)
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
    # audit_log
    # -------------------------
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,          -- quien realizó la acción
            source TEXT NOT NULL,            -- RGM | Newmont | Administrator
            action_type TEXT NOT NULL,       -- USER_LOGIN | SHIFT_MODIFICATION | DATA_EXPORT | DATA_IMPORT | ...
            detail TEXT,
            ts TEXT NOT NULL DEFAULT (datetime('now'))
    )''')

    # Indexes útiles
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_users_source ON users(source)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_schedules_source ON schedules(source)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_schedules_badge ON schedules(badge)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_schedules_date ON schedules(date)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_audit_ts ON audit_log(ts)")

    conn.commit()
    conn.close()


# ---------------------------------------------------------------------
# Audit log
# ---------------------------------------------------------------------
def log_event(username: str, source: str, action_type: str, detail: str = ""):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO audit_log (username, source, action_type, detail) VALUES (?, ?, ?, ?)",
        (username or "Unknown", source or "", action_type or "", detail or "")
    )
    conn.commit()
    conn.close()


def get_audit_log(source: Optional[str] = None) -> List[Dict]:
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    if source:
        cursor.execute(
            "SELECT ts, username, source, action_type, detail FROM audit_log "
            "WHERE source = ? ORDER BY ts DESC",
            (source,))
    else:
        cursor.execute(
            "SELECT ts, username, source, action_type, detail FROM audit_log "
            "ORDER BY ts DESC")
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


# ---------------------------------------------------------------------
# Users CRUD
# ---------------------------------------------------------------------
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
    """
    Add users in bulk, avoiding duplicates by (badge).
    Returns the number of actually inserted users.
    """
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    cursor.execute("SELECT badge FROM users WHERE source = ?", (source,))
    existing_badges = {row[0] for row in cursor.fetchall()}

    new_users = [user for user in users if str(user.get('badge')) not in existing_badges]

    if not new_users:
        conn.close()
        return 0

    user_data = [(user['name'], user['role'], str(user['badge']), source) for user in new_users]

    added_count = 0
    try:
        cursor.executemany(
            "INSERT INTO users (name, role, badge, source) VALUES (?, ?, ?, ?)",
            user_data
        )
        conn.commit()
        added_count = cursor.rowcount if cursor.rowcount is not None else len(new_users)
    except sqlite3.Error as e:
        # En caso de conflicto global de UNIQUE(badge), se omiten esos registros.
        # No abortamos la app.
        print(f"Database error when adding users in bulk: {e}")
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
        if cursor.rowcount and cursor.rowcount > 0:
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
        if cursor.rowcount and cursor.rowcount > 0:
            return True, "User deleted successfully."
        else:
            return False, "Error: User not found for deletion."
    except sqlite3.Error as e:
        return False, f"Database error: {e}"
    finally:
        conn.close()


# ---------------------------------------------------------------------
# Operations & schedules
# ---------------------------------------------------------------------
def add_operation(username: str, role: str, badge: str, start_date: date, end_date: date):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO operations (username, role, badge, start_date, end_date) VALUES (?, ?, ?, ?, ?)",
        (username, role, badge, start_date.isoformat(), end_date.isoformat())
    )
    conn.commit()
    conn.close()


def upsert_schedule_day(badge: str, d: date, status: str, shift_type: Optional[str], source: str):
    """
    Upsert de un día en schedules.
    """
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    try:
        # Intentar UPDATE primero
        cursor.execute(
            "UPDATE schedules SET status = ?, shift_type = ? WHERE badge = ? AND date = ? AND source = ?",
            (status, shift_type, badge, d.isoformat(), source)
        )
        if cursor.rowcount == 0:
            cursor.execute(
                "INSERT INTO schedules (badge, date, status, shift_type, source) VALUES (?, ?, ?, ?, ?)",
                (badge, d.isoformat(), status, shift_type, source)
            )
        conn.commit()
    finally:
        conn.close()


def upsert_schedule_range(badge: str, start_d: date, end_d: date,
                          status: str, shift_type: Optional[str], source: str) -> int:
    """
    Marca ON/ON NS/OFF por rango [start_d, end_d]. Devuelve cuántos días se escribieron.
    """
    total = 0
    d = start_d
    while d <= end_d:
        upsert_schedule_day(badge, d, status, shift_type, source)
        total += 1
        d += timedelta(days=1)
    return total


def clear_schedule_range(badge: str, start_d: date, end_d: date, source: str) -> int:
    """Elimina (limpia) estado día-a-día en rango."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute(
        "DELETE FROM schedules WHERE badge = ? AND source = ? AND date >= ? AND date <= ?",
        (badge, source, start_d.isoformat(), end_d.isoformat())
    )
    deleted = cursor.rowcount if cursor.rowcount is not None else 0
    conn.commit()
    conn.close()
    return deleted


def get_schedule_map_for_range(badge: str, start_d: date, end_d: date, source: str) -> Dict[str, Dict]:
    """Devuelve { 'YYYY-MM-DD': {'status':..., 'shift_type':...} } para el rango indicado."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute(
        "SELECT date, status, shift_type FROM schedules WHERE badge = ? AND source = ? AND date >= ? AND date <= ?",
        (badge, source, start_d.isoformat(), end_d.isoformat())
    )
    res = {row['date']: {'status': row['status'], 'shift_type': row['shift_type']} for row in cursor.fetchall()}
    conn.close()
    return res


def get_schedules_for_source(source: str) -> List[Dict]:
    """Lista completa de schedules para un source."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute(
        "SELECT badge, date, status, shift_type, source FROM schedules WHERE source = ? ORDER BY date",
        (source,)
    )
    res = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return res


def get_all_operations() -> List[Dict]:
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT id, username, role, badge, start_date, end_date FROM operations ORDER BY id DESC")
    res = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return res
