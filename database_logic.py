import sqlite3
from datetime import date, timedelta
from typing import Tuple, List, Dict, Optional

DB_FILE = 'transporte_operaciones.db'


def setup_database():
    """Create database tables if they do not exist and run lightweight migrations."""
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
            status TEXT NOT NULL,            -- 'ON', 'ON NS', 'OFF' o CODIGO personalizado (p.ej. 'SOP')
            shift_type TEXT,                 -- 'Day Shift' | 'Night Shift' | Nombre del tipo personalizado | NULL
            source TEXT NOT NULL,
            in_time TEXT,                    -- HH:MM (para tipos personalizados)
            out_time TEXT,                   -- HH:MM (para tipos personalizados)
            UNIQUE (badge, date, source)
    )''')

    # --- migración blanda: agregar columnas si faltan (SQLite acepta ADD COLUMN múltiples veces con try/except) ---
    try:
        cursor.execute("ALTER TABLE schedules ADD COLUMN in_time TEXT")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE schedules ADD COLUMN out_time TEXT")
    except sqlite3.OperationalError:
        pass

    # -------------------------
    # audit_log
    # -------------------------
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,          -- quien realizó la acción
            source TEXT NOT NULL,            -- RGM | Newmont | Administrator
            action_type TEXT NOT NULL,       -- USER_LOGIN | SHIFT_MODIFICATION | DATA_EXPORT | DATA_IMPORT | SHIFT_TYPE_* ...
            detail TEXT,
            ts TEXT NOT NULL DEFAULT (datetime('now'))
    )''')

    # -------------------------
    # shift_types (nueva)
    # -------------------------
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS shift_types (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source TEXT NOT NULL,           -- RGM | Newmont (ámbito del tipo)
            name TEXT NOT NULL,             -- único por source
            code TEXT NOT NULL,             -- único por source (p.ej. 'SOP')
            color_hex TEXT NOT NULL,        -- '#RRGGBB'
            in_time TEXT NOT NULL,          -- 'HH:MM' 24h
            out_time TEXT NOT NULL,         -- 'HH:MM' 24h
            UNIQUE (source, name),
            UNIQUE (source, code)
    )''')

    # Indexes útiles
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_users_source ON users(source)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_schedules_source ON schedules(source)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_schedules_badge ON schedules(badge)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_schedules_date ON schedules(date)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_audit_ts ON audit_log(ts)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_shift_types_source ON shift_types(source)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_shift_types_code ON shift_types(code)")

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


def upsert_schedule_day(
    badge: str,
    d: date,
    status: str,
    shift_type: Optional[str],
    source: str,
    in_time: Optional[str] = None,
    out_time: Optional[str] = None
):
    """
    Upsert de un día en schedules. Soporta:
      - estados base: 'ON'/'ON NS'/'OFF' (in_time/out_time pueden ser None)
      - tipos personalizados: status=code, shift_type=name, in_time/out_time HH:MM
    """
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    try:
        # UPDATE primero
        cursor.execute(
            "UPDATE schedules SET status = ?, shift_type = ?, in_time = ?, out_time = ? "
            "WHERE badge = ? AND date = ? AND source = ?",
            (status, shift_type, in_time, out_time, badge, d.isoformat(), source)
        )
        if cursor.rowcount == 0:
            cursor.execute(
                "INSERT INTO schedules (badge, date, status, shift_type, source, in_time, out_time) "
                "VALUES (?, ?, ?, ?, ?, ?, ?)",
                (badge, d.isoformat(), status, shift_type, source, in_time, out_time)
            )
        conn.commit()
    finally:
        conn.close()


def upsert_schedule_range(
    badge: str,
    start_d: date,
    end_d: date,
    status: str,
    shift_type: Optional[str],
    source: str,
    in_time: Optional[str] = None,
    out_time: Optional[str] = None
) -> int:
    """
    Marca por rango [start_d, end_d]. Devuelve cuántos días se escribieron.
    """
    total = 0
    d = start_d
    while d <= end_d:
        upsert_schedule_day(badge, d, status, shift_type, source, in_time, out_time)
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
    """Devuelve { 'YYYY-MM-DD': {'status':..., 'shift_type':..., 'in_time':..., 'out_time':...} } para el rango."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute(
        "SELECT date, status, shift_type, in_time, out_time "
        "FROM schedules WHERE badge = ? AND source = ? AND date >= ? AND date <= ?",
        (badge, source, start_d.isoformat(), end_d.isoformat())
    )
    res = {
        row['date']: {
            'status': row['status'],
            'shift_type': row['shift_type'],
            'in_time': row['in_time'],
            'out_time': row['out_time']
        }
        for row in cursor.fetchall()
    }
    conn.close()
    return res


def get_schedules_for_source(source: str) -> List[Dict]:
    """Lista completa de schedules para un source."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute(
        "SELECT badge, date, status, shift_type, source, in_time, out_time "
        "FROM schedules WHERE source = ? ORDER BY date",
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


# ---------------------------------------------------------------------
# Shift Types (CRUD + helpers)
# ---------------------------------------------------------------------
def get_shift_types(source: str) -> List[Dict]:
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute(
        "SELECT id, source, name, code, color_hex, in_time, out_time "
        "FROM shift_types WHERE source = ? ORDER BY name",
        (source,)
    )
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


def get_shift_type_map(source: str) -> Dict[str, Dict]:
    """
    Devuelve {code_upper: {'name':..., 'color_hex':..., 'in_time':..., 'out_time':...}}
    """
    types = get_shift_types(source)
    return {
        t['code'].strip().upper(): {
            'name': t['name'],
            'color_hex': t['color_hex'],
            'in_time': t['in_time'],
            'out_time': t['out_time']
        }
        for t in types
    }


def create_shift_type(source: str, name: str, code: str, color_hex: str, in_time: str, out_time: str) -> Tuple[bool, str]:
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    try:
        cur.execute(
            "INSERT INTO shift_types (source, name, code, color_hex, in_time, out_time) VALUES (?,?,?,?,?,?)",
            (source, name.strip(), code.strip().upper(), color_hex.strip(), in_time.strip(), out_time.strip())
        )
        conn.commit()
        return True, "Shift type created."
    except sqlite3.IntegrityError as e:
        return False, f"Error: name/code already exists for {source}."
    except sqlite3.Error as e:
        return False, f"Database error: {e}"
    finally:
        conn.close()


def update_shift_type(
    type_id: int,
    source: str,
    name: str,
    code: str,
    color_hex: str,
    in_time: str,
    out_time: str
) -> Tuple[bool, str, Optional[str], Optional[str]]:
    """
    Actualiza un tipo de turno. Si el código cambia, actualiza TODAS las asignaciones en schedules
    (status viejo -> status nuevo) para el mismo source. Devuelve (ok, msg, old_code, new_code).
    """
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    try:
        cur.execute("SELECT code FROM shift_types WHERE id = ? AND source = ?", (type_id, source))
        row = cur.fetchone()
        if not row:
            return False, "Shift type not found.", None, None
        old_code = row[0]
        new_code = code.strip().upper()

        # Verificar unicidad (name/code) excepto el propio registro
        cur.execute("SELECT id FROM shift_types WHERE source=? AND name=? AND id != ?", (source, name.strip(), type_id))
        if cur.fetchone():
            return False, "Error: another shift type with the same name already exists.", None, None
        cur.execute("SELECT id FROM shift_types WHERE source=? AND code=? AND id != ?", (source, new_code, type_id))
        if cur.fetchone():
            return False, "Error: another shift type with the same code already exists.", None, None

        # Update shift_types
        cur.execute(
            "UPDATE shift_types SET name=?, code=?, color_hex=?, in_time=?, out_time=? WHERE id=? AND source=?",
            (name.strip(), new_code, color_hex.strip(), in_time.strip(), out_time.strip(), type_id, source)
        )

        # Si cambió el código, propagar a schedules
        if old_code != new_code:
            cur.execute(
                "UPDATE schedules SET status=? WHERE status=? AND source=?",
                (new_code, old_code, source)
            )

        conn.commit()
        return True, "Shift type updated.", old_code, new_code
    except sqlite3.Error as e:
        return False, f"Database error: {e}", None, None
    finally:
        conn.close()


def delete_shift_type(type_id: int) -> Tuple[bool, str, Optional[str], Optional[str]]:
    """
    Intenta eliminar; si está en uso, lo impide.
    Devuelve (ok, msg, source, code) para facilitar mensajes y acciones.
    """
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    try:
        cur.execute("SELECT source, code, name FROM shift_types WHERE id=?", (type_id,))
        row = cur.fetchone()
        if not row:
            return False, "Shift type not found.", None, None
        source, code, name = row[0], row[1], row[2]

        # Regla crítica: impedir eliminación si está asignado
        cur.execute("SELECT COUNT(1) FROM schedules WHERE source=? AND status=?", (source, code))
        cnt = cur.fetchone()[0]
        if cnt and int(cnt) > 0:
            return (
                False,
                f"No se puede eliminar el tipo de turno '{name}' porque está asignado a uno o más empleados. "
                f"Reasigne primero esos turnos.",
                source,
                code
            )

        cur.execute("DELETE FROM shift_types WHERE id=?", (type_id,))
        conn.commit()
        return True, "Shift type deleted.", source, code
    except sqlite3.Error as e:
        return False, f"Database error: {e}", None, None
    finally:
        conn.close()
