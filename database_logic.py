# Basado y extendido a partir del mÃ³dulo original. Referencia: :contentReference[oaicite:0]{index=0}
import sqlite3
import json
from datetime import date, timedelta, datetime
from typing import Tuple, List, Dict, Optional, Set

#DB_FILE = "transporte_operaciones.db"

# --- POR ESTO (Ruta absoluta segura) ---
import sys
from pathlib import Path

def _app_dir() -> Path:
    """Devuelve la carpeta donde corre el script o el .exe"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

DB_FILE = str(_app_dir() / "transporte_operaciones.db")

def setup_database():
    """Crea todas las tablas necesarias si no existen."""
    # Conectamos a la base de datos
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    # 1. Tabla de Usuarios
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            role TEXT,
            badge TEXT UNIQUE NOT NULL,
            source TEXT NOT NULL
        )
    """)
    
    # 2. Tabla de Roles
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS roles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source TEXT NOT NULL,
            name TEXT NOT NULL,
            UNIQUE (source, name)
        )""")
    
    # Migración automática de roles existentes en usuarios
    try:
        cursor.execute(
            "INSERT OR IGNORE INTO roles (source, name) "
            "SELECT source, role FROM users WHERE role IS NOT NULL AND role != ''"
        )
    except Exception:
        pass

    # 3. File Registry
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS file_registry (
            source TEXT NOT NULL,
            file_key TEXT NOT NULL,
            path TEXT NOT NULL,
            updated_by TEXT,
            updated_at TEXT NOT NULL DEFAULT (datetime('now')),
            PRIMARY KEY (source, file_key)
        )""")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_file_registry_source ON file_registry(source)")

    # 4. Tabla Locations (Corregida y consolidada aquí)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS location (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source TEXT NOT NULL,
            pickup_location TEXT NOT NULL,
            UNIQUE (source, pickup_location)
        )""")
    
    # Migración de esquema antiguo de location
    try:
        cols = [r[1] for r in cursor.execute("PRAGMA table_info(location)").fetchall()]
        if "source" not in cols:
            cursor.execute("ALTER TABLE location RENAME TO location_old")
            cursor.execute("""
                CREATE TABLE location (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    source TEXT NOT NULL,
                    pickup_location TEXT NOT NULL,
                    UNIQUE (source, pickup_location)
                )
            """)
            cursor.execute("INSERT INTO location (source, pickup_location) SELECT 'RGM', pickup_location FROM location_old")
            cursor.execute("INSERT OR IGNORE INTO location (source, pickup_location) SELECT 'Newmont', pickup_location FROM location_old")
            cursor.execute("DROP TABLE location_old")
    except sqlite3.OperationalError:
        pass

    cursor.execute("CREATE INDEX IF NOT EXISTS idx_location_source ON location(source)")

    # 5. Tabla User Locations
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS user_locations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            badge TEXT NOT NULL,
            start_date TEXT NOT NULL,
            end_date TEXT NOT NULL,
            pickup_location TEXT,
            dropoff_location TEXT,
            is_default INTEGER NOT NULL DEFAULT 0
        )""")
    
    # Migraciones suaves para user_locations
    try: cursor.execute("ALTER TABLE user_locations ADD COLUMN dropoff_location TEXT")
    except sqlite3.OperationalError: pass
    try: cursor.execute("ALTER TABLE user_locations ADD COLUMN is_default INTEGER NOT NULL DEFAULT 0")
    except sqlite3.OperationalError: pass

    cursor.execute("CREATE INDEX IF NOT EXISTS idx_ul_badge ON user_locations(badge)")

    # 6. Tabla Operations
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS operations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            role TEXT,
            badge TEXT,
            start_date TEXT NOT NULL,
            end_date TEXT NOT NULL,
            created_by TEXT,
            entry_date TEXT,
            exit_date TEXT
        )""")
    
    # Migraciones suaves para operations
    try: cursor.execute("ALTER TABLE operations ADD COLUMN created_by TEXT")
    except sqlite3.OperationalError: pass
    try: cursor.execute("ALTER TABLE operations ADD COLUMN entry_date TEXT")
    except sqlite3.OperationalError: pass
    try: cursor.execute("ALTER TABLE operations ADD COLUMN exit_date TEXT")
    except sqlite3.OperationalError: pass

    # 7. Tabla Schedules
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS schedules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            badge TEXT NOT NULL,
            date TEXT NOT NULL,
            status TEXT NOT NULL,
            shift_type TEXT,
            source TEXT NOT NULL,
            in_time TEXT,
            out_time TEXT,
            remark TEXT,
            force_new_entry INTEGER NOT NULL DEFAULT 0,
            UNIQUE (badge, date, source)
        )""")
    
    # Migraciones suaves para schedules
    try: cursor.execute("ALTER TABLE schedules ADD COLUMN force_new_entry INTEGER NOT NULL DEFAULT 0")
    except sqlite3.OperationalError: pass
    try: cursor.execute("ALTER TABLE schedules ADD COLUMN in_time TEXT")
    except sqlite3.OperationalError: pass
    try: cursor.execute("ALTER TABLE schedules ADD COLUMN out_time TEXT")
    except sqlite3.OperationalError: pass
    try: cursor.execute("ALTER TABLE schedules ADD COLUMN remark TEXT")
    except sqlite3.OperationalError: pass

    # 8. Tabla Audit Log
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            source TEXT NOT NULL,
            action_type TEXT NOT NULL,
            detail TEXT,
            ts TEXT NOT NULL DEFAULT (datetime('now'))
        )""")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_audit_ts ON audit_log(ts)")

    # 9. Tabla Shift Types
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS shift_types (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source TEXT NOT NULL,
            name TEXT NOT NULL,
            code TEXT NOT NULL,
            color_hex TEXT NOT NULL,
            in_time TEXT NOT NULL,
            out_time TEXT NOT NULL,
            is_off INTEGER DEFAULT 0,
            UNIQUE (source, name),
            UNIQUE (source, code)
        )""")
    try: cursor.execute("ALTER TABLE shift_types ADD COLUMN is_off INTEGER DEFAULT 0")
    except sqlite3.OperationalError: pass

    # 10. Report Settings
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS report_settings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            source TEXT NOT NULL,
            settings_json TEXT,
            UNIQUE(username, source)
        )
    """)
    try: cursor.execute("ALTER TABLE report_settings ADD COLUMN settings_json TEXT")
    except sqlite3.OperationalError: pass

    conn.commit()
    conn.close()
# --- Add these new CRUD functions at the end of database_logic.py ---

def get_roles(source: Optional[str] = None) -> List[Dict]:
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    if source:
        cur.execute("SELECT id, source, name FROM roles WHERE source=? ORDER BY name", (source,))
    else:
        cur.execute("SELECT id, source, name FROM roles ORDER BY source, name")
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows

def get_roles_filtered(source: Optional[str], text: Optional[str]) -> List[Dict]:
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    query = "SELECT id, source, name FROM roles"
    conditions = []
    params = []

    if source:
        conditions.append("source = ?")
        params.append(source)
    if text:
        conditions.append("name LIKE ?")
        params.append(f"%{text}%")
    
    if conditions:
        query += " WHERE " + " AND ".join(conditions)
    
    query += " ORDER BY name"
    cursor.execute(query, tuple(params))
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows

def create_role(name: str, source: str) -> Tuple[bool, str]:
    name = (name or "").strip()
    if not name:
        return False, "Role name cannot be empty."
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    try:
        cur.execute("INSERT INTO roles (source, name) VALUES (?,?)", (source, name))
        conn.commit()
        return True, f"Role '{name}' created."
    except sqlite3.IntegrityError:
        return False, f"Role '{name}' already exists for {source}."
    finally:
        conn.close()

def update_role(role_id: int, name: str, source: str) -> Tuple[bool, str]:
    name = (name or "").strip()
    if not name:
        return False, "Role name cannot be empty."
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    try:
        cur.execute("UPDATE roles SET name=? WHERE id=? AND source=?", (name, role_id, source))
        conn.commit()
        if cur.rowcount:
            return True, "Role updated."
        return False, "Role not found."
    except sqlite3.IntegrityError:
        return False, "Role name already exists."
    finally:
        conn.close()

def delete_role(role_id: int, source: str) -> Tuple[bool, str]:
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    try:
        # 1. Obtener el nombre del rol antes de borrar
        cur.execute("SELECT name FROM roles WHERE id=? AND source=?", (role_id, source))
        row = cur.fetchone()
        if not row:
            return False, "Role not found."
        role_name = row[0]

        # 2. Verificar si hay usuarios usando este rol
        cur.execute("SELECT COUNT(*) FROM users WHERE source=? AND role=?", (source, role_name))
        count = cur.fetchone()[0]
        
        if count > 0:
            return False, f"Cannot delete '{role_name}': It is assigned to {count} user(s)."

        # 3. Si no hay usuarios, proceder con el borrado
        cur.execute("DELETE FROM roles WHERE id=? AND source=?", (role_id, source))
        conn.commit()
        if cur.rowcount:
            return True, "Role deleted."
        return False, "Role not found."
    finally:
        conn.close()

  
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS location (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source TEXT NOT NULL,                -- RGM | Newmont
            pickup_location TEXT NOT NULL,
            UNIQUE (source, pickup_location)
        )""")
    
    

    # --- MigraciÃ³n desde esquema antiguo (sin 'source') ---
    try:
        cols = [r[1] for r in cursor.execute("PRAGMA table_info(location)").fetchall()]
        # Si encontramos una tabla 'location' sin 'source', la migramos:
        if "source" not in cols:  # tabla vieja
            cursor.execute("ALTER TABLE location RENAME TO location_old")
            cursor.execute(
                """
                CREATE TABLE location (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    source TEXT NOT NULL,
                    pickup_location TEXT NOT NULL,
                    UNIQUE (source, pickup_location)
                )
                """
            )
            # Duplicamos el catÃ¡logo previo para ambas empresas para no perder nada:
            cursor.execute("INSERT INTO location (source, pickup_location) SELECT 'RGM', pickup_location FROM location_old")
            cursor.execute("INSERT OR IGNORE INTO location (source, pickup_location) SELECT 'Newmont', pickup_location FROM location_old")
            cursor.execute("DROP TABLE location_old")
    except sqlite3.OperationalError:
        pass

    # Ãndices Ãºtiles
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_location_source ON location(source)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_location_src_name ON location(source, pickup_location)")


    # -------------------------
    # AsignaciÃ³n de ubicaciones por usuario y rango
    # -------------------------
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS user_locations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            badge TEXT NOT NULL,
            start_date TEXT NOT NULL,      -- YYYY-MM-DD
            end_date TEXT NOT NULL,        -- YYYY-MM-DD
            pickup_location TEXT,          -- texto tomado de location.pickup_location
            dropoff_location TEXT,         -- idem
            is_default INTEGER NOT NULL DEFAULT 0
        )"""
    )

    # Migraciones suaves por si faltan columnas nuevas
    try:
        cursor.execute("ALTER TABLE user_locations ADD COLUMN dropoff_location TEXT")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute(
            "ALTER TABLE user_locations ADD COLUMN is_default INTEGER NOT NULL DEFAULT 0"
        )
    except sqlite3.OperationalError:
        pass

    # Ãndices Ãºtiles
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_ul_badge ON user_locations(badge)")
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_ul_range ON user_locations(start_date, end_date)"
    )

    # -------------------------
    # Operations/rotations history (rangos informativos)
    # -------------------------
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS operations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            role TEXT,
            badge TEXT,
            start_date TEXT NOT NULL,
            end_date TEXT NOT NULL,
            created_by TEXT,
            entry_date TEXT,
            exit_date TEXT
        )
    """
    )
    # --- Soft migrations for new columns ---
    try:
        cursor.execute("ALTER TABLE operations ADD COLUMN created_by TEXT")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE operations ADD COLUMN entry_date TEXT")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE operations ADD COLUMN exit_date TEXT")
    except sqlite3.OperationalError:
        pass


    # -------------------------
    # schedules (estado dÃ­a a dÃ­a)
    # -------------------------
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS schedules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            badge TEXT NOT NULL,
            date TEXT NOT NULL,                  -- 'YYYY-MM-DD'
            status TEXT NOT NULL,                -- 'ON', 'ON NS', 'OFF' o CODIGO personalizado (p.ej. 'SOP')
            shift_type TEXT,                     -- 'Day Shift' | 'Night Shift' | Nombre del tipo personalizado | NULL
            source TEXT NOT NULL,
            in_time TEXT,                        -- HH:MM (para tipos personalizados)
            out_time TEXT,                       -- HH:MM (para tipos personalizados)
            remark TEXT,
            UNIQUE (badge, date, source)
        )"""
    )

    # --- migraciÃ³n blanda: agregar columnas si faltan (SQLite acepta ADD COLUMN mÃºltiples veces con try/except) ---
    try:
        cursor.execute("ALTER TABLE schedules ADD COLUMN in_time TEXT")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE schedules ADD COLUMN out_time TEXT")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE schedules ADD COLUMN remark TEXT")
    except sqlite3.OperationalError:
        pass

    # -------------------------
    # audit_log
    # -------------------------
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,              -- quien realizÃ³ la acciÃ³n
            source TEXT NOT NULL,                -- RGM | Newmont | Administrator
            action_type TEXT NOT NULL,           -- USER_LOGIN | SHIFT_MODIFICATION | DATA_EXPORT | DATA_IMPORT | SHIFT_TYPE_* ...
            detail TEXT,
            ts TEXT NOT NULL DEFAULT (datetime('now'))
        )"""
    )

    # -------------------------
    # shift_types (nueva)
    # -------------------------
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS shift_types (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source TEXT NOT NULL,                -- RGM | Newmont (Ã¡mbito del tipo)
            name TEXT NOT NULL,                  -- Ãºnico por source
            code TEXT NOT NULL,                  -- Ãºnico por source (p.ej. 'SOP')
            color_hex TEXT NOT NULL,             -- '#RRGGBB'
            in_time TEXT NOT NULL,               -- 'HH:MM' 24h
            out_time TEXT NOT NULL,
            is_off INTEGER DEFAULT 0,-- 'HH:MM' 24h
            UNIQUE (source, name),
            UNIQUE (source, code)
        )"""
    )
    
    try:
        cursor.execute("ALTER TABLE shift_types ADD COLUMN is_off INTEGER DEFAULT 0")
        print("MigraciÃ³n: Columna 'is_off' agregada a 'shift_types'.")
    except sqlite3.OperationalError:
        pass # La columna ya existe

    # -------------------------
    # Report Settings
    # -------------------------
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS report_settings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            source TEXT NOT NULL,
            settings_json TEXT,
            UNIQUE(username, source)
        )
    """)
    # --- MIGRATION: Add settings_json column if it doesn't exist ---
    try:
        cursor.execute("ALTER TABLE report_settings ADD COLUMN settings_json TEXT")
    except sqlite3.OperationalError as e:
        if "duplicate column name" not in str(e).lower():
            raise e


    # Indexes Ãºtiles
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_users_source ON users(source)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_schedules_source ON schedules(source)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_schedules_badge ON schedules(badge)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_schedules_date ON schedules(date)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_audit_ts ON audit_log(ts)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_shift_types_source ON shift_types(source)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_shift_types_code ON shift_types(code)")
    # TR-ST-01: Index for shift type usage status filter
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_schedules_source_status ON schedules(source, status)")

    conn.commit()
    conn.close()


# ---------------------------------------------------------------------
# Report Settings
# ---------------------------------------------------------------------
def get_report_settings(username: str, source: str) -> Dict:
    """Get report settings for a user and source, providing defaults if none exist."""
    
    # Default for Newmont based on the provided image
    if source == "Newmont":
        defaults = {
            "font_name": "Calibri",
            "font_color": "#000000",
            "header_bg_color": "#70AD47",  # Green from image
            "header_font_color": "#FFFFFF", # White
            "date_format": "dd-mmm-yy",     # Default date format for Newmont
            "column_colors": {}
        }
    # Default for RGM or any other source
    else:
        defaults = {
            "font_name": "Arial",
            "font_color": "#000000",
            "header_bg_color": "#4472C4",
            "header_font_color": "#FFFFFF",
            "date_format": "dd/mm/yyyy",      # Default date format for RGM
            "column_colors": {}
        }
        
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute(
        "SELECT settings_json FROM report_settings WHERE username = ? AND source = ?",
        (username, source)
    )
    row = cursor.fetchone()
    conn.close()
    if row and row[0]:
        try:
            settings = json.loads(row[0])
            # User's saved settings will override the defaults
            defaults.update(settings)
            return defaults
        except (json.JSONDecodeError, TypeError):
            return defaults
    return defaults

def save_report_settings(username: str, source: str, settings: Dict):
    """Save or update report settings for a user and source."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    settings_json = json.dumps(settings)
    cursor.execute(
        """
        INSERT INTO report_settings (username, source, settings_json)
        VALUES (?, ?, ?)
        ON CONFLICT(username, source) DO UPDATE SET settings_json = excluded.settings_json
        """,
        (username, source, settings_json)
    )
    conn.commit()
    conn.close()

def delete_report_settings(username: str, source: str):
    """Deletes the report settings for a specific user and source."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute(
        "DELETE FROM report_settings WHERE username = ? AND source = ?",
        (username, source)
    )
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
        (username or "Unknown", source or "", action_type or "", detail or ""),
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
            (source,),
        )
    else:
        cursor.execute(
            "SELECT ts, username, source, action_type, detail FROM audit_log "
            "ORDER BY ts DESC"
        )
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
            (name, role, badge, source),
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

    new_users = [
        user for user in users if str(user.get("badge")) not in existing_badges
    ]

    if not new_users:
        conn.close()
        return 0

    user_data = [
        (user["name"], user["role"], str(user["badge"]), source) for user in new_users
    ]

    added_count = 0
    try:
        cursor.executemany(
            "INSERT INTO users (name, role, badge, source) VALUES (?, ?, ?, ?)",
            user_data,
        )
        conn.commit()
        added_count = (
            cursor.rowcount if cursor.rowcount is not None else len(new_users)
        )
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
        (source,),
    )
    users = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return users

def get_users_filtered(source: str, text: Optional[str], role: Optional[str], badge_prefix: Optional[str], active_since: Optional[date] = None) -> list:
    """Get filtered list of users for a source."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    query = "SELECT id, name, role, badge FROM users WHERE source = ?"
    params: List = [source]

    if text:
        query += " AND (name LIKE ? OR badge LIKE ?)"
        params.extend([f"%{text}%", f"%{text}%"])
    
    if role:
        query += " AND role = ?"
        params.append(role)
        
    if badge_prefix:
        query += " AND badge LIKE ?"
        params.append(f"{badge_prefix}%")
        
    if active_since:
        query += " AND EXISTS (SELECT 1 FROM schedules s WHERE s.badge = users.badge AND s.source = users.source AND s.date >= ?)"
        params.append(active_since.isoformat())

    query += " ORDER BY name"
    
    cursor.execute(query, tuple(params))
    users = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return users


def update_user(
    user_id: int, name: str, role: str, badge: str, source: str
) -> Tuple[bool, str]:
    """Update an existing user's data."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    try:
        # Check if the new badge is already in use by ANOTHER user from the same source
        cursor.execute(
            "SELECT id FROM users WHERE badge = ? AND source = ? AND id != ?",
            (badge, source, user_id),
        )
        if cursor.fetchone():
            return False, f"Error: The badge '{badge}' is already assigned to another user."

        cursor.execute(
            "UPDATE users SET name = ?, role = ?, badge = ? WHERE id = ?",
            (name, role, badge, user_id),
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

# -------------------------
# Locations (CRUD) â€” con Ã¡mbito por 'source'
# -------------------------
def get_locations(source: Optional[str] = None) -> List[Dict]:
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    if source:
        cur.execute("SELECT id, source, pickup_location FROM location WHERE source=? ORDER BY pickup_location", (source,))
    else:
        cur.execute("SELECT id, source, pickup_location FROM location ORDER BY source, pickup_location")
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows

def get_locations_filtered(source: Optional[str], text: Optional[str], sort_by: str = 'name', unassigned_only: bool = False) -> List[Dict]:
    """Get filtered and sorted locations."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    query = "SELECT id, source, pickup_location FROM location l"
    conditions = []
    params: List = []

    if source:
        conditions.append("source = ?")
        params.append(source)
    
    if text:
        conditions.append("pickup_location LIKE ?")
        params.append(f"%{text}%")
    
    if unassigned_only:
        conditions.append("NOT EXISTS (SELECT 1 FROM user_locations ul WHERE ul.pickup_location = l.pickup_location AND ul.is_default = 1)")
    
    if conditions:
        query += " WHERE " + " AND ".join(conditions)

    if sort_by == 'source':
        query += " ORDER BY source, pickup_location"
    else: # name
        query += " ORDER BY pickup_location, source"

    cursor.execute(query, tuple(params))
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


def create_location(pickup_location: str, source: str) -> Tuple[bool, str]:
    pickup_location = (pickup_location or "").strip()
    if not pickup_location:
        return False, "Location name cannot be empty."
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    try:
        cur.execute("INSERT INTO location (source, pickup_location) VALUES (?,?)", (source, pickup_location))
        conn.commit()
        return True, f"Location created for {source}."
    except sqlite3.IntegrityError:
        return False, f"This location already exists for {source}."
    finally:
        conn.close()

def update_location(loc_id: int, pickup_location: str, source: str) -> Tuple[bool, str]:
    pickup_location = (pickup_location or "").strip()
    if not pickup_location:
        return False, "Location name cannot be empty."
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    try:
        # Solo actualiza si el registro pertenece al 'source' (seguridad por Ã¡mbito)
        cur.execute("UPDATE location SET pickup_location=? WHERE id=? AND source=?", (pickup_location, loc_id, source))
        conn.commit()
        if cur.rowcount:
            return True, "Location updated."
        return False, "Location not found for this company."
    finally:
        conn.close()

def delete_location(loc_id: int, source: str) -> Tuple[bool, str]:
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    try:
        # 1. Verificar nombre
        cur.execute("SELECT pickup_location FROM location WHERE id=? AND source=?", (loc_id, source))
        row = cur.fetchone()
        if not row:
            return False, "Location not found for this company."
        loc_name = row[0]

        # 2. Verificar uso en user_locations (pickup o dropoff)
        # Hacemos JOIN con users para asegurar que filtramos por el source correcto
        query = """
            SELECT COUNT(*) FROM user_locations ul
            JOIN users u ON u.badge = ul.badge
            WHERE u.source = ? AND (ul.pickup_location = ? OR ul.dropoff_location = ?)
        """
        cur.execute(query, (source, loc_name, loc_name))
        count = cur.fetchone()[0]

        if count > 0:
            return False, f"Cannot delete '{loc_name}': It is assigned to {count} user(s)."

        # 3. Borrar
        cur.execute("DELETE FROM location WHERE id=? AND source=?", (loc_id, source))
        conn.commit()
        if cur.rowcount:
            return True, "Location deleted."
        return False, "Location not found for this company."
    finally:
        conn.close()

# --- Variantes para Administrador (pueden cambiar 'source' o operar sin Ã¡mbito) ---
def update_location_admin(loc_id: int, pickup_location: str, new_source: str) -> Tuple[bool, str]:
    pickup_location = (pickup_location or "").strip()
    if not pickup_location:
        return False, "Location name cannot be empty."
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    try:
        cur.execute("UPDATE location SET pickup_location=?, source=? WHERE id=?", (pickup_location, new_source, loc_id))
        conn.commit()
        if cur.rowcount:
            return True, "Location updated (admin)."
        return False, "Location not found."
    except sqlite3.IntegrityError:
        return False, f"Another location with the same name already exists for {new_source}."
    finally:
        conn.close()

def delete_location_admin(loc_id: int) -> Tuple[bool, str]:
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    try:
        # 1. Verificar nombre y source
        cur.execute("SELECT source, pickup_location FROM location WHERE id=?", (loc_id,))
        row = cur.fetchone()
        if not row:
            return False, "Location not found."
        source, loc_name = row[0], row[1]

        # 2. Verificar uso
        query = """
            SELECT COUNT(*) FROM user_locations ul
            JOIN users u ON u.badge = ul.badge
            WHERE u.source = ? AND (ul.pickup_location = ? OR ul.dropoff_location = ?)
        """
        cur.execute(query, (source, loc_name, loc_name))
        count = cur.fetchone()[0]

        if count > 0:
            return False, f"Cannot delete '{loc_name}' ({source}): It is assigned to {count} user(s)."

        # 3. Borrar
        cur.execute("DELETE FROM location WHERE id=?", (loc_id,))
        conn.commit()
        if cur.rowcount:
            return True, "Location deleted (admin)."
        return False, "Location not found."
    finally:
        conn.close()


def assign_user_location_range(badge: str, start_date: date, end_date: date,
                               pickup: Optional[str], dropoff: Optional[str],
                               is_default: int = 0) -> None:
    """Inserta una asignaciÃ³n de pickup/dropoff para un rango de fechas."""
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO user_locations (badge, start_date, end_date, pickup_location, dropoff_location, is_default) "
        "VALUES (?,?,?,?,?,?)",
        (str(badge), start_date.isoformat(), end_date.isoformat(),
         (pickup or None), (dropoff or None), int(bool(is_default)))
    )
    conn.commit()
    conn.close()

def get_users_with_defaults(source: str) -> list:
    """
    Get all users joined with their default locations (is_default=1).
    Optimized for CRUD table rendering to avoid N+1 queries.
    """
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    # Realizamos un LEFT JOIN con user_locations filtrando por is_default=1
    # Esto nos da el usuario Y su configuraciÃ³n por defecto en una sola fila.
    query = """
        SELECT 
            u.id, u.name, u.role, u.badge,
            ul.pickup_location, ul.dropoff_location
        FROM users u
        LEFT JOIN user_locations ul 
            ON u.badge = ul.badge 
            AND ul.is_default = 1
        WHERE u.source = ?
        ORDER BY u.name
    """
    
    cursor.execute(query, (source,))
    # Convertimos a lista de diccionarios para facilitar el manejo en UI
    users = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return users

def set_user_default_locations(badge: str, pickup: Optional[str], dropoff: Optional[str]) -> None:
    """
    Define un default permanente (sin rango finito) para el usuario.
    Se implementa con is_default=1 y un rango amplio.
    """
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("DELETE FROM user_locations WHERE badge=? AND is_default=1", (str(badge),))
    cur.execute(
        "INSERT INTO user_locations (badge, start_date, end_date, pickup_location, dropoff_location, is_default) "
        "VALUES (?,?,?,?,?,1)",
        (str(badge), "1900-01-01", "9999-12-31", (pickup or None), (dropoff or None))
    )
    conn.commit()
    conn.close()

def get_user_location_for_date(badge: str, d: date) -> Tuple[Optional[str], Optional[str]]:
    """Busca primero una asignaciÃ³n de rango que cubra la fecha; si no existe, cae al default."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    iso = d.isoformat()

    # rango especÃ­fico
    cur.execute(
        "SELECT pickup_location, dropoff_location FROM user_locations "
        "WHERE badge=? AND is_default=0 AND start_date<=? AND end_date>=? "
        "ORDER BY id DESC LIMIT 1",
        (str(badge), iso, iso)
    )
    row = cur.fetchone()
    if row and (row["pickup_location"] or row["dropoff_location"]):
        conn.close()
        return row["pickup_location"], row["dropoff_location"]

    # default
    cur.execute(
        "SELECT pickup_location, dropoff_location FROM user_locations "
        "WHERE badge=? AND is_default=1 ORDER BY id DESC LIMIT 1",
        (str(badge),)
    )
    row = cur.fetchone()
    conn.close()
    if row:
        return row["pickup_location"], row["dropoff_location"]
    return None, None

def list_user_default_locations(source: str) -> List[Dict]:
    """Listado para UI (tabla por usuario con su default actual)."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute(
        "SELECT u.name, u.role, u.badge, "
        "       COALESCE(ul.pickup_location,'') AS pickup_location, "
        "       COALESCE(ul.dropoff_location,'') AS dropoff_location "
        "FROM users u "
        "LEFT JOIN user_locations ul ON ul.badge = u.badge AND ul.is_default = 1 "
        "WHERE u.source = ? "
        "ORDER BY u.name", (source,)
    )
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows



# ---------------------------------------------------------------------
# Operations & schedules
# ---------------------------------------------------------------------
def add_operation(
    username: str, 
    role: str, 
    badge: str, 
    start_date: date, 
    end_date: date, 
    created_by: str, 
    entry_date: Optional[datetime] = None, 
    exit_date: Optional[datetime] = None,
    cursor: Optional[sqlite3.Cursor] = None # <--- NEW ARGUMENT
):
    """
    MODIFIED: Handles datetime for entry/exit and supports external cursor.
    """
    should_close = False
    if cursor is None:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        should_close = True

    try:
        cursor.execute(
            "INSERT INTO operations (username, role, badge, start_date, end_date, created_by, entry_date, exit_date) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (
                username, role, badge,
                start_date.isoformat(), end_date.isoformat(),
                created_by,
                entry_date.strftime('%Y-%m-%d %H:%M') if entry_date else None,
                exit_date.strftime('%Y-%m-%d %H:%M') if exit_date else None
            ),
        )
        if should_close:
            cursor.connection.commit()
    except Exception:
        if should_close:
            cursor.connection.rollback()
        raise
    finally:
        if should_close:
            cursor.connection.close()


# EN database_logic.py (Agregar al final o en la sección de Operations)

def update_operation_exit_time_by_date(badge: str, target_date: date, new_exit_time: datetime.time, cursor=None):
    """
    Actualiza la hora de salida (exit_date) de la operación que cubre 'target_date'.
    Se usa cuando se rompe un imán (Separar Viaje) para definir la hora real de salida del viaje previo.
    Soporta cursor externo para uso dentro de transacciones activas (ej: paste).
    """
    should_close = False
    if cursor is None:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        should_close = True

    try:
        iso_date = target_date.isoformat()
        
        # 1. Buscar la operación activa en esa fecha
        cursor.execute(
            "SELECT id, exit_date FROM operations "
            "WHERE badge = ? AND start_date <= ? AND end_date >= ? "
            "ORDER BY id DESC LIMIT 1",
            (badge, iso_date, iso_date)
        )
        row = cursor.fetchone()
        
        if row:
            op_id = row[0]
            new_dt = datetime.combine(target_date, new_exit_time)
            
            cursor.execute(
                "UPDATE operations SET exit_date = ? WHERE id = ?",
                (new_dt.strftime('%Y-%m-%d %H:%M'), op_id)
            )
            if should_close:
                cursor.connection.commit()
            return True, f"Updated exit time for OP #{op_id}"
            
        return False, "No active operation found for this date."
        
    except Exception as e:
        return False, f"DB Error: {e}"
    finally:
        if should_close:
            cursor.connection.close()

def upsert_schedule_day(
    badge: str,
    d: date,
    status: str,
    shift_type: Optional[str],
    source: str,
    in_time: Optional[str] = None,
    out_time: Optional[str] = None,
    remark: Optional[str] = None,
    force_new_entry: Optional[int] = None,
    cursor: Optional[sqlite3.Cursor] = None,  # <--- NEW ARGUMENT
):
    """
    Upsert de un dia en schedules.
    Updated to support external transactions via 'cursor'.
    force_new_entry: None = no tocar en UPDATE (preservar valor existente),
                     0/1 = setear explicitamente.
    """
    # Normalizar force_new_entry a 0/1 (o None para "no tocar" en UPDATE)
    if force_new_entry is None:
        fne_db = None
    else:
        fne_db = 1 if bool(force_new_entry) else 0

    should_close = False
    
    # If no cursor passed, we open a new connection (Legacy behavior)
    if cursor is None:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        should_close = True

    try:
        # UPDATE first — COALESCE(?, force_new_entry) preserva si fne_db es None
        cursor.execute(
            "UPDATE schedules SET status = ?, shift_type = ?, in_time = ?, out_time = ?, remark = ?, "
            "    force_new_entry = COALESCE(?, force_new_entry) "
            "WHERE badge = ? AND date = ? AND source = ?",
            (status, shift_type, in_time, out_time, remark, fne_db, badge, d.isoformat(), source),
        )
        if cursor.rowcount == 0:
            fne_insert = 1 if bool(force_new_entry) else 0
            cursor.execute(
                "INSERT INTO schedules (badge, date, status, shift_type, source, in_time, out_time, remark, force_new_entry) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (badge, d.isoformat(), status, shift_type, source, in_time, out_time, remark, fne_insert),
            )
        
        # Only commit if we opened the connection ourselves
        if should_close:
            cursor.connection.commit()
            
    except Exception:
        if should_close:
            cursor.connection.rollback()
        raise
    finally:
        if should_close:
            cursor.connection.close()


def upsert_schedule_range(
    badge: str,
    start_d: date,
    end_d: date,
    status: str,
    shift_type: Optional[str],
    source: str,
    in_time: Optional[str] = None,
    out_time: Optional[str] = None,
    remark: Optional[str] = None,
    force_new_entry_start: Optional[int] = None,
    cursor: Optional[sqlite3.Cursor] = None,  # <--- NEW ARGUMENT
) -> int:
    """
    Marca por rango [start_d, end_d]. Devuelve cuantos dias se escribieron.
    CORREGIDO: Propagación correcta de force_new_entry para limpiar residuos.
    """
    total = 0
    d = start_d
    
    # If no cursor provided, wrap the ENTIRE loop in a single transaction
    should_close = False
    if cursor is None:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("BEGIN TRANSACTION") # Explicit transaction for speed
        should_close = True

    try:
        while d <= end_d:
            # ────────────────────────────────────────────────────────
            # CORRECCIÓN: Lógica estricta para Unir/Separar
            # ────────────────────────────────────────────────────────
            if force_new_entry_start is None:
                # No se pidió cambio explícito → preservar lo que haya en DB (None)
                fne = None
            elif force_new_entry_start == 0:
                # "Unir" → limpiar forzosamente TODOS los días del rango a 0
                fne = 0
            else:
                # "Separar" (1) → marcar solo el primer día como 1, limpiar el resto a 0
                fne = 1 if d == start_d else 0

            # Pass the cursor to the day function
            upsert_schedule_day(
                badge, d, status, shift_type, source, 
                in_time, out_time, remark,
                force_new_entry=fne,
                cursor=cursor 
            )
            total += 1
            d += timedelta(days=1)
        
        if should_close:
            cursor.connection.commit()
            
    except Exception:
        if should_close:
            cursor.connection.rollback()
        raise
    finally:
        if should_close:
            cursor.connection.close()
    
    return total


def clear_schedule_range(badge: str, start_d: date, end_d: date, source: str) -> int:
    """Elimina (limpia) estado dÃ­a-a-dÃ­a en rango."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute(
        "DELETE FROM schedules WHERE badge = ? AND source = ? AND date >= ? AND date <= ?",
        (badge, source, start_d.isoformat(), end_d.isoformat()),
    )
    deleted = cursor.rowcount if cursor.rowcount is not None else 0
    conn.commit()
    conn.close()
    return deleted


def get_schedule_map_for_range(
    badge: str, start_d: date, end_d: date, source: str, 
    cursor: Optional[sqlite3.Cursor] = None
) -> Dict[str, Dict]:
    """Devuelve mapa de schedules para el rango. Soporta cursor externo."""
    should_close = False
    if cursor is None:
        conn = sqlite3.connect(DB_FILE)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        should_close = True

    try:
        cursor.execute(
            "SELECT date, status, shift_type, in_time, out_time, remark, force_new_entry "
            "FROM schedules WHERE badge = ? AND source = ? AND date >= ? AND date <= ?",
            (badge, source, start_d.isoformat(), end_d.isoformat()),
        )
        res = {}
        for row in cursor.fetchall():
            # Cuando usamos cursor externo, row puede ser tuple (no Row)
            # Necesitamos manejar ambos formatos
            if isinstance(row, dict):
                res[row["date"]] = {
                    "status": row["status"],
                    "shift_type": row["shift_type"],
                    "in_time": row["in_time"],
                    "out_time": row["out_time"],
                    "remark": row["remark"],
                    "force_new_entry": row["force_new_entry"],
                }
            else:
                # sqlite3.Row soporta indexado por nombre
                res[row["date"]] = {
                    "status": row["status"],
                    "shift_type": row["shift_type"],
                    "in_time": row["in_time"],
                    "out_time": row["out_time"],
                    "remark": row["remark"],
                    "force_new_entry": row["force_new_entry"],
                }
        return res
    finally:
        if should_close:
            cursor.connection.close()

def get_schedules_for_source(source: str) -> List[Dict]:
    """Lista completa de schedules para un source."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute(
        "SELECT badge, date, status, shift_type, source, in_time, out_time, remark "
        "FROM schedules WHERE source = ? ORDER BY date",
        (source,),
    )
    res = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return res


def get_all_operations() -> List[Dict]:
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id, username, role, badge, start_date, end_date, created_by, entry_date, exit_date FROM operations ORDER BY id DESC"
    )
    res = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return res

def get_operations_filtered(text: Optional[str] = None, role: Optional[str] = None, d_from: Optional[date] = None, d_to: Optional[date] = None, sort_by: str = 'start_date_desc', created_by: Optional[str] = None) -> List[Dict]:
    """ Get filtered list of operations history. """
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    query = "SELECT id, username, role, badge, start_date, end_date, created_by, entry_date, exit_date FROM operations"
    conditions = []
    params: List = []

    if text:
        conditions.append("(username LIKE ? OR badge LIKE ? OR role LIKE ?)")
        params.extend([f"%{text}%", f"%{text}%", f"%{text}%"])

    if role:
        conditions.append("role = ?")
        params.append(role)
    
    if d_from and d_to:
        # Overlap logic: (StartA <= EndB) and (EndA >= StartB)
        conditions.append("(start_date <= ? AND end_date >= ?)")
        params.extend([d_to.isoformat(), d_from.isoformat()])
    
    if created_by:
        conditions.append("created_by = ?")
        params.append(created_by)

    if conditions:
        query += " WHERE " + " AND ".join(conditions)

    if sort_by == 'name_asc':
        query += " ORDER BY username ASC"
    else: # start_date_desc
        query += " ORDER BY start_date DESC"

    cursor.execute(query, tuple(params))
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
        "SELECT id, source, name, code, color_hex, in_time, out_time, is_off "
        "FROM shift_types WHERE source = ? ORDER BY name",
        (source,),
    )
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows

def get_shift_types_filtered(source: str, text: Optional[str], in_from: Optional[str], in_to: Optional[str], usage: Optional[str]) -> List[Dict]:
    """Get filtered shift types for a source."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    query = "SELECT id, source, name, code, color_hex, in_time, out_time FROM shift_types st WHERE source = ?"
    params: List = [source]

    if text:
        query += " AND (name LIKE ? OR code LIKE ?)"
        params.extend([f"%{text}%", f"%{text}%"])
    
    if in_from and in_to:
        query += " AND in_time BETWEEN ? AND ?"
        params.extend([in_from, in_to])
        
    if usage == "In use":
        query += " AND EXISTS (SELECT 1 FROM schedules s WHERE s.source = st.source AND s.status = st.code)"
    elif usage == "Not in use":
        query += " AND NOT EXISTS (SELECT 1 FROM schedules s WHERE s.source = st.source AND s.status = st.code)"

    query += " ORDER BY name"
    
    cursor.execute(query, tuple(params))
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows


def get_shift_type_map(source: str) -> Dict[str, Dict]:
    """
    Retorna un diccionario mapeando CODE -> {name, color, times, is_off}.
    Crucial para que los reportes de Excel sepan distinguir dÃ­as libres customizados.
    """
    types = get_shift_types(source)
    return {
        t["code"].strip().upper(): {
            "name": t["name"],
            "color_hex": t["color_hex"],
            "in_time": t["in_time"],
            "out_time": t["out_time"],
            # Aseguramos que retorne un booleano (0 o 1 en SQLite -> False o True)
            "is_off": bool(t.get("is_off", 0)) 
        }
        for t in types
    }


# ACTUALIZAR create_shift_type
def create_shift_type(
    source: str, name: str, code: str, color_hex: str, in_time: str, out_time: str, is_off: bool = False
) -> Tuple[bool, str]:
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    try:
        cur.execute(
            "INSERT INTO shift_types (source, name, code, color_hex, in_time, out_time, is_off) VALUES (?,?,?,?,?,?,?)",
            (
                source,
                name.strip(),
                code.strip().upper(),
                color_hex.strip(),
                in_time.strip(),
                out_time.strip(),
                1 if is_off else 0, # Guardar como entero
            ),
        )
        conn.commit()
        return True, "Shift type created."
    except sqlite3.IntegrityError:
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
    out_time: str,
    is_off: bool # Nuevo parÃ¡metro
) -> Tuple[bool, str, Optional[str], Optional[str]]:
    """
    Actualiza un tipo de turno. Si el cÃ³digo cambia, actualiza TODAS las asignaciones en schedules
    (status viejo -> status nuevo) para el mismo source. Devuelve (ok, msg, old_code, new_code).
    """
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    try:
        cur.execute(
            "SELECT code FROM shift_types WHERE id = ? AND source = ?", (type_id, source)
        )
        row = cur.fetchone()
        if not row:
            return False, "Shift type not found.", None, None
        old_code = row[0]
        new_code = code.strip().upper()

        # Verificar unicidad (name/code) excepto el propio registro
        cur.execute(
            "SELECT id FROM shift_types WHERE source=? AND name=? AND id != ?",
            (source, name.strip(), type_id),
        )
        if cur.fetchone():
            return (
                False,
                "Error: another shift type with the same name already exists.",
                None,
                None,
            )
        cur.execute(
            "SELECT id FROM shift_types WHERE source=? AND code=? AND id != ?",
            (source, new_code, type_id),
        )
        if cur.fetchone():
            return (
                False,
                "Error: another shift type with the same code already exists.",
                None,
                None,
            )

        # Update shift_types
        cur.execute(
            "UPDATE shift_types SET name=?, code=?, color_hex=?, in_time=?, out_time=?, is_off=? WHERE id=? AND source=?",
            (
                name.strip(),
                new_code,
                color_hex.strip(),
                in_time.strip(),
                out_time.strip(),
                1 if is_off else 0,
                type_id,
                source,
            ),
        )

        # Si cambiÃ³ el cÃ³digo, propagar a schedules
        if old_code != new_code:
            cur.execute(
                "UPDATE schedules SET status=? WHERE status=? AND source=?",
                (new_code, old_code, source),
            )

        conn.commit()
        return True, "Shift type updated.", old_code, new_code
    except sqlite3.Error as e:
        return False, f"Database error: {e}", None, None
    finally:
        conn.close()


def delete_shift_type(type_id: int) -> Tuple[bool, str, Optional[str], Optional[str]]:
    """
    Intenta eliminar; si estÃ¡ en uso, lo impide.
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

        # Regla crÃ­tica: impedir eliminaciÃ³n si estÃ¡ asignado
        cur.execute(
            "SELECT COUNT(1) FROM schedules WHERE source=? AND status=?", (source, code)
        )
        cnt = cur.fetchone()[0]
        if cnt and int(cnt) > 0:
            return (
                False,
                f"No se puede eliminar el tipo de turno '{name}' porque estÃ¡ asignado a uno o mÃ¡s empleados. "
                f"Reasigne primero esos turnos.",
                source,
                code,
            )

        cur.execute("DELETE FROM shift_types WHERE id=?", (type_id,))
        conn.commit()
        return True, "Shift type deleted.", source, code
    except sqlite3.Error as e:
        return False, f"Database error: {e}", None, None
    finally:
        conn.close()
        
def get_operation_overlapping(badge: str, check_date: date, cursor: Optional[sqlite3.Cursor] = None) -> Optional[Dict]:
    """
    Busca si existe una operaciÃ³n activa que cubra una fecha especÃ­fica.
    Soporta cursor externo para evitar bloqueos (database is locked).
    """
    should_close = False
    
    # Si no nos pasan un cursor, abrimos una conexiÃ³n propia (comportamiento original)
    if cursor is None:
        conn = sqlite3.connect(DB_FILE)
        conn.row_factory = sqlite3.Row  # Importante para poder convertir a dict despuÃ©s
        cursor = conn.cursor()
        should_close = True

    try:
        iso = check_date.isoformat()
        # Buscamos una operaciÃ³n donde start <= date <= end
        cursor.execute(
            "SELECT * FROM operations WHERE badge = ? AND start_date <= ? AND end_date >= ? "
            "ORDER BY id DESC LIMIT 1",
            (badge, iso, iso)
        )
        row = cursor.fetchone()
        
        # Convertimos a diccionario si encontramos datos
        return dict(row) if row else None
        
    finally:
        # Solo cerramos la conexiÃ³n si NOSOTROS la abrimos.
        # Si vino de fuera (transacciÃ³n), la dejamos abierta.
        if should_close:
            cursor.connection.close()

def delete_operations_in_range(
    badge: str, 
    start_d: date, 
    end_d: date,
    cursor: Optional[sqlite3.Cursor] = None # <--- NEW ARGUMENT
):
    """
    Elimina cualquier operaciÃ³n que estÃ© TOTAL o PARCIALMENTE contenida en el rango.
    Supports external cursor.
    """
    should_close = False
    if cursor is None:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        should_close = True

    try:
        cursor.execute(
            """
            DELETE FROM operations 
            WHERE badge = ? 
              AND (
                (start_date >= ? AND start_date <= ?) OR 
                (end_date >= ? AND end_date <= ?) OR
                (start_date <= ? AND end_date >= ?)
              )
            """,
            (badge, start_d.isoformat(), end_d.isoformat(), 
             start_d.isoformat(), end_d.isoformat(),
             start_d.isoformat(), end_d.isoformat())
        )
        if should_close:
            cursor.connection.commit()
    except Exception:
        if should_close:
            cursor.connection.rollback()
        raise
    finally:
        if should_close:
            cursor.connection.close()


# ---------------------------------------------------------------------
# Shift Collision Detector — query para reportes
# ---------------------------------------------------------------------
def get_force_new_entry_map(source: str, start_d: date, end_d: date) -> Dict[str, Set[str]]:
    """
    Devuelve los dias marcados con force_new_entry=1 en el rango,
    agrupados por badge.
    Retorna: { 'BADGE1': {'YYYY-MM-DD', ...}, 'BADGE2': {...} }
    """
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    try:
        cur.execute(
            """
            SELECT badge, date
              FROM schedules
             WHERE source = ?
               AND date >= ?
               AND date <= ?
               AND force_new_entry = 1
            """,
            (source, start_d.isoformat(), end_d.isoformat()),
        )
        res: Dict[str, Set[str]] = {}
        for row in cur.fetchall():
            b = str(row["badge"]).strip()
            d = str(row["date"]).strip()
            if b and d:
                res.setdefault(b, set()).add(d)
        return res
    except sqlite3.OperationalError:
        # Columna aun no existe (DB sin migrar) -> retornar vacio
        return {}
    finally:
        conn.close()


def is_working_status(status: str, source: str) -> bool:
    """Helper rápido para saber si un status cuenta como día de trabajo (para fusionar)."""
    if not status:
        return False
    su = status.strip().upper()
    if su in ("OFF", "BREAK", "KO", "LEAVE", ""):
        return False
    
    # Chequear si es un custom type marcado como 'is_off'
    types = get_shift_types(source)
    for t in types:
        if str(t.get("code", "")).strip().upper() == su and t.get("is_off"):
            return False
            
    return True

def get_file_path(source: str, file_key: str) -> Optional[str]:
    """Recupera la ruta guardada de un archivo."""
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("SELECT path FROM file_registry WHERE source=? AND file_key=?", (source, file_key))
    row = cur.fetchone()
    conn.close()
    return row[0] if row and row[0] else None

def set_file_path(source: str, file_key: str, path: str, updated_by: Optional[str] = None) -> None:
    """Guarda o actualiza la ruta de un archivo."""
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute(
        """
        INSERT OR REPLACE INTO file_registry (source, file_key, path, updated_by, updated_at)
        VALUES (?, ?, ?, ?, datetime('now'))
        """,
        (source, file_key, path, updated_by),
    )
    conn.commit()
    conn.close()