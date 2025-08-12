# database_logic.py
import sqlite3
from datetime import date

DB_NAME = "transporte_operaciones.db"

def setup_database():
    """Prepara la base de datos, asegurando que la tabla y columnas existan."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    # Aseguramos que la tabla, la columna badge y el índice único existan.
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS operations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            role TEXT,
            badge TEXT UNIQUE,
            start_date DATE NOT NULL,
            end_date DATE NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    try:
        cursor.execute("ALTER TABLE operations ADD COLUMN badge TEXT")
        cursor.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_badge ON operations (badge)")
    except sqlite3.OperationalError as e:
        if "duplicate column name" not in str(e):
            raise e
    conn.commit()
    conn.close()

def add_operation(username: str, role: str, badge: str, start_date: date, end_date: date):
    """Añade o reemplaza una operación en la base de datos usando el BADGE como clave."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT OR REPLACE INTO operations (id, username, role, badge, start_date, end_date) "
        "VALUES ((SELECT id FROM operations WHERE badge = ?), ?, ?, ?, ?, ?)",
        (badge, username, role, badge, start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d'))
    )
    conn.commit()
    conn.close()

def get_all_operations() -> list:
    """Obtiene todas las operaciones de la base de datos."""
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT id, username, role, badge, start_date, end_date FROM operations ORDER BY username")
    records = cursor.fetchall()
    conn.close()
    return records

def get_operations_for_report(report_start: date, report_end: date) -> list:
    """Obtiene las operaciones relevantes para el reporte de transporte."""
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute(
        "SELECT * FROM operations WHERE start_date <= ? AND end_date >= ?",
        (report_end.strftime('%Y-%m-%d'), report_start.strftime('%Y-%m-%d'))
    )
    records = cursor.fetchall()
    conn.close()
    return records