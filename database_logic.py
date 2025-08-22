import sqlite3
from datetime import date

DB_NAME = "transporte_operaciones.db"

def setup_database():
    """
    Prepara la base de datos. Crea la tabla 'operations' con la estructura correcta
    solo si no existe. Esta es una operación segura para ejecutar cada vez que
    inicia la aplicación.
    """
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
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
    conn.commit()
    conn.close()

def add_operation(username: str, role: str, badge: str, start_date: date, end_date: date):
    """Añade o reemplaza una operación en la base de datos usando el BADGE como clave única."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    # Usamos INSERT OR REPLACE para simplificar. Si el badge ya existe, actualiza la fila.
    # Si no existe, inserta una nueva.
    cursor.execute(
        """
        INSERT OR REPLACE INTO operations (id, username, role, badge, start_date, end_date)
        VALUES ((SELECT id FROM operations WHERE badge = ?), ?, ?, ?, ?, ?)
        """,
        (badge, username, role, badge, start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d'))
    )
    conn.commit()
    conn.close()

def get_all_operations() -> list:
    """Obtiene todas las operaciones de la base de datos, ordenadas por nombre."""
    conn = sqlite3.connect(DB_NAME)
    # sqlite3.Row permite acceder a los resultados por nombre de columna
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT id, username, role, badge, start_date, end_date FROM operations ORDER BY username")
    records = [dict(row) for row in cursor.fetchall()] # Convertir a lista de diccionarios
    conn.close()
    return records

def get_operations_for_report(report_start: date, report_end: date) -> list:
    """
    Obtiene las operaciones activas dentro de un rango de fechas.
    Una operación se considera activa si su período se superpone con el rango del reporte.
    """
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    # La lógica de la consulta asegura que cualquier rotación que se cruce con el
    # rango del reporte sea incluida.
    cursor.execute(
        "SELECT * FROM operations WHERE start_date <= ? AND end_date >= ?",
        (report_end.strftime('%Y-%m-%d'), report_start.strftime('%Y-%m-%d'))
    )
    records = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return records