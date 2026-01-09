import sqlite3

DB_FILE = "transporte_operaciones.db"

def limpiar_base_datos():
    confirmacion = input("¿Estás seguro de BORRAR TODO el contenido de la base de datos? (escribe 'SI'): ")
    if confirmacion != 'SI':
        print("Cancelado.")
        return

    # Conectamos
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    # Lista de tablas a vaciar
    tablas = [
        "schedules",
        "operations",
        "user_locations",
        "users",
        "audit_log",
        # "shift_types",  <-- Descomenta si quieres borrar tipos de turno
        # "location",     <-- Descomenta si quieres borrar lugares
        # "report_settings" 
    ]

    try:
        # 1. Borrar contenido de las tablas
        for tabla in tablas:
            cursor.execute(f"DELETE FROM {tabla}")
            print(f"Tabla '{tabla}' vaciada.")
        
        # 2. Reiniciar contadores de ID (Autoincrement)
        # Esto asegura que el próximo usuario sea ID 1, no ID 500
        cursor.execute("DELETE FROM sqlite_sequence")
        print("Contadores de ID reiniciados.")

        # 3. CRÍTICO: Guardar cambios (Commit) ANTES de VACUUM
        # Esto cierra la transacción abierta por los DELETE
        conn.commit() 
        
        # 4. Optimizar espacio (VACUUM)
        # Para evitar el error, nos aseguramos de estar en modo autocommit
        conn.isolation_level = None 
        cursor.execute("VACUUM")
        
        print("\n✅ Base de datos limpiada, IDs reiniciados y archivo optimizado.")

    except Exception as e:
        print(f"❌ Error: {e}")
    finally:
        conn.close()

if __name__ == '__main__':
    limpiar_base_datos()