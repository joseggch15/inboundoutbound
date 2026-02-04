import sqlite3
import os
from database_logic import setup_database, DB_FILE

def limpiar_base_datos():
    print(f"Base de datos objetivo: {DB_FILE}")
    confirmacion = input("⚠️  ADVERTENCIA: Esto BORRARÁ TODAS las tablas y datos.\n¿Estás seguro? (escribe 'SI'): ")
    
    if confirmacion.strip().upper() != 'SI':
        print("Operación cancelada.")
        return

    # 1. Intentar borrar el archivo físico primero (la opción más limpia)
    file_removed = False
    if os.path.exists(DB_FILE):
        try:
            os.remove(DB_FILE)
            print(f"🗑️  Archivo '{DB_FILE}' eliminado del disco.")
            file_removed = True
        except PermissionError:
            print("⚠️  No se pudo borrar el archivo (está en uso). Se intentará limpiar vía SQL.")
        except Exception as e:
            print(f"⚠️  Error al borrar archivo: {e}")

    # 2. Si no se pudo borrar el archivo, usamos SQL para borrar tablas
    if not file_removed:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        
        tablas = [
            "schedules",
            "operations",
            "user_locations",
            "users",
            "audit_log",
            "shift_types",  
            "location",   
            "report_settings",
            "sqlite_sequence" # Para reiniciar los autoincrement
        ]

        try:
            # Desactivamos llaves foráneas temporalmente para evitar conflictos al borrar
            cursor.execute("PRAGMA foreign_keys = OFF;")
            
            for tabla in tablas:
                # Usamos DROP TABLE IF EXISTS para que no falle si la tabla no existe
                cursor.execute(f"DROP TABLE IF EXISTS {tabla}")
                print(f"Tabla '{tabla}' eliminada (si existía).")
            
            conn.commit()
            cursor.execute("VACUUM") # Optimizar espacio
            print("🧹 Limpieza SQL completada.")

        except Exception as e:
            print(f"❌ Error SQL: {e}")
        finally:
            conn.close()

    # 3. REGENERAR LA ESTRUCTURA (CRÍTICO)
    print("🛠️  Regenerando estructura de la base de datos...")
    try:
        setup_database()
        print(f"\n✅ ÉXITO: Base de datos '{DB_FILE}' reseteada y lista para usar.")
    except Exception as e:
        print(f"\n❌ Error al regenerar la base de datos: {e}")

if __name__ == '__main__':
    limpiar_base_datos()