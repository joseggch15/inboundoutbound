import sqlite3
import os
from database_logic import setup_database, DB_FILE

def limpiar_base_datos():
    print(f"🎯 Base de datos objetivo: {DB_FILE}")
    print("--- MODO: LIMPIEZA PROFUNDA (CONSERVANDO ADMINS) ---")
    print("Se borrarán: Schedules, Operaciones, Ubicaciones, Roles y Usuarios de prueba.")
    print("⚠️  SE CONSERVARÁN: 'javierteheran', 'miguelvenegas', 'admin' y sus configuraciones.")
    
    confirmacion = input("\n¿Estás seguro de reiniciar los datos para el TEST? (escribe 'SI'): ")
    
    if confirmacion.strip().upper() != 'SI':
        print("Operación cancelada.")
        return

    # Nos conectamos. NO borramos el archivo físico para poder salvar a los admins.
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    try:
        # Desactivamos llaves foráneas para poder borrar sin errores de dependencias
        cursor.execute("PRAGMA foreign_keys = OFF;")

        # ---------------------------------------------------------
        # 1. TABLAS A VACIAR COMPLETAMENTE (Datos transaccionales)
        # ---------------------------------------------------------
        tablas_a_vaciar = [
            "schedules",        # Turnos día a día
            "operations",       # Historial de rotaciones
            "user_locations",   # Asignaciones de transporte
            "location",         # Maestro de puntos de recogida (solicitaste borrarlo)
            "roles",            # Maestro de roles (se regenerará solo después)
            "audit_log",        # Logs
            # "shift_types"     # OPCIONAL: Descomenta si quieres borrar los tipos de turno también
        ]

        print("\n🧹 Limpiando tablas transaccionales...")
        for tabla in tablas_a_vaciar:
            try:
                cursor.execute(f"DELETE FROM {tabla}")
                # Reiniciamos el contador de ID (autoincrement) para esta tabla
                cursor.execute(f"DELETE FROM sqlite_sequence WHERE name='{tabla}'")
                print(f"   -> Tabla '{tabla}' vaciada.")
            except sqlite3.OperationalError:
                # Si la tabla no existe aún, no pasa nada
                pass

        # ---------------------------------------------------------
        # 2. LIMPIEZA SELECTIVA DE USUARIOS (El núcleo de tu petición)
        # ---------------------------------------------------------
        print("🧹 Limpiando Usuarios (protegiendo a Admin/RGM/Newmont)...")
        
        # Palabras clave que NO se deben borrar
        keywords = ['javierteheran', 'admin', 'miguelvenegas', 'Javier', 'Miguel']
        
        # Construimos la query: "Borrar todo usuario que NO tenga estos nombres/badges"
        query_users = "DELETE FROM users WHERE "
        conditions = []
        params = []
        
        for kw in keywords:
            # La condición es: name NOT LIKE '%kw%' AND badge NOT LIKE '%kw%'
            conditions.append(f"(name NOT LIKE ? AND badge NOT LIKE ?)")
            params.append(f"%{kw}%")
            params.append(f"%{kw}%")
        
        # Unimos todas las condiciones con AND (debe cumplir TODAS las negaciones para ser borrado)
        full_query = query_users + " AND ".join(conditions)
        
        cursor.execute(full_query, tuple(params))
        deleted_count = cursor.rowcount
        print(f"   -> {deleted_count} usuarios de prueba eliminados.")

        # ---------------------------------------------------------
        # 3. LIMPIEZA SELECTIVA DE SETTINGS
        # ---------------------------------------------------------
        # Borramos configuraciones de reporte de usuarios que ya no existen
        cursor.execute("""
            DELETE FROM report_settings 
            WHERE username NOT IN (SELECT name FROM users) 
              AND username NOT IN (SELECT badge FROM users)
        """)
        print("   -> Configuraciones huérfanas eliminadas.")

        conn.commit()
        
        # Optimizar base de datos (compactar archivo)
        cursor.execute("VACUUM")
        print("✨ Base de datos optimizada (VACUUM).")

    except Exception as e:
        print(f"❌ Error SQL durante la limpieza: {e}")
        conn.rollback()
    finally:
        conn.close()

    # ---------------------------------------------------------
    # 4. REGENERAR ESTRUCTURA Y DATOS BASE
    # ---------------------------------------------------------
    print("\n🛠️  Verificando estructura y regenerando básicos...")
    try:
        # Esto es vital: setup_database() contiene un script que dice:
        # "INSERT OR IGNORE INTO roles ... SELECT role FROM users"
        # Como no borramos a Javier ni a Miguel, sus roles volverán a llenar la tabla 'roles' automáticamente.
        setup_database()
        print(f"✅ ÉXITO: Base de datos lista para el test de cero.")
    except Exception as e:
        print(f"❌ Error al regenerar estructura: {e}")

if __name__ == '__main__':
    limpiar_base_datos()