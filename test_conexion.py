# -*- coding: utf-8 -*-
from conexion import conectar_db, liberar_conexion


def probar_conexion():
    """Prueba la conexión al pool de Supabase y libera la conexión correctamente."""
    conn = conectar_db()
    if conn is None:
        print("No se pudo conectar.")
        return
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT 1")
        print("Conexión OK:", cursor.fetchone())
    except Exception as e:
        print("Error:", e)
    finally:
        liberar_conexion(conn)


if __name__ == "__main__":
    probar_conexion()
