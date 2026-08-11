# -*- coding: utf-8 -*-
import psycopg2

def conectar_db(silencioso=False):
    """
    Motor de conexión puro para la nube. 
    Cero interfaces gráficas (Sin Tkinter).
    """
    try:
        conn = psycopg2.connect(
            host="aws-0-us-east-2.pooler.supabase.com", 
            database="postgres",      
            user="postgres.lnmuzwlxcxewobdvlgrh",        
            password="Ruc20613993682*",   
            port="5432"
        )
        return conn
    except Exception as e:
        if not silencioso:
            # En la nube usamos print() en lugar de messagebox
            print(f"❌ Error crítico de conexión a Supabase: {e}")
        return None

def registrar_auditoria(usuario, modulo, accion):
    """
    Función de respaldo por si algún otro módulo la llama.
    Totalmente silenciosa.
    """
    conn = conectar_db(silencioso=True)
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO auditoria (usuario, modulo, accion) VALUES (%s, %s, %s)",
                (usuario, modulo, accion)
            )
            conn.commit()
        except Exception as e:
            print(f"⚠️ Error interno en auditoría: {e}")
        finally:
            conn.close()