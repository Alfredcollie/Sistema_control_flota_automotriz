# -*- coding: utf-8 -*-
"""
CONEXION.PY (v3 SEGURA + OPTIMIZADA)
- Credenciales de Supabase en llavero del sistema (keyring).
- Pool de Conexiones Persistente (ThreadedConnectionPool).
- Auditoría Asíncrona (Background Threading).
"""
import logging
import psycopg2
from psycopg2 import pool
import keyring
import threading
from datetime import datetime

# Bundle de certificados raíz (necesario para SSL a Supabase en macOS/PyInstaller)
try:
    import certifi
    _CA_BUNDLE = certifi.where()
except Exception:
    _CA_BUNDLE = None

SERVICE_NAME = "ControlFlota"

# =========================================================
# ⚙️ CREDENCIALES DE RESPALDO (Supabase)
# Se usan solo si el llavero del sistema está vacío o no es
# accesible (p. ej. en un .app empaquetado de macOS, donde el
# Keychain no tiene las credenciales configuradas). Así la app
# conecta igual en Windows y macOS sin configuración extra.
# (Mismos valores que usa validacion_licencia.py)
# =========================================================
SUPABASE_HOST = "aws-1-us-west-2.pooler.supabase.com"
SUPABASE_DB_NAME = "postgres"
SUPABASE_USER = "postgres.nqjfptmupnrkmgvnbyly"
SUPABASE_PASSWORD = "Ve-10339092"
SUPABASE_PORT = "6543"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] conexion_supabase: %(message)s"
)

# Variable global para el Pool de conexiones
_connection_pool = None


def _credenciales_llavero():
    """Credenciales leídas del llavero del sistema (None si no existen)."""
    def _get(clave):
        try:
            return keyring.get_password(SERVICE_NAME, clave)
        except Exception:
            return None
    return {
        "host": _get("SUPABASE_DB_HOST"),
        "port": _get("SUPABASE_DB_PORT"),
        "dbname": _get("SUPABASE_DB_NAME"),
        "user": _get("SUPABASE_DB_USER"),
        "password": _get("SUPABASE_DB_PASSWORD"),
    }


def _credenciales_respaldo():
    """Credenciales fijas de Supabase (respaldo multiplataforma)."""
    return {
        "host": SUPABASE_HOST,
        "port": SUPABASE_PORT,
        "dbname": SUPABASE_DB_NAME,
        "user": SUPABASE_USER,
        "password": SUPABASE_PASSWORD,
    }


def leer_credenciales():
    """Credenciales del llavero; si falta algún campo, completa con el respaldo."""
    ll = _credenciales_llavero()
    r = _credenciales_respaldo()
    return {
        "host": ll["host"] or r["host"],
        "port": ll["port"] or r["port"],
        "dbname": ll["dbname"] or r["dbname"],
        "user": ll["user"] or r["user"],
        "password": ll["password"] or r["password"],
    }


def _crear_pool(cred):
    """Crea el ThreadedConnectionPool con unas credenciales dadas (o None)."""
    if not cred["host"] or not cred["user"] or not cred["password"]:
        return None
    kwargs = {
        "minconn": 1,
        "maxconn": 15,
        "host": cred["host"],
        "port": int(cred["port"]),
        "database": cred["dbname"],
        "user": cred["user"],
        "password": cred["password"],
        "connect_timeout": 10,
        # Supabase exige SSL. En macOS (app de PyInstaller) no hay CA raíz del
        # sistema, por eso forzamos SSL y usamos el bundle de certifi.
        "sslmode": "require",
    }
    if _CA_BUNDLE:
        kwargs["sslrootcert"] = _CA_BUNDLE
    return psycopg2.pool.ThreadedConnectionPool(**kwargs)


def inicializar_pool(silencioso=False):
    """Inicializa el pool de conexiones persistentes.

    Intenta primero con las credenciales del llavero y, si la conexión falla
    (p. ej. credenciales viejas/incorrectas en el Keychain de macOS), reintenta
    automáticamente con las credenciales de respaldo de Supabase."""
    global _connection_pool
    if _connection_pool is not None:
        return

    cred = leer_credenciales()
    try:
        _connection_pool = _crear_pool(cred)
        if _connection_pool is not None:
            return
    except Exception as e:
        if not silencioso:
            logging.error(f"Error al conectar con credenciales del llavero: {e}")
        _connection_pool = None

    # Reintento con el respaldo puro de Supabase (solo si no es idéntico al intento anterior)
    respaldo = _credenciales_respaldo()
    if cred == respaldo:
        return
    try:
        _connection_pool = _crear_pool(respaldo)
        if _connection_pool is not None and not silencioso:
            logging.info("Conectado con credenciales de respaldo de Supabase.")
    except Exception as e:
        if not silencioso:
            logging.error(f"Error al conectar con credenciales de respaldo: {e}")
        _connection_pool = None


def conectar_db(silencioso=False):
    """Obtiene una conexión pre-creada del Pool en lugar de crear una nueva."""
    global _connection_pool
    if _connection_pool is None:
        inicializar_pool(silencioso)
        
    try:
        if _connection_pool:
            return _connection_pool.getconn()
    except Exception as e:
        if not silencioso:
            logging.error(f"Error al obtener conexión del pool: {e}")
    return None


def liberar_conexion(conn):
    """Devuelve la conexión al pool para que sea reutilizada por otro proceso."""
    global _connection_pool
    if _connection_pool and conn:
        try:
            _connection_pool.putconn(conn)
        except Exception:
            pass


def _tarea_auditoria_asincrona(usuario, modulo, accion):
    """Función interna que se ejecuta en un hilo separado (Background)."""
    conn = conectar_db(silencioso=True)
    if not conn:
        return
    try:
        ahora = datetime.now()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO bitacora_auditoria (fecha, hora, usuario, modulo, accion) VALUES (%s, %s, %s, %s, %s)",
            (ahora.strftime("%d/%m/%Y"), ahora.strftime("%H:%M:%S"), usuario, modulo, accion)
        )
        conn.commit()
    except Exception as e:
        logging.error(f"Error en auditoría asíncrona: {e}")
    finally:
        # IMPORTANTE: Liberamos la conexión en lugar de cerrarla
        liberar_conexion(conn)


def registrar_auditoria(usuario, modulo, accion):
    """Registra una acción en la bitácora sin congelar la interfaz (Asíncrono)."""
    if usuario in ["Desconocido", "Invitado", None]:
        return
    
    # Lanzamos la escritura a la base de datos en un "hilo" paralelo (Daemon)
    hilo = threading.Thread(
        target=_tarea_auditoria_asincrona, 
        args=(usuario, modulo, accion),
        daemon=True
    )
    hilo.start()


if __name__ == "__main__":
    c = conectar_db()
    if c:
        print("✅ Conexión (Pool) correcta leyendo desde el llavero del sistema.")
        liberar_conexion(c)
    else:
        print("❌ Sin conexión. Ejecuta primero: python configurar_credenciales.py")