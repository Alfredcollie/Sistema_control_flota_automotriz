# -*- coding: utf-8 -*-
"""
KEEPALIVE_SUPABASE.PY
=====================
Mantiene "activo" el proyecto gratuito de Supabase para que NO se pause por
inactividad.

¿Por qué hace falta?
    Supabase pausa los proyectos del plan Free cuando detectan poca actividad
    de base de datos durante 7 días seguidos. La documentación oficial indica
    que "unos pocos pedidos de usuario a la base de datos cada día" alcanzan
    para que el proyecto no se pause.

¿Qué hace este script?
    Abre una conexión a la base de datos (el mismo pooler que usa la app) y
    ejecuta varias consultas reales del usuario. Eso cuenta como actividad.

Uso:
    python keepalive_supabase.py

Credenciales (en este orden):
    1) Variables de entorno: SUPABASE_DB_HOST, SUPABASE_DB_PORT,
       SUPABASE_DB_NAME, SUPABASE_DB_USER, SUPABASE_DB_PASSWORD
    2) Archivo config_db.json junto a este script
    3) Llavero del sistema (keyring), si está instalado

Códigos de salida:
    0 = OK (el proyecto respondió)
    2 = El proyecto está pausado / no se pudo conectar
    3 = Faltan credenciales
"""
import json
import os
import sys
import time
from datetime import datetime, timezone

try:  # evita errores de acentos en consolas Windows (cp1252)
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

SERVICE_NAME = "ControlFlota"
CONSULTAS = 3          # pedidos "de usuario" a la base de datos
PAUSA_ENTRE = 2.0      # segundos entre consultas

# Bundle de certificados raíz (macOS / PyInstaller no traen CA del sistema)
try:
    import certifi
    _CA_BUNDLE = certifi.where()
except Exception:
    _CA_BUNDLE = None


LOG_ARCHIVO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "keepalive_log.txt")
LOG_MAX_LINEAS = 300


def _log(msg):
    """Imprime el mensaje y lo agrega al archivo keepalive_log.txt."""
    linea = f"{datetime.now():%Y-%m-%d %H:%M:%S} | {msg}"
    print(linea, flush=True)
    if os.environ.get("KEEPALIVE_SIN_LOG"):
        return
    try:
        lineas = []
        if os.path.exists(LOG_ARCHIVO):
            with open(LOG_ARCHIVO, "r", encoding="utf-8", errors="replace") as f:
                lineas = f.read().splitlines()
        lineas.append(linea)
        if len(lineas) > LOG_MAX_LINEAS:
            lineas = lineas[-LOG_MAX_LINEAS:]
        with open(LOG_ARCHIVO, "w", encoding="utf-8") as f:
            f.write("\n".join(lineas) + "\n")
    except Exception:
        pass


def _credenciales_entorno():
    return {
        "host": os.environ.get("SUPABASE_DB_HOST", ""),
        "port": str(os.environ.get("SUPABASE_DB_PORT", "")),
        "dbname": os.environ.get("SUPABASE_DB_NAME", ""),
        "user": os.environ.get("SUPABASE_DB_USER", ""),
        "password": os.environ.get("SUPABASE_DB_PASSWORD", ""),
    }


def _credenciales_archivo():
    base = os.path.dirname(os.path.abspath(__file__))
    ruta = os.path.join(base, "config_db.json")
    if not os.path.exists(ruta):
        return {"host": "", "port": "", "dbname": "", "user": "", "password": ""}
    try:
        with open(ruta, "r", encoding="utf-8-sig") as f:
            datos = json.load(f)
        return {
            "host": datos.get("host", ""),
            "port": str(datos.get("port", "")),
            "dbname": datos.get("dbname", ""),
            "user": datos.get("user", ""),
            "password": datos.get("password", ""),
        }
    except Exception as e:
        _log(f"AVISO: no se pudo leer config_db.json ({e})")
        return {"host": "", "port": "", "dbname": "", "user": "", "password": ""}


def _credenciales_llavero():
    try:
        import keyring
        return {
            "host": keyring.get_password(SERVICE_NAME, "SUPABASE_DB_HOST") or "",
            "port": str(keyring.get_password(SERVICE_NAME, "SUPABASE_DB_PORT") or ""),
            "dbname": keyring.get_password(SERVICE_NAME, "SUPABASE_DB_NAME") or "",
            "user": keyring.get_password(SERVICE_NAME, "SUPABASE_DB_USER") or "",
            "password": keyring.get_password(SERVICE_NAME, "SUPABASE_DB_PASSWORD") or "",
        }
    except Exception:
        return {"host": "", "port": "", "dbname": "", "user": "", "password": ""}


def leer_credenciales():
    env = _credenciales_entorno()
    arch = _credenciales_archivo()
    ll = _credenciales_llavero()
    return {
        "host": env["host"] or arch["host"] or ll["host"],
        "port": env["port"] or arch["port"] or ll["port"],
        "dbname": env["dbname"] or arch["dbname"] or ll["dbname"],
        "user": env["user"] or arch["user"] or ll["user"],
        "password": env["password"] or arch["password"] or ll["password"],
    }


def mantener_activo():
    try:
        import psycopg2
    except ImportError:
        _log("ERROR: falta psycopg2. Instálalo con: pip install psycopg2-binary")
        return 2

    cred = leer_credenciales()
    if not (cred["host"] and cred["user"] and cred["password"]):
        _log("ERROR: faltan credenciales (define SUPABASE_DB_* o crea config_db.json).")
        return 3

    conexion = {
        "host": cred["host"],
        "port": int(cred["port"] or 5432),
        "dbname": cred["dbname"] or "postgres",
        "user": cred["user"],
        "password": cred["password"],
        "connect_timeout": 15,
        "sslmode": "require",
    }
    _log(f"Conectando a {cred['host']}:{conexion['port']} (proyecto {cred['user'].split('.')[-1][:8]}...)")

    con = None
    if _CA_BUNDLE and os.path.exists(_CA_BUNDLE):
        # Igual que conexion.py: si certifi no valida la cadena del pooler,
        # se reintenta con sslmode=require (cifrado sin verificar la CA).
        try:
            con = psycopg2.connect(**{**conexion, "sslrootcert": _CA_BUNDLE})
        except Exception as e:
            _log(f"AVISO: certifi no validó el certificado ({str(e)[:80]}); "
                 "reintentando con sslmode=require.")
    if con is None:
        try:
            con = psycopg2.connect(**conexion)
        except Exception as e:
            msg = str(e)
            _log(f"ERROR de conexión: {msg}")
            bajo = msg.lower()
            if any(p in bajo for p in ("tenant", "not found", "no such", "tenant/user", "project")):
                _log("PISTA: el proyecto parece estar PAUSADO. Entra al dashboard de "
                     "Supabase y pulsa 'Resume project'.")
            return 2

    try:
        con.autocommit = True
        with con.cursor() as cur:
            for i in range(1, CONSULTAS + 1):
                cur.execute("SELECT now(), current_user")
                fila = cur.fetchone()
                _log(f"Consulta {i}/{CONSULTAS} OK -> {fila[0]:%Y-%m-%d %H:%M:%S %Z} (usuario {fila[1]})")
                if i < CONSULTAS:
                    time.sleep(PAUSA_ENTRE)
        _log("Proyecto ACTIVO: la actividad quedó registrada en Supabase.")
        return 0
    except Exception as e:
        _log(f"ERROR al ejecutar las consultas: {e}")
        return 2
    finally:
        try:
            con.close()
        except Exception:
            pass


if __name__ == "__main__":
    sys.exit(mantener_activo())
