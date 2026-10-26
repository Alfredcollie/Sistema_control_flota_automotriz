# -*- coding: utf-8 -*-
"""KEEPALIVE DE LA BASE DE DATOS "Licencias Colliesoftware" (Supabase).

¿Para qué sirve?
    Supabase (plan gratuito) PAUSA el proyecto cuando pasa ~7 días sin
    actividad de base de datos. Este script abre una conexión real y ejecuta
    varias consultas de usuario: eso cuenta como actividad y evita la pausa.

¿Qué base usa?
    La MISMA que valida las licencias del sistema (proyecto "Licencias
    Colliesoftware" en Supabase): aws-1-us-west-2.pooler.supabase.com:6543.

Credenciales (en este orden):
    1) Variables de entorno:
         LICENCIAS_DB_HOST / LICENCIAS_DB_PORT / LICENCIAS_DB_NAME
         LICENCIAS_DB_USER / LICENCIAS_DB_PASSWORD
       (también acepta los nombres SUPABASE_LIC_DB_*)
    2) Archivo config_licencia.json junto a este script
    3) Llavero del sistema (servicio "ControlFlotaLicencias")

Uso:
    python keepalive_licencias.py

Códigos de salida:
    0 = OK (el proyecto respondió)
    2 = El proyecto está pausado / no se pudo conectar
    3 = Faltan credenciales
"""
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

try:                      # acentos en consolas Windows (cp1252)
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

SERVICE_NAME = "ControlFlotaLicencias"     # mismo servicio que usa validacion_licencia.py
CONSULTAS = 3                              # pedidos "de usuario" a la base
PAUSA_ENTRE = 2.0                          # segundos entre consultas

try:                                       # CA bundle para macOS / PyInstaller
    import certifi
    _CA_BUNDLE = certifi.where()
except Exception:
    _CA_BUNDLE = None

BASE = Path(__file__).resolve().parent
ARCHIVO_CONFIG = BASE / "config_licencia.json"
LOG_ARCHIVO = BASE / "keepalive_licencias_log.txt"
LOG_MAX_LINEAS = 300


def _log(mensaje):
    """Imprime el mensaje y lo agrega al archivo de registro."""
    linea = "%s | %s" % (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), mensaje)
    print(linea, flush=True)
    if os.environ.get("KEEPALIVE_SIN_LOG"):
        return
    try:
        lineas = []
        if LOG_ARCHIVO.is_file():
            lineas = LOG_ARCHIVO.read_text(encoding="utf-8", errors="replace").splitlines()
        lineas.append(linea)
        if len(lineas) > LOG_MAX_LINEAS:
            lineas = lineas[-LOG_MAX_LINEAS:]
        LOG_ARCHIVO.write_text("\n".join(lineas) + "\n", encoding="utf-8")
    except Exception:
        pass


def _desde_entorno():
    return {
        "host": os.environ.get("LICENCIAS_DB_HOST") or os.environ.get("SUPABASE_LIC_DB_HOST", ""),
        "port": str(os.environ.get("LICENCIAS_DB_PORT") or os.environ.get("SUPABASE_LIC_DB_PORT", "")),
        "dbname": os.environ.get("LICENCIAS_DB_NAME") or os.environ.get("SUPABASE_LIC_DB_NAME", ""),
        "user": os.environ.get("LICENCIAS_DB_USER") or os.environ.get("SUPABASE_LIC_DB_USER", ""),
        "password": (os.environ.get("LICENCIAS_DB_PASSWORD")
                     or os.environ.get("SUPABASE_LIC_DB_PASSWORD", "")),
    }


def _desde_archivo():
    vacio = {"host": "", "port": "", "dbname": "", "user": "", "password": ""}
    if not ARCHIVO_CONFIG.is_file():
        return vacio
    try:
        datos = json.loads(ARCHIVO_CONFIG.read_text(encoding="utf-8-sig"))
        return {
            "host": datos.get("host", ""),
            "port": str(datos.get("port", "")),
            "dbname": datos.get("dbname", ""),
            "user": datos.get("user", ""),
            "password": datos.get("password", ""),
        }
    except Exception as e:
        _log("AVISO: no se pudo leer config_licencia.json (%s)" % e)
        return vacio


def _desde_llavero():
    vacio = {"host": "", "port": "", "dbname": "", "user": "", "password": ""}
    try:
        import keyring
        # la app guarda la contraseña como "SUPABASE_DB_PASSWORD"; se aceptan ambos nombres
        contrasena = (keyring.get_password(SERVICE_NAME, "SUPABASE_DB_PASSWORD")
                      or keyring.get_password(SERVICE_NAME, "LICENCIAS_DB_PASSWORD") or "")
        return {
            "host": keyring.get_password(SERVICE_NAME, "LICENCIAS_DB_HOST") or "",
            "port": str(keyring.get_password(SERVICE_NAME, "LICENCIAS_DB_PORT") or ""),
            "dbname": keyring.get_password(SERVICE_NAME, "LICENCIAS_DB_NAME") or "",
            "user": keyring.get_password(SERVICE_NAME, "LICENCIAS_DB_USER") or "",
            "password": contrasena,
        }
    except Exception:
        return vacio


def leer_credenciales():
    entorno, archivo, llavero = _desde_entorno(), _desde_archivo(), _desde_llavero()
    return {clave: entorno[clave] or archivo[clave] or llavero[clave]
            for clave in ("host", "port", "dbname", "user", "password")}


def mantener_activo():
    try:
        import psycopg2
    except ImportError:
        _log("ERROR: falta psycopg2. Instálelo con: pip install psycopg2-binary")
        return 2

    cred = leer_credenciales()
    if not (cred["host"] and cred["user"] and cred["password"]):
        _log("ERROR: faltan credenciales (defina LICENCIAS_DB_* o cree config_licencia.json).")
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
    proyecto = cred["user"].split(".")[-1]
    _log("Conectando a %s:%s (proyecto Licencias Colliesoftware %s...)"
         % (cred["host"], conexion["port"], proyecto[:8]))

    con = None
    if _CA_BUNDLE and Path(_CA_BUNDLE).is_file():
        try:
            con = psycopg2.connect(**dict(conexion, sslrootcert=_CA_BUNDLE))
        except Exception as e:
            _log("AVISO: certifi no validó el certificado (%s); reintentando con sslmode=require."
                 % str(e)[:80])
    if con is None:
        try:
            con = psycopg2.connect(**conexion)
        except Exception as e:
            mensaje = str(e)
            _log("ERROR de conexión: %s" % mensaje)
            bajo = mensaje.lower()
            if any(pista in bajo for pista in ("tenant", "not found", "no such", "project")):
                _log("PISTA: el proyecto parece estar PAUSADO. Entre al panel de Supabase "
                     "y pulse 'Resume project'.")
            return 2

    try:
        con.autocommit = True
        with con.cursor() as cur:
            for i in range(1, CONSULTAS + 1):
                cur.execute("SELECT now(), current_user")
                fila = cur.fetchone()
                _log("Consulta %d/%d OK -> %s (usuario %s)"
                     % (i, CONSULTAS, fila[0].strftime("%Y-%m-%d %H:%M:%S %Z"), fila[1]))
                if i < CONSULTAS:
                    time.sleep(PAUSA_ENTRE)
        _log("Proyecto ACTIVO: la actividad quedó registrada en Supabase (Licencias Colliesoftware).")
        return 0
    except Exception as e:
        _log("ERROR al ejecutar las consultas: %s" % e)
        return 2
    finally:
        try:
            con.close()
        except Exception:
            pass


if __name__ == "__main__":
    sys.exit(mantener_activo())
