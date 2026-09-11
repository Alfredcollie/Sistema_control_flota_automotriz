# -*- coding: utf-8 -*-
"""
CONFIG_NUBE.PY - Configuración compartida guardada en Supabase (nube).

Los datos bancarios y la cuenta del App Grifo viven en la tabla
`config_general` de Supabase (NO en el archivo local), para que sean
consistentes entre todos los equipos (Windows / macOS).
"""
import json
from conexion import conectar_db, liberar_conexion

CLAVE_BANCOS = "cuentas_bancarias"
CLAVE_GRIFO = "cuenta_grifo_pagos"
CLAVE_RCLONE = "rclone_sync"


def _asegurar_tabla(cursor):
    cursor.execute(
        "CREATE TABLE IF NOT EXISTS config_general "
        "(clave VARCHAR(255) PRIMARY KEY, valor TEXT)"
    )


def _cargar_clave(clave, default=""):
    conn = conectar_db(silencioso=True)
    if not conn:
        return default
    try:
        with conn.cursor() as cursor:
            _asegurar_tabla(cursor)
            cursor.execute("SELECT valor FROM config_general WHERE clave = %s", (clave,))
            fila = cursor.fetchone()
            if fila and fila[0] is not None:
                return fila[0]
    except Exception:
        pass
    finally:
        liberar_conexion(conn)
    return default


def _guardar_clave(clave, valor):
    conn = conectar_db(silencioso=True)
    if not conn:
        return False
    try:
        with conn.cursor() as cursor:
            _asegurar_tabla(cursor)
            cursor.execute(
                "INSERT INTO config_general (clave, valor) VALUES (%s, %s) "
                "ON CONFLICT (clave) DO UPDATE SET valor = EXCLUDED.valor",
                (clave, valor),
            )
        conn.commit()
        return True
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        return False
    finally:
        liberar_conexion(conn)


def clave_existe_en_nube(clave):
    """True si la clave ya existe en la tabla config_general de Supabase.

    Sirve para no pisar el valor local con uno vacío durante la migración:
    solo se toma la nube como fuente de verdad cuando ya existe la clave.
    """
    conn = conectar_db(silencioso=True)
    if not conn:
        return False
    try:
        with conn.cursor() as cursor:
            _asegurar_tabla(cursor)
            cursor.execute("SELECT 1 FROM config_general WHERE clave = %s", (clave,))
            return cursor.fetchone() is not None
    except Exception:
        return False
    finally:
        liberar_conexion(conn)


def cargar_bancos():
    """Lista de cuentas bancarias de la empresa (desde Supabase)."""
    raw = _cargar_clave(CLAVE_BANCOS, "[]")
    try:
        datos = json.loads(raw) if raw else []
        if isinstance(datos, list):
            return datos
    except Exception:
        pass
    return []


def guardar_bancos(lista_bancos):
    """Guarda la lista de cuentas bancarias en Supabase (como JSON)."""
    return _guardar_clave(CLAVE_BANCOS, json.dumps(lista_bancos or [], ensure_ascii=False))


def cargar_cuenta_grifo():
    """Cuenta asignada para pagos del App Grifo (desde Supabase)."""
    return (_cargar_clave(CLAVE_GRIFO, "") or "").strip()


def guardar_cuenta_grifo(cuenta):
    """Guarda la cuenta del App Grifo en Supabase."""
    return _guardar_clave(CLAVE_GRIFO, cuenta or "")


def cargar_rclone_sync():
    """Devuelve el dict de sincronización en la nube registrado en Supabase,
    o None si todavía ningún equipo lo registró.

    Estructura esperada:
        {"rclone_remote": "gdrive:", "rclone_ruta_nube": "BlackCube",
         "linked_by_device": "<uuid>", "linked_at": "..."}
    """
    raw = _cargar_clave(CLAVE_RCLONE, "")
    if not raw:
        return None
    try:
        datos = json.loads(raw)
        if isinstance(datos, dict):
            return datos
    except Exception:
        pass
    return None


def registrar_rclone_sync(datos):
    """Registra la configuración de sincronización SOLO si aún no existe
    (regla "first-write-wins"): el primer equipo que la registra queda como
    fuente de verdad para todos los demás y nadie la pisa.

    Devuelve el dict que finalmente quedó guardado en la nube (el nuevo si
    estaba vacío, o el existente si otro equipo ya lo había registrado).
    """
    conn = conectar_db(silencioso=True)
    if not conn:
        return None
    try:
        with conn.cursor() as cursor:
            _asegurar_tabla(cursor)
            cursor.execute(
                "INSERT INTO config_general (clave, valor) VALUES (%s, %s) "
                "ON CONFLICT (clave) DO NOTHING",
                (CLAVE_RCLONE, json.dumps(datos or {}, ensure_ascii=False)),
            )
        conn.commit()
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
    finally:
        liberar_conexion(conn)
    return cargar_rclone_sync()


def borrar_rclone_sync():
    """Borra el registro de sincronización en la nube (solo lo hace el equipo
    propietario desde la interfaz; aquí no se valida quién lo pide)."""
    conn = conectar_db(silencioso=True)
    if not conn:
        return False
    try:
        with conn.cursor() as cursor:
            _asegurar_tabla(cursor)
            cursor.execute("DELETE FROM config_general WHERE clave = %s", (CLAVE_RCLONE,))
        conn.commit()
        return True
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        return False
    finally:
        liberar_conexion(conn)
