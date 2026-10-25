# -*- coding: utf-8 -*-
"""
DOCUMENTOS_ADJUNTOS.PY - Archivos (PDF/escaneos) de vehículos y choferes
=========================================================================
Guarda en una sola tabla qué archivo corresponde a cada documento:

    tipo        "vehiculo" | "chofer"
    entidad     placa (vehículo) o id (chofer)
    documento   "soat", "revision_tecnica", "extintor", "revision_gas", "brevete"...

La ruta se guarda RELATIVA a la carpeta sincronizada (cuando está dentro),
así la misma fila sirve en Windows y en Mac.
"""
import os
import datetime

from conexion import conectar_db, liberar_conexion

_TABLA_OK = False

SUBCARPETAS = {"vehiculo": "documentos_vehiculos", "chofer": "documentos_choferes"}


def _asegurar(cursor):
    global _TABLA_OK
    if _TABLA_OK:
        return
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS documentos_adjuntos (
            id SERIAL PRIMARY KEY,
            entidad_tipo VARCHAR(20) NOT NULL,
            entidad VARCHAR(120) NOT NULL,
            documento VARCHAR(60) NOT NULL,
            archivo_ruta TEXT DEFAULT '',
            actualizado VARCHAR(30) DEFAULT '',
            UNIQUE (entidad_tipo, entidad, documento)
        )""")
    _TABLA_OK = True


def _conectar():
    return conectar_db(silencioso=True)


def guardar(tipo, entidad, documento, ruta):
    """Guarda (o reemplaza) el archivo de un documento. Devuelve True/False."""
    tipo = str(tipo or "").strip().lower()
    entidad = str(entidad or "").strip()
    documento = str(documento or "").strip().lower()
    if not tipo or not entidad or not documento or not ruta:
        return False
    try:
        from app_paths import ruta_para_guardar
        ruta_bd = ruta_para_guardar(ruta) or str(ruta)
    except Exception:
        ruta_bd = str(ruta)
    conn = _conectar()
    if not conn:
        return False
    try:
        with conn.cursor() as c:
            _asegurar(c)
            c.execute("""
                INSERT INTO documentos_adjuntos
                    (entidad_tipo, entidad, documento, archivo_ruta, actualizado)
                VALUES (%s, %s, %s, %s, %s)
                ON CONFLICT (entidad_tipo, entidad, documento)
                DO UPDATE SET archivo_ruta = EXCLUDED.archivo_ruta,
                              actualizado  = EXCLUDED.actualizado""",
                      (tipo, entidad, documento, ruta_bd,
                       datetime.datetime.now().strftime("%d/%m/%Y %H:%M")))
        conn.commit()
        return True
    except Exception as e:
        print("[documentos_adjuntos] no se pudo guardar:", e)
        try:
            conn.rollback()
        except Exception:
            pass
        return False
    finally:
        liberar_conexion(conn)


def obtener(tipo, entidad, documento):
    """Ruta REAL del archivo en este equipo (o '' si no hay / no se encuentra)."""
    tipo = str(tipo or "").strip().lower()
    entidad = str(entidad or "").strip()
    documento = str(documento or "").strip().lower()
    if not tipo or not entidad or not documento:
        return ""
    conn = _conectar()
    if not conn:
        return ""
    guardada = ""
    try:
        with conn.cursor() as c:
            _asegurar(c)
            c.execute("""SELECT COALESCE(archivo_ruta, '') FROM documentos_adjuntos
                         WHERE entidad_tipo=%s AND entidad=%s AND documento=%s""",
                      (tipo, entidad, documento))
            fila = c.fetchone()
            if fila:
                guardada = str(fila[0] or "").strip()
    except Exception as e:
        print("[documentos_adjuntos] no se pudo leer:", e)
    finally:
        liberar_conexion(conn)
    if not guardada:
        return ""
    try:
        from app_paths import resolver_ruta_archivo
        return resolver_ruta_archivo(guardada) or ""
    except Exception:
        return guardada if os.path.isfile(guardada) else ""


def quitar(tipo, entidad, documento):
    """Borra el registro (no borra el archivo del disco)."""
    conn = _conectar()
    if not conn:
        return False
    try:
        with conn.cursor() as c:
            _asegurar(c)
            c.execute("""DELETE FROM documentos_adjuntos
                         WHERE entidad_tipo=%s AND entidad=%s AND documento=%s""",
                      (str(tipo).lower(), str(entidad), str(documento).lower()))
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


def carpeta_destino(tipo, entidad):
    """Carpeta donde se dejan los archivos de este vehículo/chofer."""
    tipo = str(tipo or "").strip().lower()
    entidad = str(entidad or "").strip()
    try:
        from app_paths import ruta_drive_configurada
        base = ruta_drive_configurada()
    except Exception:
        base = ""
    if not base:
        base = os.path.join(os.path.expanduser("~"), "Archivos_control_flota")
    sub = SUBCARPETAS.get(tipo, "documentos")
    seguro = "".join(ch for ch in entidad if ch.isalnum() or ch in "-_ ").strip().replace(" ", "_")
    carpeta = os.path.join(base, sub, seguro or "sin_identificar")
    try:
        os.makedirs(carpeta, exist_ok=True)
    except Exception:
        pass
    return carpeta


def ruta_destino(tipo, entidad, documento, extension=".pdf"):
    """Ruta nueva (con fecha) para guardar el escaneo de un documento."""
    marca = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre = "%s_%s%s" % (str(documento or "documento").lower(), marca, extension)
    return os.path.join(carpeta_destino(tipo, entidad), nombre)
