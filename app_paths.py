# -*- coding: utf-8 -*-
"""
APP_PATHS.PY - Rutas de configuración y recursos.
- BASE_DIR / RESOURCES_DIR: carpeta del ejecutable (recursos de SOLO LECTURA).
- DATA_DIR / CONFIG_FILE: carpeta de datos del usuario (ESCRIBIBLE):
    Windows -> %APPDATA%\\ControlFlota
    macOS   -> ~/Library/Application Support/ControlFlota
    Linux   -> ~/.config/ControlFlota
En desarrollo (python) la config sigue en la carpeta del proyecto.
"""
import os
import sys
import shutil
import uuid
from pathlib import Path

APP_NAME = "ControlFlota"


def _base_dir():
    """Carpeta base: la del .exe/.app si está compilado, o la del script."""
    if getattr(sys, "frozen", False):
        return Path(os.path.dirname(sys.executable))
    return Path(__file__).resolve().parent


def _data_dir():
    """Carpeta de datos del usuario (escribible)."""
    if sys.platform == "win32":
        raiz = Path(os.environ.get("APPDATA", str(Path.home() / "AppData" / "Roaming")))
    elif sys.platform == "darwin":
        raiz = Path.home() / "Library" / "Application Support"
    else:
        raiz = Path(os.environ.get("XDG_CONFIG_HOME", str(Path.home() / ".config")))
    carpeta = raiz / APP_NAME
    try:
        carpeta.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    return carpeta


BASE_DIR = _base_dir()
DATA_DIR = _data_dir()

# Archivo de configuración: en la carpeta de datos si está compilado
# (para poder escribir sin permisos de administrador), o en el proyecto en desarrollo.
CONFIG_FILE = (DATA_DIR if getattr(sys, "frozen", False) else BASE_DIR) / "config_local.json"

# Carpeta para recursos (logos, etc.) — junto al ejecutable
RESOURCES_DIR = BASE_DIR

# Identificador único y persistente de este equipo (para saber qué máquina
# realizó el enlace con Rclone). Vive en DATA_DIR, fuera de config_local.json.
DEVICE_ID_FILE = DATA_DIR / "device_id.txt"


def obtener_device_id():
    """Devuelve (y crea si hace falta) un UUID único y persistente por equipo."""
    try:
        if DEVICE_ID_FILE.exists():
            valor = DEVICE_ID_FILE.read_text(encoding="utf-8").strip()
            if valor:
                return valor
    except Exception:
        pass
    nuevo = str(uuid.uuid4())
    try:
        DEVICE_ID_FILE.write_text(nuevo, encoding="utf-8")
    except Exception:
        pass
    return nuevo


# =========================================================
# RESOLUCIÓN PORTABLE DE RUTAS DE ARCHIVOS
# =========================================================
# Subcarpetas que usa el programa para guardar documentos. Sirven para reconocer
# una ruta guardada (aunque venga de otro equipo o de otro sistema operativo) y
# volver a encontrarla en la carpeta local de ESTE equipo.
SUBCARPETAS_CONOCIDAS = (
    "facturas_recibidas", "facturas_emitidas", "comprobantes_egresos", "comprobantes_ingresos",
    "comprobantes_impuestos", "ordenes_generadas", "ordenes_anuladas", "ordenes_compra_recibidas",
    "cobranzas_generadas", "Inspecciones", "Inspecciones_PDF", "archivos_flota",
    "Cotizaciones", "soportes_pagos_terceros", "cartas_eventos", "formularios_proveedores",
    "Pautas_Eventos", "tarjetas_propiedad", "expedientes_choferes",
)


def ruta_drive_configurada():
    """Carpeta local configurada en Configuración del Sistema (ruta_drive)."""
    try:
        import json as _json
        if CONFIG_FILE.exists():
            with open(str(CONFIG_FILE), "r", encoding="utf-8") as f:
                datos = _json.load(f)
            if isinstance(datos, dict):
                return str(datos.get("ruta_drive", "") or "").strip()
    except Exception:
        pass
    return ""


def _normalizar_nombre(nombre):
    """Normaliza un nombre de archivo para compararlo (acentos NFC/NFD y mayúsculas)."""
    try:
        import unicodedata
        return unicodedata.normalize("NFC", str(nombre)).casefold()
    except Exception:
        return str(nombre).casefold()


def partir_ruta_guardada(ruta):
    """(subcarpeta, nombre_archivo) de una ruta guardada, sea Windows o macOS."""
    texto = str(ruta or "").replace("\\", "/").strip()
    if not texto:
        return "", ""
    partes = [p for p in texto.split("/") if p not in ("", ".")]
    if not partes:
        return "", ""
    nombre = partes[-1]
    subcarpeta = ""
    for i in range(len(partes) - 1, -1, -1):
        if partes[i] in SUBCARPETAS_CONOCIDAS:
            subcarpeta = "/".join(partes[i:-1])
            break
    return subcarpeta, nombre


def resolver_ruta_archivo(ruta_guardada, ruta_base=""):
    """Devuelve la ruta REAL del archivo en este equipo, o "" si no se encuentra.

    Sirve cuando la ruta guardada en la base de datos fue creada en otro equipo
    o en otro sistema operativo (por ejemplo, una ruta de macOS en un Windows):
      1) si la ruta existe tal cual, se devuelve;
      2) si es relativa, se resuelve contra la carpeta local (ruta_drive);
      3) se reconstruye subcarpeta + nombre dentro de la carpeta local;
      4) se busca el nombre en las subcarpetas conocidas de la carpeta local.
    """
    if not ruta_guardada:
        return ""
    base = os.path.expanduser(str(ruta_base or "").strip()) or os.path.expanduser(ruta_drive_configurada())
    candidatas = []

    texto = str(ruta_guardada).replace("\\", os.sep).strip()
    if os.path.isabs(texto):
        candidatas.append(texto)
    elif base:
        candidatas.append(os.path.join(base, texto))

    subcarpeta, nombre = partir_ruta_guardada(ruta_guardada)
    if base and nombre:
        if subcarpeta:
            candidatas.append(os.path.join(base, *subcarpeta.split("/"), nombre))
        candidatas.append(os.path.join(base, nombre))
        for conocida in SUBCARPETAS_CONOCIDAS:
            candidatas.append(os.path.join(base, conocida, nombre))
    if nombre:
        candidatas.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), nombre))

    for candidata in candidatas:
        try:
            if candidata and os.path.exists(candidata):
                return os.path.normpath(candidata)
        except Exception:
            continue

    # Último recurso: comparar nombres normalizados (macOS guarda los nombres
    # descompuestos -NFD- y Windows compuestos -NFC-; además puede variar el
    # uso de mayúsculas/minúsculas entre equipos).
    objetivo = _normalizar_nombre(nombre) if nombre else ""
    if objetivo:
        carpetas = []
        if base:
            carpetas.append(base)
            if subcarpeta:
                carpetas.append(os.path.join(base, *subcarpeta.split("/")))
            carpetas.extend(os.path.join(base, conocida) for conocida in SUBCARPETAS_CONOCIDAS)
        for carpeta in carpetas:
            try:
                if not os.path.isdir(carpeta):
                    continue
                for entrada in os.listdir(carpeta):
                    if _normalizar_nombre(entrada) == objetivo:
                        ruta = os.path.join(carpeta, entrada)
                        if os.path.isfile(ruta):
                            return os.path.normpath(ruta)
            except Exception:
                continue
    return ""


def ruta_para_guardar(ruta, ruta_base=""):
    """Devuelve la ruta que se debe GUARDAR en la base de datos.

    Si el archivo está dentro de la carpeta local sincronizada (ruta_drive) se
    guarda RELATIVA (por ejemplo 'facturas_recibidas/archivo.jpg'): así la misma
    fila sirve en Windows y en Mac y no queda atada a ningún equipo.
    Si está fuera de esa carpeta (expedientes de choferes, tarjetas de propiedad,
    inspecciones, etc.) se guarda la ruta absoluta como hasta ahora.
    """
    if not ruta:
        return ruta
    try:
        base = os.path.expanduser(str(ruta_base or "").strip()) or os.path.expanduser(ruta_drive_configurada())
        if not base:
            return ruta
        ruta_abs = os.path.abspath(os.path.expanduser(str(ruta)))
        base_abs = os.path.abspath(base)
        if os.path.commonpath([ruta_abs, base_abs]) != base_abs:
            return ruta  # está fuera de la carpeta local: se guarda completa
        relativa = os.path.relpath(ruta_abs, base_abs)
        if relativa.startswith(".."):
            return ruta
        return relativa.replace(os.sep, "/")
    except Exception:
        return ruta


def eliminar_archivo(ruta_guardada, ruta_base=""):
    """Elimina el archivo resolviendo la ruta (aunque se guardó en otro equipo/SO).

    Devuelve True si se eliminó y False si no se encontró.
    """
    ruta = resolver_ruta_archivo(ruta_guardada, ruta_base)
    if not ruta:
        return False
    try:
        os.remove(ruta)
        return True
    except Exception:
        return False


def _migrar_config():
    """Copia config_local.json desde el ejecutable (config por defecto del
    instalador) a la carpeta de datos del usuario, solo la primera vez."""
    if not getattr(sys, "frozen", False):
        return
    try:
        if not CONFIG_FILE.exists():
            origen = BASE_DIR / "config_local.json"
            if origen.exists():
                shutil.copy2(str(origen), str(CONFIG_FILE))
    except Exception:
        pass


_migrar_config()
