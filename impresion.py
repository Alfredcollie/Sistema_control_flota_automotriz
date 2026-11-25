# -*- coding: utf-8 -*-
"""Impresión directa de PDF en Windows y macOS.

Usa la impresora multifuncional configurada en Configuración General
(config_local.json, campo «impresora»). Si está vacía o falla, imprime en
la impresora predeterminada del sistema.
"""
import json
import os
import subprocess
import sys


def _config():
    try:
        from app_paths import CONFIG_FILE
        with open(str(CONFIG_FILE), "r", encoding="utf-8") as archivo:
            return json.load(archivo) or {}
    except Exception:
        return {}


def impresora_configurada():
    """Nombre de la impresora configurada en el sistema (puede venir vacío)."""
    return str(_config().get("impresora") or "").strip()


def imprimir_pdf(ruta, impresora=None):
    """Envía el PDF a la impresora indicada (o a la configurada).

    Devuelve (ok, detalle): si se envió, «detalle» es la impresora usada;
    si falló, es el mensaje de error.
    """
    ruta = str(ruta or "").strip()
    if not ruta or not os.path.isfile(ruta):
        return False, "No se encontró el archivo PDF."
    if impresora is None:
        impresora = impresora_configurada()
    impresora = str(impresora or "").strip()

    if sys.platform == "win32":
        return _imprimir_windows(ruta, impresora)
    return _imprimir_unix(ruta, impresora)


def _imprimir_windows(ruta, impresora):
    """Windows: «Imprimir en» la impresora configurada y, si falla, la predeterminada."""
    try:
        import ctypes
    except Exception as error:
        return False, str(error)
    if impresora:
        try:
            resultado = ctypes.windll.shell32.ShellExecuteW(
                None, "printto", ruta, '"%s"' % impresora, None, 0)
            if int(resultado) > 32:
                return True, impresora
        except Exception:
            pass
    try:
        os.startfile(ruta, "print")          # noqa: S606 (solo Windows)
        return True, "predeterminada"
    except Exception as error:
        return False, str(error)


def _imprimir_unix(ruta, impresora):
    """macOS/Linux: envía el PDF con lp (CUPS)."""
    orden = ["lp"]
    if impresora:
        orden += ["-d", impresora]
    orden.append(ruta)
    error = ""
    try:
        resultado = subprocess.run(orden, capture_output=True, text=True)
        if resultado.returncode == 0:
            return True, impresora or "predeterminada"
        error = (resultado.stderr or resultado.stdout or "").strip()
    except Exception as fallo:
        error = str(fallo)
    if impresora:
        # La impresora configurada no existe o no responde: se intenta la predeterminada
        try:
            resultado = subprocess.run(["lp", ruta], capture_output=True, text=True)
            if resultado.returncode == 0:
                return True, "predeterminada"
            error = (resultado.stderr or resultado.stdout or error).strip()
        except Exception as fallo:
            error = str(fallo)
    return False, error or "No se pudo imprimir."


def imprimir_con_aviso(ruta, parent=None):
    """Imprime y muestra el resultado al usuario. Devuelve True si se envió."""
    try:
        from tkinter import messagebox
    except Exception:
        messagebox = None
    ok, detalle = imprimir_pdf(ruta)
    if messagebox is not None:
        if ok:
            messagebox.showinfo("Imprimir",
                                "Documento enviado a la impresora: " + detalle + ".",
                                parent=parent)
        else:
            messagebox.showerror("Imprimir", "No se pudo imprimir:\n" + detalle,
                                 parent=parent)
    return ok
