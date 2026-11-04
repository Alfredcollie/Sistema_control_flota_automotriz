# -*- coding: utf-8 -*-
"""
DATOS PARA LOS SECRETOS DE GITHUB  (keepalive de licencias)
===========================================================
El workflow «Mantener activo Supabase (Licencias Colliesoftware)» falla con
«faltan credenciales» porque el repositorio todavía no tiene los secretos.

Este programa muestra los datos del equipo y los copia al portapapeles para
que los pegues en GitHub:

    https://github.com/Alfredcollie/Sistema_control_flota_automotriz/settings/secrets/actions

Secretos que hay que crear (New repository secret), uno por uno:
    LICENCIAS_DB_HOST       ·  LICENCIAS_DB_PORT   ·  LICENCIAS_DB_NAME
    LICENCIAS_DB_USER       ·  LICENCIAS_DB_PASSWORD

La contraseña NO se muestra en pantalla: solo se copia al portapapeles.
"""

import json
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(BASE, "config_licencia.json")
URL_SECRETOS = ("https://github.com/Alfredcollie/Sistema_control_flota_automotriz"
                "/settings/secrets/actions")


def _datos():
    datos = {"host": "", "port": "", "dbname": "", "user": "", "password": ""}
    # 1) archivo de configuración del programa
    try:
        if os.path.isfile(CONFIG):
            with open(CONFIG, "r", encoding="utf-8-sig") as archivo:
                cfg = json.load(archivo)
            for clave in ("host", "port", "dbname", "user"):
                datos[clave] = str(cfg.get(clave) or "")
    except Exception as e:
        print(" (no se pudo leer config_licencia.json: %s)" % e)
    # 2) contraseña desde el llavero del sistema
    try:
        import keyring
        for servicio in ("ControlFlotaLicencias", "Sistema Control Flota"):
            for clave in ("LICENCIAS_DB_PASSWORD", "SUPABASE_DB_PASSWORD"):
                valor = keyring.get_password(servicio, clave)
                if valor:
                    datos["password"] = valor
                    break
            if datos["password"]:
                break
    except Exception as e:
        print(" (no se pudo leer el llavero: %s)" % e)
    return datos


def _copiar(texto):
    """Copia al portapapeles (Windows/macOS) sin mostrar el valor."""
    try:
        import tkinter
        raiz = tkinter.Tk()
        raiz.withdraw()
        raiz.clipboard_clear()
        raiz.clipboard_append(texto)
        raiz.update()
        raiz.destroy()
        return True
    except Exception as e:
        print("   (no se pudo copiar: %s)" % e)
        return False


def main():
    print("=" * 70)
    print(" DATOS PARA LOS SECRETOS DE GITHUB (keepalive de licencias)")
    print("=" * 70)
    datos = _datos()
    print("")
    print(" Datos del equipo:")
    print("   LICENCIAS_DB_HOST      = " + (datos["host"] or "(no encontrado)"))
    print("   LICENCIAS_DB_PORT      = " + (datos["port"] or "(no encontrado)"))
    print("   LICENCIAS_DB_NAME      = " + (datos["dbname"] or "(no encontrado)"))
    print("   LICENCIAS_DB_USER      = " + (datos["user"] or "(no encontrado)"))
    print("   LICENCIAS_DB_PASSWORD  = " +
          ("(guardada en el llavero, %d caracteres)" % len(datos["password"])
           if datos["password"] else "(no encontrada)"))
    print("")
    print(" Página para crear los secretos:")
    print("   " + URL_SECRETOS)
    print("")
    print(" Elegí un número: se copia al portapapeles y lo pegás en GitHub")
    print("   [1] HOST        [2] PORT        [3] NAME")
    print("   [4] USER        [5] PASSWORD    [0] Salir")
    print("")

    opciones = {"1": ("host", "LICENCIAS_DB_HOST"),
                "2": ("port", "LICENCIAS_DB_PORT"),
                "3": ("dbname", "LICENCIAS_DB_NAME"),
                "4": ("user", "LICENCIAS_DB_USER"),
                "5": ("password", "LICENCIAS_DB_PASSWORD")}
    while True:
        try:
            opcion = input(" Opción: ").strip()
        except Exception:
            return 1
        if opcion in ("0", ""):
            print(" Listo.")
            return 0
        if opcion not in opciones:
            print("   (opción no válida)")
            continue
        clave, nombre_secreto = opciones[opcion]
        valor = datos.get(clave, "")
        if not valor:
            print("   ⚠️  No se encontró ese dato en este equipo.")
            continue
        if _copiar(valor):
            if clave == "password":
                print("   ✔ Contraseña copiada al portapapeles (%d caracteres)." % len(valor))
            else:
                print("   ✔ %s copiado: %s" % (nombre_secreto, valor))
            print("     Pegalo en GitHub → New repository secret → %s" % nombre_secreto)
        print("")


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nCancelado.")
        sys.exit(1)
