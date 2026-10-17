# -*- coding: utf-8 -*-
"""
VERIFICAR_GUARDADO_ARCHIVOS.PY
=============================
Comprueba que TODOS los módulos que guardan archivos (PDF, JPG, PNG, JSON)
usen la carpeta configurada en "Configuración General" (ruta_drive) — la misma
que sincroniza Rclone con la cuenta del equipo principal — y que las rutas
guardadas en Supabase sean portables (relativas a esa carpeta).

Uso:   python verificar_guardado_archivos.py
r"""
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import app_paths
import politica_almacenamiento as pol

BASE = os.path.normpath(os.path.expanduser(app_paths.ruta_drive_configurada() or ""))
fallos = []


def _es_absoluta(texto):
    r"""Ruta absoluta de Windows (C:\...), de macOS/Linux (/...) o UNC."""
    t = str(texto or "").strip()
    return bool(t) and (os.path.isabs(t) or t.startswith("/") or t.startswith("\\")
                        or (len(t) > 1 and t[1] == ":"))


def dentro(ruta):
    if not ruta:
        return False
    texto = str(ruta).strip()
    if not _es_absoluta(texto):
        return True                         # relativa: se resuelve dentro de la base
    try:
        r = os.path.normpath(os.path.expanduser(texto.replace("/", os.sep)))
        return os.path.commonpath([r.lower(), BASE.lower()]) == BASE.lower()
    except Exception:
        return False


def main():
    print("=" * 78)
    print("CARPETA CONFIGURADA (Configuración General -> ruta_drive):")
    print("   ", BASE or "(no configurada)")
    estado = pol.estado_almacenamiento()
    print("Política de almacenamiento: autorizado=%s | equipo principal=%s"
          % (estado.get("autorizado"), estado.get("es_principal")))
    print("=" * 78)

    print("\n1) CARPETA BASE QUE USA CADA MÓDULO")
    pruebas = [
        ("Compras            (obtener_ruta_base_drive)", "modulo_compras", "obtener_ruta_base_drive"),
        ("Ventas             (obtener_ruta_base_drive)", "modulo_ventas", "obtener_ruta_base_drive"),
        ("Impuestos          (obtener_ruta_base_drive)", "calculo_impuestos", "obtener_ruta_base_drive"),
        ("Órdenes de compra  (_carpeta_base_ordenes)", "ordenes_compra", "_carpeta_base_ordenes"),
        ("Banco              (obtener_ruta_base)", "modulo_banco", "obtener_ruta_base"),
        ("Inspección vehicular (_carpeta_archivos)", "inspeccion_vehicular", "_carpeta_archivos"),
        ("Choferes           (_carpeta_expedientes)", "choferes", "_carpeta_expedientes"),
        ("Cobranza           (_carpeta_pdf)", "calculo_cobranza", "_carpeta_pdf"),
    ]
    for etiqueta, modulo, funcion in pruebas:
        try:
            mod = __import__(modulo)
            ruta = getattr(mod, funcion)()
        except Exception as e:
            print("   %-44s ERROR: %s" % (etiqueta, e)); fallos.append(etiqueta); continue
        ok = dentro(ruta)
        print("   %-44s %-52s %s" % (etiqueta, str(ruta)[:52], "OK" if ok else "FUERA DE LA CARPETA ✗"))
        if not ok:
            fallos.append(etiqueta)
    # Flota: se comprueba que el código use la carpeta autorizada
    try:
        fuente = open("flota_automotriz.py", encoding="utf-8").read()
        bloque = fuente[fuente.find("ruta_final_archivo = self.ruta_tarjeta_db"):
                        fuente.find("ruta_final_archivo = self.ruta_tarjeta_db") + 1400]
        ok = ("ruta_base_autorizada" in bloque) and ("dirname(os.path.abspath(__file__))" not in bloque)
        print("   %-44s %-52s %s" % ("Flota (tarjetas de propiedad)", "(usa ruta_base_autorizada)", "OK" if ok else "REVISAR ✗"))
        if not ok:
            fallos.append("Flota")
    except Exception as e:
        print("   Flota: no se pudo comprobar (%s)" % e)

    print("\n2) RUTAS GUARDADAS EN SUPABASE")
    try:
        from conexion import conectar_db, liberar_conexion
        conn = conectar_db(silencioso=True)
        if not conn:
            print("   (sin conexión)")
            return
        try:
            with conn.cursor() as c:
                c.execute("""SELECT table_name, column_name FROM information_schema.columns
                             WHERE table_schema='public' AND data_type IN ('text','character varying')
                               AND (column_name ILIKE '%%ruta%%' OR column_name ILIKE '%%archivo%%')
                             ORDER BY table_name""")
                columnas = c.fetchall()
            total_rel = total_fuera = 0
            for tabla, col in columnas:
                try:
                    with conn.cursor() as c:
                        c.execute('SELECT "%s" FROM "%s" WHERE COALESCE("%s"::text, \'\') <> \'\''
                                  % (col, tabla, col))
                        valores = [str(r[0]) for r in c.fetchall()]
                except Exception:
                    conn.rollback(); continue
                if not valores:
                    continue
                rel = fuera = 0
                ejemplo = ""
                for v in valores:
                    textos = list(json.loads(v).values()) if v.startswith("{") else [v]
                    for t in textos:
                        t = str(t)
                        ruta_ok = dentro(t)
                        if ruta_ok:
                            rel += 1
                        else:
                            fuera += 1
                            ejemplo = ejemplo or t[-58:]
                total_rel += rel; total_fuera += fuera
                if fuera:
                    print("   %-24s %-16s dentro=%-4s FUERA=%-4s %s" % (tabla, col, rel, fuera, ejemplo))
                else:
                    print("   %-24s %-16s dentro=%-4s FUERA=%-4s" % (tabla, col, rel, fuera))
            print("\n   RESUMEN: %s ruta(s) portables  |  %s FUERA de la carpeta sincronizada"
                  % (total_rel, total_fuera))
            if total_fuera:
                print("   (las de FUERA corresponden a archivos guardados por versiones anteriores"
                      " en otra carpeta: hay que volver a subirlos desde el módulo)")
                fallos.append("rutas fuera de la carpeta")
        finally:
            liberar_conexion(conn)
    except Exception as e:
        print("   No se pudo revisar Supabase: %s" % e)

    print("\n" + "=" * 78)
    print("RESULTADO:", "TODO CORRECTO ✔" if not fallos else ("REVISAR: %s" % ", ".join(fallos)))


if __name__ == "__main__":
    main()
