# -*- coding: utf-8 -*-
"""
LIMPIAR_DATOS_NOMINA.PY - Deja el modulo de Nomina como recien instalado
========================================================================
Vacia las tablas nom_* del modulo de Nomina (marcaciones, asistencia, planilla,
empleados de nomina, turnos, horarios, vinculos del reloj, incidencias, etc.).

  * NO toca la tabla 'choferes' ni ningun otro modulo del sistema.
  * Los turnos tipicos y los parametros se vuelven a crear solos, con sus valores
    por defecto, la proxima vez que se abra el modulo.

Uso:
    python limpiar_datos_nomina.py

Pide confirmacion escribiendo la palabra BORRAR. No se puede deshacer.
"""
import sys
import os
import warnings

warnings.filterwarnings("ignore")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import nomina_core as nc

TABLAS = ["nom_parametros", "nom_turnos", "nom_horarios", "nom_horario_detalle", "nom_empleado",
          "nom_asignaciones", "nom_calendario", "nom_mapeo_reloj", "nom_marcaciones", "nom_asistencia",
          "nom_incidencias", "nom_horas_extra", "nom_importaciones", "nom_planilla",
          "nom_planilla_detalle"]

ETIQUETAS = {
    "nom_marcaciones": "marcaciones del reloj", "nom_asistencia": "dias de asistencia calculada",
    "nom_empleado": "fichas de nomina", "nom_mapeo_reloj": "vinculos reloj-DNI",
    "nom_importaciones": "importaciones registradas", "nom_horas_extra": "horas extra",
    "nom_incidencias": "incidencias", "nom_turnos": "turnos", "nom_horarios": "horarios",
    "nom_planilla": "planillas", "nom_planilla_detalle": "detalle de planilla",
    "nom_calendario": "feriados", "nom_asignaciones": "asignaciones", "nom_parametros": "parametros",
}


def main():
    conn = nc.conectar_db(silencioso=True)
    if not conn:
        print("Sin conexion a la base de datos.")
        return 1
    nc.liberar_conexion(conn)

    print("=" * 74)
    print(" LIMPIAR LOS DATOS DEL MODULO DE NOMINA")
    print("=" * 74)
    print("\nEsto es lo que hay ahora:")
    total = 0
    for tabla in TABLAS:
        fila = nc._consultar('SELECT COUNT(*) FROM "%s"' % tabla, uno=True)
        cantidad = nc.a_int(fila[0]) if fila else 0
        total += cantidad
        if cantidad:
            print("   %-24s %6s  (%s)" % (tabla, format(cantidad, ","),
                                          ETIQUETAS.get(tabla, "datos")))
    print("\n   TOTAL: %s filas" % format(total, ","))
    print("\nNo se toca la tabla 'choferes' ni los demas modulos del sistema.")
    print("Los turnos tipicos y los parametros se vuelven a crear solos al abrir el modulo.\n")

    try:
        respuesta = input("Escriba BORRAR (en mayusculas) para confirmar: ")
    except EOFError:
        respuesta = ""
    if respuesta.strip().upper() != "BORRAR":
        print("\nCancelado. No se borro nada.")
        return 1

    conn = nc.conectar_db(silencioso=True)
    if not conn:
        print("Sin conexion a la base de datos.")
        return 1
    try:
        with conn.cursor() as cursor:
            for tabla in TABLAS:
                cursor.execute('TRUNCATE TABLE "%s" RESTART IDENTITY CASCADE' % tabla)
        conn.commit()
    except Exception as e:
        conn.rollback()
        print("Error al borrar: %s" % e)
        return 1
    finally:
        nc.liberar_conexion(conn)

    ok, error = nc.inicializar_esquema_nomina(forzar=True)
    print("\nListo: se vaciaron %d tablas y los contadores de ID volvieron a 1." % len(TABLAS))
    if ok:
        turnos = nc._consultar("SELECT COUNT(*) FROM nom_turnos", uno=True)
        parametros = nc._consultar("SELECT COUNT(*) FROM nom_parametros", uno=True)
        print("El modulo quedo listo con %s turnos tipicos y %s parametros por defecto."
              % (turnos[0] if turnos else 0, parametros[0] if parametros else 0))
    else:
        print("Aviso: no se pudo recrear el esquema: %s" % error)
    print("\nAbra el modulo de Nomina y use el ASISTENTE PASO A PASO para empezar de cero.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
