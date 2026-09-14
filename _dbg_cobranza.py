# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from conexion import conectar_db, liberar_conexion
conn = conectar_db(silencioso=True)
cur = conn.cursor()
sql = """SELECT id, COALESCE(cliente_nombre,''), COALESCE(cliente_ruc,''),
                COALESCE(anio,0), COALESCE(mes,0), COALESCE(quincena,0),
                COALESCE(plan_cobro,''), COALESCE(total,0), COALESCE(facturado,FALSE),
                COALESCE(factura_referencia,''), COALESCE(fecha_registro,''),
                COALESCE(notas,''), COALESCE(id_cliente,0)
         FROM cobranza_quincenas"""
try:
    cur.execute(sql)
    filas = cur.fetchall()
    print("FILAS:", len(filas), "COLS:", len(filas[0]) if filas else 0)
    for f in filas:
        print(repr(f))
except Exception as e:
    print("ERROR:", type(e).__name__, e)
liberar_conexion(conn)
