# -*- coding: utf-8 -*-
"""ARNÉS DE PRUEBA (dry-run) de la conciliación inteligente.

NO MODIFICA NADA: ejecuta el SQL real y siempre hace ROLLBACK.
Uso:  python prueba_conciliacion.py [banco]
"""
import os, sys, time, datetime

FL = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, FL)
import conexion
import conciliacion_inteligente as ci

BANCO = sys.argv[1] if len(sys.argv) > 1 else "Interbank"


def conectar():
    for _ in range(5):
        con = conexion.conectar_db(silencioso=True)
        if con:
            return con
        time.sleep(4)
    return None


def main():
    con = conectar()
    if not con:
        print("sin conexión")
        return 1
    con.autocommit = False
    cur = con.cursor()
    cur.execute("""SELECT id, fecha, descripcion, monto, saldo
                   FROM conciliacion_movimientos_pdf
                   WHERE UPPER(COALESCE(banco,'')) LIKE UPPER(%s) ORDER BY id""",
                ("%" + BANCO + "%",))
    movs = [{"id": i + 1, "banco": BANCO, "fecha": f[1], "descripcion": f[2],
             "monto": f[3], "pagina": 0} for i, f in enumerate(cur.fetchall())]
    print("movimientos del estado de cuenta:", len(movs))

    asientos = []
    cur.execute("""SELECT id, fecha, descripcion, monto FROM conciliacion_bancaria
                   WHERE UPPER(COALESCE(banco,'')) LIKE UPPER(%s)""", ("%" + BANCO + "%",))
    for f in cur.fetchall():
        asientos.append({"id": f[0], "id_asiento": f[0], "tabla": "conciliacion_bancaria",
                         "fecha": f[1], "descripcion": f[2], "monto": f[3], "banco": BANCO})
    cur.execute("""SELECT p.id, p.fecha_pago, p.monto_pagado, COALESCE(p.proveedor_nombre,'')
                   FROM pagos_comprobantes p WHERE COALESCE(p.monto_pagado,0) > 0""")
    for f in cur.fetchall():
        asientos.append({"id": "pc:%s" % f[0], "id_asiento": None, "tabla": "pagos_comprobantes",
                         "fecha": f[1], "descripcion": f[3], "monto": -abs(float(f[2] or 0)), "banco": BANCO})
    cur.execute("""SELECT p.id, p.fecha_pago, p.monto_pagado, COALESCE(p.cliente_nombre,'')
                   FROM pagos_clientes p WHERE COALESCE(p.monto_pagado,0) > 0""")
    for f in cur.fetchall():
        asientos.append({"id": "pc2:%s" % f[0], "id_asiento": None, "tabla": "pagos_clientes",
                         "fecha": f[1], "descripcion": f[3], "monto": abs(float(f[2] or 0)), "banco": BANCO})
    if "INTERBANK" in BANCO.upper():
        cur.execute("""SELECT id, fecha, proveedor, total FROM facturas_recibidas
                       WHERE categoria = 'Combustible y Peajes' AND COALESCE(total,0) > 0""")
        for f in cur.fetchall():
            asientos.append({"id": "fr:%s" % f[0], "id_asiento": None, "tabla": "facturas_recibidas",
                             "fecha": f[1], "descripcion": f[2], "monto": -abs(float(f[3] or 0)), "banco": BANCO})
    print("asientos contables disponibles:", len(asientos))

    propuestas = ci.proponer(movs, asientos, ventana_dias=5)
    print("propuestas:", len(propuestas))

    errores, simuladas = [], 0
    for p in propuestas[:80]:
        ident = p.get("id_asiento")
        ident_int = ident if isinstance(ident, int) else None
        try:
            cur.execute("SAVEPOINT prueba")
            if ident_int:
                cur.execute("""UPDATE conciliacion_bancaria
                               SET estado='conciliado', fecha_conciliacion=%s
                               WHERE id=%s""",
                            (datetime.datetime.now().strftime("%d/%m/%Y"), ident_int))
            cur.execute("""INSERT INTO conciliacion_auditoria
                (id_asiento, banco, cuenta, fecha_contable, fecha_banco, dias_desfase,
                 descripcion_contable, descripcion_banco, importe_contable, importe_banco,
                 diferencia, sentido, referencia, pagina_pdf, nivel, estado, confianza,
                 reglas, observacion, accion_sugerida, origen_decision, aprobado_por,
                 creado, activo)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,'PRUEBA',%s,%s,%s,%s,
                        'prueba','',%s,TRUE)""",
                (ident_int, p["banco"], p.get("cuenta", ""), p["fecha_contable"],
                 p["fecha_banco"], p["dias_desfase"], p["descripcion_contable"],
                 p["descripcion_banco"], p["importe_contable"], p["importe_banco"],
                 0, p["sentido"], p.get("referencia", ""), p.get("pagina_pdf"),
                 p["nivel"], p["confianza"], p["reglas"], p.get("observacion", ""),
                 p.get("accion_sugerida", ""), datetime.datetime.now().strftime("%Y-%m-%d %H:%M")))
            cur.execute("ROLLBACK TO SAVEPOINT prueba")
            simuladas += 1
        except Exception as error:
            errores.append("%s -> %s" % (ident, str(error).strip().splitlines()[0][:80]))
            try:
                cur.execute("ROLLBACK TO SAVEPOINT prueba")
            except Exception:
                pass
    con.rollback()

    print()
    print("=== RESULTADO DEL DRY-RUN (nada se guardó) ===")
    print("   propuestas simuladas ........", simuladas)
    print("   errores de SQL/tipos ........", len(errores))
    for e in errores[:10]:
        print("      ", e)
    cur.execute("SELECT COUNT(*) FROM conciliacion_auditoria")
    print("   filas en auditoría (debe seguir igual):", cur.fetchone()[0])
    cur.close()
    try:
        conexion.liberar_conexion(con)
    except Exception:
        con.close()
    return 0 if not errores else 2


if __name__ == "__main__":
    raise SystemExit(main())
