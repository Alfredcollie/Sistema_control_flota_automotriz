# -*- coding: utf-8 -*-
import json, os, psycopg2
for k in ("SUPABASE_DB_HOST","SUPABASE_DB_PASSWORD","SUPABASE_DB_USER","SUPABASE_DB_PORT","SUPABASE_DB_NAME"):
    os.environ.pop(k, None)
cfg = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "config_db.json"), encoding="utf-8"))
conn = psycopg2.connect(host=cfg["host"], port=cfg["port"], dbname=cfg["dbname"],
    user=cfg["user"], password=cfg["password"], sslmode="require", connect_timeout=10)
cur = conn.cursor()

for t in ("pagos_comprobantes","pagos_clientes","facturas_recibidas","facturas_emitidas"):
    cur.execute("SELECT COUNT(*) FROM " + t)
    print(t, "count =", cur.fetchone()[0])

cur.execute("SELECT COALESCE(cuenta_origen,'') FROM pagos_comprobantes")
from collections import Counter
c = Counter(r[0] for r in cur.fetchall())
print("\n=== pagos_comprobantes por cuenta_origen ===")
for k,v in sorted(c.items(), key=lambda x:-x[1]):
    print(" %-40s %d" % (repr(k), v))

if True:
    cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name='pagos_clientes' ORDER BY ordinal_position")
    print("\npagos_clientes cols:", [r[0] for r in cur.fetchall()])
    cur.execute("SELECT id, id_factura, monto_pagado, cliente_nombre, fecha_pago, COALESCE(cuenta_destino,''), COALESCE(codigo_cotizacion,'') FROM pagos_clientes ORDER BY id")
    rows = cur.fetchall()
    print("pagos_clientes total", len(rows))
    for r in rows[:40]:
        print(" id=%s id_fac=%s monto=%s cliente=%r fecha=%r cuenta=%r ref=%r" % (r[0],r[1],r[2],r[3],r[4],r[5],r[6]))

cur.close(); conn.close()
