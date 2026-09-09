# -*- coding: utf-8 -*-
import json, os, csv, psycopg2
from datetime import datetime
for k in ("SUPABASE_DB_HOST","SUPABASE_DB_PASSWORD","SUPABASE_DB_USER","SUPABASE_DB_PORT","SUPABASE_DB_NAME"):
    os.environ.pop(k, None)
cfg = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "config_db.json"), encoding="utf-8"))
conn = psycopg2.connect(host=cfg["host"], port=cfg["port"], dbname=cfg["dbname"],
    user=cfg["user"], password=cfg["password"], sslmode="require", connect_timeout=10)
cur = conn.cursor()
stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "backup_pagos_comprobantes_%s.csv" % stamp)
cur.execute("SELECT id, codigo_cotizacion, categoria_suministro, monto_pagado, archivo_ruta, proveedor_nombre, fecha_pago, id_factura, cuenta_origen FROM pagos_comprobantes ORDER BY id")
rows = cur.fetchall()
cols = ["id","codigo_cotizacion","categoria_suministro","monto_pagado","archivo_ruta","proveedor_nombre","fecha_pago","id_factura","cuenta_origen"]
with open(out, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f); w.writerow(cols); w.writerows(rows)
print("BACKUP:", out)
print("rows backed up:", len(rows))
cur.close(); conn.close()
