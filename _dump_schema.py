# -*- coding: utf-8 -*-
import json, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from conexion import conectar_db, liberar_conexion

conn = conectar_db(silencioso=True)
if not conn:
    print("NO_CONN"); sys.exit(1)
cur = conn.cursor()
cur.execute("""
    SELECT table_name FROM information_schema.tables
    WHERE table_schema='public' AND table_type='BASE TABLE'
    ORDER BY table_name
""")
tablas = [r[0] for r in cur.fetchall()]
out = {}
for t in tablas:
    cur.execute("""SELECT column_name, data_type FROM information_schema.columns
                   WHERE table_schema='public' AND table_name=%s ORDER BY ordinal_position""", (t,))
    cols = [{"c": r[0], "t": r[1]} for r in cur.fetchall()]
    try:
        cur.execute(f'SELECT count(*) FROM public."{t}"')
        n = cur.fetchone()[0]
    except Exception as e:
        conn.rollback(); n = -1
    out[t] = {"n": n, "cols": cols}
liberar_conexion(conn)
with open("_schema_dump.json", "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=1)
print("TABLAS:", len(out))
