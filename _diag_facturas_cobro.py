
import json, os, psycopg2
cfg = json.load(open(r"C:\Users\Alberto\Desktop\Programa de control de flotilla automotriz para Win-Mac 210826\config_db.json", encoding="utf-8"))
conn = psycopg2.connect(host=cfg["host"], port=cfg["port"], dbname=cfg["dbname"], user=cfg["user"], password=cfg["password"], sslmode="require", connect_timeout=15)
cur = conn.cursor()
cur.execute("SELECT COUNT(*) FROM facturas_emitidas")
print("facturas_emitidas TOTAL:", cur.fetchone()[0])
cur.execute("SELECT COUNT(*) FROM facturas_emitidas WHERE COALESCE(enlace_pdf_sunat,'')='' AND COALESCE(archivo_ruta,'')=''")
print("sin enlace_sunat Y sin archivo_ruta:", cur.fetchone()[0])
cur.execute("SELECT COUNT(*) FROM facturas_emitidas WHERE COALESCE(estado_sunat,'') ILIKE '%Anulada%'")
print("anuladas:", cur.fetchone()[0])
cur.execute("SELECT COUNT(*) FROM facturas_emitidas WHERE COALESCE(estado_sunat,'') NOT ILIKE '%Anulada%' AND (COALESCE(enlace_pdf_sunat,'')<>'' OR COALESCE(archivo_ruta,'')<>'')")
print("visibles actualmente (no anulada y con enlace o archivo):", cur.fetchone()[0])
print()
print("=== ultimas 20 facturas_emitidas (id, fecha, nro_doc, cliente, estado, enlace, archivo) ===")
cur.execute("SELECT id, fecha, numero_documento, cliente, COALESCE(estado_sunat,''), COALESCE(enlace_pdf_sunat,''), COALESCE(archivo_ruta,''), total FROM facturas_emitidas ORDER BY id DESC LIMIT 20")
for r in cur.fetchall():
    print(" id=%s fecha=%s nro=%r cliente=%r estado=%r enlace=%r archivo=%r total=%s" % r)
cur.close(); conn.close()
