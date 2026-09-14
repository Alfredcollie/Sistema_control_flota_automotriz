# -*- coding: utf-8 -*-
import json, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from conexion import conectar_db, liberar_conexion
conn = conectar_db(silencioso=True)
cur = conn.cursor()
out = {}
def q(label, sql):
    try:
        cur.execute(sql); out[label] = [list(r) for r in cur.fetchall()]
    except Exception as e:
        conn.rollback(); out[label] = "ERR: " + str(e)[:120]

q("fr_categoria", "SELECT COALESCE(categoria,'(null)'), count(*), sum(total), sum(coalesce(impuesto,0)) FROM facturas_recibidas GROUP BY 1 ORDER BY 2 DESC")
q("fr_tipo_doc", "SELECT COALESCE(tipo_documento,'?'), count(*) FROM facturas_recibidas GROUP BY 1 ORDER BY 2 DESC")
q("fr_evento", "SELECT COALESCE(NULLIF(TRIM(evento_asociado),''),'(vacio)'), count(*) FROM facturas_recibidas GROUP BY 1 ORDER BY 2 DESC")
q("fr_proveedor", "SELECT proveedor, count(*) FROM facturas_recibidas GROUP BY 1 ORDER BY 2 DESC LIMIT 40")
q("pc_cuenta", "SELECT COALESCE(NULLIF(TRIM(cuenta_origen),''),'(vacio)'), count(*), sum(monto_pagado) FROM pagos_comprobantes GROUP BY 1 ORDER BY 2 DESC")
q("pc_categoria", "SELECT COALESCE(NULLIF(TRIM(categoria_suministro),''),'(vacio)'), count(*), sum(monto_pagado) FROM pagos_comprobantes GROUP BY 1 ORDER BY 2 DESC")
q("flota", "SELECT placa, marca, modelo, categoria, estado, tipo_combustible, kilometraje FROM flota_vehiculos ORDER BY placa")
q("conc", "SELECT * FROM conciliacion_bancaria")
q("config", "SELECT clave, left(valor,200) FROM config_general")
q("osf", "SELECT placa, servicio, costo_total, fecha_emision, estado, proveedor FROM ordenes_servicio_flota ORDER BY fecha_emision")
q("prov", "SELECT nombre, categoria, categoria_2, categoria_3, categoria_4, categoria_5 FROM proveedores")
q("cli", "SELECT id, ruc, nombre_empresa, plan_cobro, limite_credito FROM clientes")
q("cliun", "SELECT * FROM clientes_unidades")
q("fr_fechas", "SELECT min(fecha), max(fecha), count(DISTINCT fecha) FROM facturas_recibidas")
q("pc_fechas", "SELECT min(fecha_pago), max(fecha_pago) FROM pagos_comprobantes")
q("cq", "SELECT cliente_nombre, anio, mes, quincena, total, facturado, fecha_registro FROM cobranza_quincenas")
q("fr_tercero", "SELECT COALESCE(NULLIF(TRIM(pagado_por_tercero),''),'(no)'), es_compra_cruzada, count(*) FROM facturas_recibidas GROUP BY 1,2")
q("fr_km", "SELECT count(*) FROM facturas_recibidas WHERE COALESCE(NULLIF(TRIM(cantidad_combustible),''),'') <> ''")
q("bitacora", "SELECT modulo, count(*) FROM bitacora_auditoria GROUP BY 1 ORDER BY 2 DESC")
liberar_conexion(conn)
with open("_data_dump.json","w",encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=1, default=str)
print("ok")
