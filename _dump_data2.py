# -*- coding: utf-8 -*-
import json, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from conexion import conectar_db, liberar_conexion
conn = conectar_db(silencioso=True); cur = conn.cursor()
out={}
def q(l,s):
    try:
        cur.execute(s); out[l]=[list(r) for r in cur.fetchall()]
    except Exception as e:
        conn.rollback(); out[l]="ERR "+str(e)[:100]
q("muestra_fr","SELECT fecha, numero_documento, tipo_documento, proveedor, subtotal, impuesto, det_monto, total, cantidad_combustible, kilometraje, evento_asociado, descripcion FROM facturas_recibidas ORDER BY fecha LIMIT 6")
q("combustible_dist","SELECT cantidad_combustible, count(*) FROM facturas_recibidas GROUP BY 1 ORDER BY 2 DESC LIMIT 15")
q("km_dist","SELECT kilometraje, count(*) FROM facturas_recibidas GROUP BY 1 ORDER BY 2 DESC LIMIT 15")
q("cta_uniq","SELECT DISTINCT cuenta_origen FROM pagos_comprobantes")
q("pagos_sample","SELECT fecha_pago, proveedor_nombre, monto_pagado, cuenta_origen, categoria_suministro, id_factura FROM pagos_comprobantes ORDER BY fecha_pago LIMIT 5")
q("conc_all","SELECT banco,cuenta,fecha,descripcion,monto,tipo,origen,estado FROM conciliacion_bancaria")
q("tras","SELECT count(*) FROM transferencias_bancarias")
q("cob_unid","SELECT * FROM cobranza_quincena_unidades LIMIT 3")
q("cob_dias","SELECT categoria, count(*) FROM cobranza_detalle_dias GROUP BY 1")
q("cob_viajes","SELECT * FROM cobranza_quincena_viajes LIMIT 3")
q("tipos_entrada","SELECT * FROM tipos_entrada_cronograma")
q("insp","SELECT count(*) FROM inspecciones")
q("choferes","SELECT count(*) FROM choferes")
q("reg_imp","SELECT count(*) FROM registro_impuestos")
q("users","SELECT usuario, rol, activo FROM usuarios")
q("bitacora_sample","SELECT fecha,hora,usuario,modulo,accion FROM bitacora_auditoria ORDER BY id DESC LIMIT 5")
liberar_conexion(conn)
open("_data2.json","w",encoding="utf-8").write(json.dumps(out,ensure_ascii=False,indent=1,default=str))
print("ok")
