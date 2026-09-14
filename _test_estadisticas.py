# -*- coding: utf-8 -*-
import sys, os, json, io, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import estadisticas_financiera as EF
from config_nube import cargar_bancos

out = io.StringIO()
def p(*a): out.write(" ".join(str(x) for x in a) + "\n")

bancos = cargar_bancos()
def crear_app():
    obj = object.__new__(EF.EstadisticasFinancieraApp)
    obj.bancos_config = bancos
    obj.tasa_renta_anual = 0.0
    obj.tasa_renta_mensual = 1.5
    obj._placa_objetivo = ""
    return obj

def base_filtros(**kw):
    f = {"banco": "Todas las Cuentas", "categoria": "Todas las Categorías",
         "placa": "Todas las Placas / Vehículos", "cliente": "Todos los Clientes",
         "proveedor": "Todos los Proveedores", "modo_fecha": "Mensual/Anual",
         "mes": "Todos los meses", "anio": "2026", "dt_ini": None, "dt_fin": None,
         "periodo_txt": "Año 2026"}
    f.update(kw); return f

casos = [
    ("SIN FILTROS 2026", base_filtros()),
    ("PLACA CHD817 (sin guion)", base_filtros(placa="CHD817")),
    ("PLACA BSK-758", base_filtros(placa="BSK-758 | Ram V 700")),
    ("CATEGORIA Gastos Fijos", base_filtros(categoria="Gastos Fijos")),
    ("CATEGORIA Alquiler (sub)", base_filtros(categoria="Gastos Fijos - Alquiler")),
    ("CATEGORIA Taller Mecanico", base_filtros(categoria="Taller Mecánico")),
    ("CATEGORIA del banco: Planilla", base_filtros(categoria="Gastos Fijos - Planilla / Sueldos")),
    ("BANCO Interbank", base_filtros(banco="Interbank - 297-3007235833")),
    ("BANCO BCP", base_filtros(banco="BCP - 1937218112021")),
    ("CLIENTE BAGUETERIA", base_filtros(cliente="BAGUETERIA DON MARIO SAC")),
    ("MES 09/2026", base_filtros(mes="09")),
    ("RANGO 01-15 sep 2026", base_filtros(modo_fecha="Rango", periodo_txt="Rango 01/09-15/09/2026",
                                          dt_ini=datetime.datetime(2026,9,1), dt_fin=datetime.datetime(2026,9,15))),
    ("PLACA + CATEGORIA", base_filtros(placa="CHD817", categoria="Combustible y Peajes")),
]

for nombre, f in casos:
    try:
        app = crear_app(); d = app._recolectar_datos(f)
    except Exception:
        import traceback
        p("###", nombre, "EXCEPCION:"); p(traceback.format_exc()); continue
    if d is None:
        p("###", nombre, "-> SIN CONEXION"); continue
    S = d["secciones"]
    res = {fila[0]: fila[1] for fila in S["resumen"]["filas"]}
    p("###", nombre)
    p("   Compras:", res.get("Compras del Periodo (Neto)"),
      "| Pagado:", res.get("Pagado en el Periodo (Dinero Real)"),
      "| SaldosBancos:", res.get("Saldo Total Disponible en Bancos"))
    p("   Cobranza:", res.get("Cobranza Registrada (Total)"),
      "| Combustible:", res.get("Importe de Combustible"), res.get("Galones / m³ de Combustible"))
    p("   conteos:", {k: len(S[k]["filas"]) for k in
       ["categorias","placas","proveedores_res","cuentas","compras","pagos","combustible",
        "cobranza","cobranza_unidades","flota","ordenes_servicio","conciliacion","bitacora","clientes"]})
    p("   conciliacion:", json.dumps(S["conciliacion"]["filas"], ensure_ascii=False, default=str))
    if S["cobranza"]["filas"]:
        p("   cobranza fila0:", json.dumps(S["cobranza"]["filas"][0], ensure_ascii=False, default=str))
    p("")
open("_test_resultado.txt","w",encoding="utf-8").write(out.getvalue())
print("OK")
