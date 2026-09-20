# -*- coding: utf-8 -*-
import io, sys, warnings
warnings.filterwarnings("ignore")
sys.path.insert(0, r"C:\\Users\\Alberto\\Desktop\\Proyecto Control Flota 140926")
import nomina_core as nc
out = io.StringIO()
def p(*a): out.write(" ".join(str(x) for x in a) + "\n")
for periodo in ("2026-09", "2026-04"):
    estado = nc.estado_configuracion(periodo)
    p("### PERIODO %s" % periodo)
    p("avance: %d%% | listo: %s | criticos %d/%d | pendientes %d | atenciones %d | ok %d" % (
        estado["avance"], estado["listo"], estado["criticos_ok"], estado["criticos_total"],
        len(estado["pendientes"]), len(estado["atenciones"]), len(estado["listos"])))
    for indice, paso in enumerate(estado["pasos"], 1):
        p("%2d. [%-9s] %-38s (%s)" % (indice, paso["estado"], paso["titulo"], paso["fase"]))
        p("      detalle: %s" % paso["detalle"])
        p("      destino: %s | accion: %s" % (paso["destino"], paso["accion"]))
    if estado["siguiente"]:
        p("SIGUIENTE: %s" % estado["siguiente"]["titulo"])
    p("")
open(r"C:\\Users\\Alberto\\Desktop\\Proyecto Control Flota 140926\_test_asistente.txt", "w", encoding="utf-8").write(out.getvalue())
print("OK")
