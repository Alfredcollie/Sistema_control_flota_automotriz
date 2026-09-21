# -*- coding: utf-8 -*-
import io, os, sys, time, warnings
warnings.filterwarnings("ignore")
sys.path.insert(0, r"C:\Users\Alberto\Desktop\Proyecto Control Flota 140926")
out = io.StringIO()
def p(*a): out.write(" ".join(str(x) for x in a) + "\n")
import tkinter as tk
from tkinter import messagebox
for f in ("showinfo", "showerror", "showwarning"):
    setattr(messagebox, f, lambda *a, **k: True)
import customtkinter as ctk
import modulo_nomina
def bombear(s):
    fin = time.time() + s
    while time.time() < fin:
        root.update(); time.sleep(0.05)
root = ctk.CTk(); root.geometry("1200x820+10+10")
ruta = os.path.join(r"C:\Users\Alberto\Desktop\Proyecto Control Flota 140926", "INSTRUCTIVO_MODULO_NOMINA.pdf")
visor = modulo_nomina.VisorInstructivo(root, ruta)
bombear(4)

def buscar(termino):
    visor.ent_buscar.delete(0, tk.END); visor.ent_buscar.insert(0, termino)
    t0 = time.time()
    visor.buscar()
    bombear(1.2)
    return time.time() - t0

p("### BUSQUEDA RAPIDA")
for termino in ("turno", "nomina", "nómina", "horas extra", "sueldo", "tardanza", "palabra-que-no-existe"):
    segundos = buscar(termino)
    p("  %-22r -> %-70s (%.2f s, en la pagina %d)" % (termino, visor.lbl_busqueda.cget("text"),
                                                      segundos, visor.numero + 1))

p("")
p("### NAVEGAR ENTRE COINCIDENCIAS")
buscar("turno")
p("  total de coincidencias: %d" % len(visor.coincidencias))
recorrido = [visor.numero + 1]
for _ in range(5):
    visor.ir_a_coincidencia(1); bombear(0.8)
    recorrido.append(visor.numero + 1)
p("  paginas visitadas yendo hacia adelante: %s" % recorrido)
visor.ir_a_coincidencia(-1); bombear(0.8)
p("  al retroceder -> %s" % visor.lbl_busqueda.cget("text"))
p("  la imagen se dibujo con resaltado: %s" % (visor._imagen is not None))

p("")
p("### SALTO DE PAGINA CON BUSQUEDA ACTIVA")
visor.mostrar_pagina(9); bombear(1.2)
p("  pagina 10 -> %s | imagen=%s" % (visor.lbl_estado.cget("text"), visor._imagen is not None))

p("")
p("### LIMPIAR")
visor.limpiar_busqueda(); bombear(1.2)
p("  mensaje: %s" % visor.lbl_busqueda.cget("text"))
p("  coincidencias: %d | consulta: %r" % (len(visor.coincidencias), visor.consulta))
p("  imagen sigue dibujada: %s" % (visor._imagen is not None))
try:
    visor.destroy(); root.destroy()
except Exception:
    pass
open(r"C:\Users\Alberto\Desktop\Proyecto Control Flota 140926\_test_busqueda.txt", "w", encoding="utf-8").write(out.getvalue())
print("OK")
