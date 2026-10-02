# -*- coding: utf-8 -*-
"""Busca nombres usados como globales que no existen en el modulo (posibles NameError silenciosos)."""
import ast, builtins, glob, io, os, symtable, sys, traceback
sys.path.insert(0, r"C:\Users\Alberto\Desktop\Proyecto Control Flota 140926")
os.chdir(r"C:\Users\Alberto\Desktop\Proyecto Control Flota 140926")

archivos = sorted(glob.glob("*.py"))
hallazgos = []
for archivo in archivos:
    try:
        fuente = open(archivo, "r", encoding="utf-8", errors="ignore").read()
        tabla = symtable.symtable(fuente, archivo, "exec")
    except Exception:
        continue
    definidos = set()
    for sym in tabla.get_symbols():
        if sym.is_assigned() or sym.is_imported() or sym.is_namespace():
            definidos.add(sym.get_name())
    # nombres traidos con "from x import *"
    comodin = any(isinstance(n, ast.ImportFrom) and any(a.name == "*" for a in n.names)
                  for n in ast.walk(ast.parse(fuente)))
    def revisar(t, ruta):
        for sym in t.get_symbols():
            n = sym.get_name()
            if sym.is_referenced() and sym.is_global() and n not in definidos and not hasattr(builtins, n):
                hallazgos.append((archivo, n, ruta))
        for hijo in t.get_children():
            revisar(hijo, ruta + "." + hijo.get_name())
    revisar(tabla, archivo)

print("ARCHIVOS REVISADOS:", len(archivos))
vistos = {}
for archivo, nombre, donde in hallazgos:
    vistos.setdefault((archivo, nombre), []).append(donde)
print("POSIBLES NOMBRES NO DEFINIDOS:", len(vistos))
for (archivo, nombre), donde in sorted(vistos.items()):
    print("  %-28s %-32s (%d usos) ej: %s" % (archivo, nombre, len(donde), donde[0][:70]))
