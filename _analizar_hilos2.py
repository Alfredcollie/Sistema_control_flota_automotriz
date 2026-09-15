# -*- coding: utf-8 -*-
"""Análisis ampliado: OTRAS llamadas a la interfaz hechas dentro de hilos.

Busca, en funciones alcanzables desde threading.Thread(target=...), llamadas a:
  - messagebox / filedialog / simpledialog  (diálogos desde un hilo: cuelga en macOS)
  - after / update / update_idletasks / clipboard_*
  - métodos de widgets: configure, insert, delete, pack, grid, place, destroy, winfo_*, set
"""
import ast
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PELIGROSOS = {
    "after", "after_cancel", "update", "update_idletasks",
    "clipboard_clear", "clipboard_append", "clipboard_get",
    "configure", "pack", "grid", "place", "destroy", "winfo_children",
    "showinfo", "showerror", "showwarning", "askyesno", "askstring", "askopenfilename",
    "askopenfilenames", "askdirectory", "asksaveasfilename", "setvar",
}
DIALOGOS = {"messagebox", "filedialog", "simpledialog"}
# Métodos muy usados también fuera de la interfaz: se revisan aparte
AMBIGUOS = {"insert", "delete", "set", "get", "focus"}

def nombre_llamada(nodo):
    f = nodo.func
    if isinstance(f, ast.Name):
        return f.id, None
    if isinstance(f, ast.Attribute):
        base = f.value.id if isinstance(f.value, ast.Name) else None
        return f.attr, base
    return None, None

archivos = sorted(f for f in os.listdir(".") if f.endswith(".py") and not f.startswith("_"))
total = 0
for ruta in archivos:
    try:
        arbol = ast.parse(open(ruta, encoding="utf-8").read())
    except Exception:
        continue
    funciones = {}
    for nodo in ast.walk(arbol):
        if isinstance(nodo, (ast.FunctionDef, ast.AsyncFunctionDef)):
            llamadas, hallazgos = set(), []
            for sub in ast.walk(nodo):
                if isinstance(sub, ast.Call):
                    nombre, base = nombre_llamada(sub)
                    if nombre:
                        llamadas.add(nombre)
                    if nombre in PELIGROSOS or (base in DIALOGOS):
                        hallazgos.append((sub.lineno, f"{base + '.' if base else ''}{nombre}"))
            funciones[nodo.name] = {"llamadas": llamadas, "hallazgos": hallazgos, "linea": nodo.lineno}
    objetivos = set()
    for nodo in ast.walk(arbol):
        if isinstance(nodo, ast.Call):
            nombre, _ = nombre_llamada(nodo)
            if nombre == "Thread":
                for kw in nodo.keywords:
                    if kw.arg == "target":
                        t = kw.value
                        if isinstance(t, ast.Name):
                            objetivos.add(t.id)
                        elif isinstance(t, ast.Attribute):
                            objetivos.add(t.attr)
    # Solo las funciones del propio hilo (sin expandir a todo lo que llaman) para evitar ruido
    problemas = []
    for f in sorted(objetivos):
        info = funciones.get(f)
        if info and info["hallazgos"]:
            problemas.append((f, info["linea"], info["hallazgos"]))
    if problemas:
        print(f"\n{ruta}")
        for nombre, linea, hallazgos in problemas:
            print(f"   {nombre}() línea {linea}:")
            for ln, que in hallazgos:
                print(f"        línea {ln}: {que}()")
            total += len(hallazgos)
print(f"\nTOTAL de llamadas de interfaz directamente en funciones de hilo: {total}")
