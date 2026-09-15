# -*- coding: utf-8 -*-
"""Detecta funciones que corren en hilos y que llaman a .after() (riesgo de cuelgue en macOS)."""
import ast
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ARCHIVOS = [f for f in os.listdir(".") if f.endswith(".py") and not f.startswith("_")]

def nombre_llamada(nodo):
    f = nodo.func
    if isinstance(f, ast.Name):
        return f.id
    if isinstance(f, ast.Attribute):
        return f.attr
    return None

def analizar(ruta):
    fuente = open(ruta, encoding="utf-8").read()
    try:
        arbol = ast.parse(fuente)
    except Exception as e:
        return None, f"ERROR de sintaxis: {e}"

    # Mapa: nombre de función -> lista de (nombre de función llamada)
    funciones = {}
    for nodo in ast.walk(arbol):
        if isinstance(nodo, (ast.FunctionDef, ast.AsyncFunctionDef)):
            llamadas = set()
            afters = []
            for sub in ast.walk(nodo):
                if isinstance(sub, ast.Call):
                    n = nombre_llamada(sub)
                    if n:
                        llamadas.add(n)
                    if n == "after":
                        afters.append(sub.lineno)
            funciones[nodo.name] = {"llamadas": llamadas, "afters": afters, "linea": nodo.lineno}

    # Objetivos de hilos
    objetivos = set()
    for nodo in ast.walk(arbol):
        if isinstance(nodo, ast.Call):
            n = nombre_llamada(nodo)
            if n == "Thread":
                for kw in nodo.keywords:
                    if kw.arg == "target":
                        t = kw.value
                        if isinstance(t, ast.Name):
                            objetivos.add(t.id)
                        elif isinstance(t, ast.Attribute):
                            objetivos.add(t.attr)

    # Expansión transitiva (funciones llamadas por los hilos)
    pendientes = list(objetivos)
    vistos = set()
    while pendientes:
        f = pendientes.pop()
        if f in vistos:
            continue
        vistos.add(f)
        info = funciones.get(f)
        if info:
            pendientes.extend(info["llamadas"])

    problemas = [(f, funciones[f]["linea"], funciones[f]["afters"])
                 for f in sorted(vistos) if f in funciones and funciones[f]["afters"]]
    return problemas, None

total = 0
for ruta in sorted(ARCHIVOS):
    problemas, err = analizar(ruta)
    if err:
        continue
    if problemas:
        total += sum(len(p[2]) for p in problemas)
        print(f"\n{ruta}")
        for nombre, linea, afters in problemas:
            print(f"   {nombre}()  (línea {linea})  ->  .after() en líneas {afters}")
print(f"\nTOTAL de llamadas .after() dentro de hilos: {total}")
