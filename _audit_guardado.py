# -*- coding: utf-8 -*-
"""Auditoría: dónde guarda archivos cada módulo (PDF/JPG/PNG)."""
import io, os, re, sys, json
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

IGNORAR = {"deepseek-harness", "__pycache__", ".git", "node_modules"}
PATRONES = [
    ("copy", re.compile(r"\bshutil\.copy2?\s*\(")),
    ("makedirs", re.compile(r"\bos\.makedirs\s*\(")),
    ("open_wb", re.compile(r"open\s*\([^)]*['\"](?:wb|ab)['\"]")),
    ("img_save", re.compile(r"\.save\s*\(")),
    ("guardar_dialogo", re.compile(r"guardar_archivo_dialogo\s*\(")),
    ("ruta_para_guardar", re.compile(r"ruta_para_guardar\s*\(")),
]
BASE = re.compile(r"(os\.path\.dirname\(os\.path\.abspath\(__file__\)\)|DATA_DIR|ruta_drive_configurada|ruta_base_autorizada|obtener_ruta_base_drive|_carpeta_archivos|_carpeta_expedientes|obtener_ruta_base|ruta_base|ruta_drive|APPDATA|expanduser)")

filas = []
for raiz, dirs, archivos in os.walk("."):
    dirs[:] = [d for d in dirs if d not in IGNORAR and not d.startswith(".")]
    for nombre in archivos:
        if not nombre.endswith(".py") or nombre.startswith("_"):
            continue
        ruta = os.path.join(raiz, nombre)
        try:
            lineas = io.open(ruta, encoding="utf-8").read().split("\n")
        except Exception:
            continue
        for i, linea in enumerate(lineas, 1):
            for etiqueta, pat in PATRONES:
                if pat.search(linea):
                    contexto = " ".join(l.strip() for l in lineas[max(0, i-6):i+2])
                    bases = set(BASE.findall(contexto))
                    filas.append((nombre, i, etiqueta, sorted(bases), linea.strip()[:110]))
                    break

print("TOTAL de puntos de guardado detectados:", len(filas))
por_modulo = {}
for mod, ln, et, bases, txt in filas:
    por_modulo.setdefault(mod, []).append((ln, et, bases, txt))
for mod in sorted(por_modulo, key=lambda m: -len(por_modulo[m])):
    usos = por_modulo[mod]
    todas = sorted({b for _, _, bs, _ in usos for b in bs})
    print("\n=== %s (%d) -> bases: %s" % (mod, len(usos), ", ".join(todas) or "SIN REFERENCIA CLARA"))
    for ln, et, bases, txt in usos:
        print("   %-5s %-16s %-42s %s" % (ln, et, ",".join(bases)[:42], txt[:95]))
