# -*- coding: utf-8 -*-
"""Reordena el capitulo 4 del instructivo para que siga el orden de las pestanas."""
import io, re, sys
ruta = r"C:\\Users\\Alberto\\Desktop\\Proyecto Control Flota 140926\INSTRUCTIVO_MODULO_NOMINA.md"
texto = open(ruta, encoding="utf-8").read()
salida = io.StringIO()
def p(*a): salida.write(" ".join(str(x) for x in a) + "\n")

inicio = texto.index("## 4. Descripción de cada pestaña")
fin = texto.index("## 5. El archivo del captahuella")
cap4, antes, despues = texto[inicio:fin], texto[:inicio], texto[fin:]

partes = re.split(r"(?m)^### 4\.\d+ ", cap4)
intro, secciones = partes[0], partes[1:]
p("Secciones encontradas: %s" % ", ".join(s.split(chr(10), 1)[0].strip() for s in secciones))
mapa = {s.split(chr(10), 1)[0].strip(): s for s in secciones}

orden = ["Panel", "Personal", "Configuración", "Marcaciones", "Asistencia", "Planilla", "Reportes"]
faltan = [n for n in orden if n not in mapa]
if faltan:
    p("ERROR: faltan secciones: %s" % faltan); print(salida.getvalue()); sys.exit(1)
sobran = [n for n in mapa if n not in orden]
if sobran:
    p("AVISO: secciones no contempladas: %s" % sobran)

nuevo_intro = """## 4. Descripción de cada pestaña

**Las pestañas están ordenadas igual que los pasos del asistente**, para que usted avance de
izquierda a derecha sin tener que pensar qué toca hacer:

| Pestaña | Pasos del asistente que se resuelven aquí |
|---|---|
| 📊 Panel | Inicio: estado general del mes y botón del asistente |
| 👥 Personal | 1 padrón · 2 turnos · 3 horarios · 4 asignar horario · 5 sueldos |
| ⚙️ Configuración | 6 parámetros de cálculo · 7 feriados |
| 📥 Marcaciones | 8 importar las marcaciones · 9 vincular las personas del reloj |
| 🗓️ Asistencia | 10 recalcular y revisar la asistencia |
| 💵 Planilla | 11 calcular la planilla y emitir boletas |
| 📈 Reportes | Consultas y exportación de resultados |

"""
nuevo_cap4 = nuevo_intro + "\n".join("### 4.%d %s" % (i, mapa[nombre]) for i, nombre in enumerate(orden, 1))
if not nuevo_cap4.endswith("\n"):
    nuevo_cap4 += "\n"
open(ruta, "w", encoding="utf-8").write(antes + nuevo_cap4 + despues)

p("")
p("### NUEVO ORDEN DE LAS SECCIONES")
for linea in re.findall(r"(?m)^### 4\.\d+ .+$", nuevo_cap4):
    p("  " + linea)
p("")
p("Longitud del capitulo: antes %d, despues %d caracteres" % (len(cap4), len(nuevo_cap4)))
p("El resto del documento quedo intacto: %s" % (len(antes) + len(despues) == len(texto) - len(cap4)))
print(salida.getvalue())
