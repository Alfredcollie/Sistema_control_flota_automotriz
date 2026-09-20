# -*- coding: utf-8 -*-
"""
GENERAR_INSTRUCTIVO_PDF.PY - Convierte el instructivo de nómina a PDF
=====================================================================
Lee 'INSTRUCTIVO_MODULO_NOMINA.md' y genera 'INSTRUCTIVO_MODULO_NOMINA.pdf'
con portada, índice, encabezado, pie de página y tablas con el color del sistema.

Uso:  python generar_instructivo_pdf.py

Soporta el subconjunto de Markdown que usa el instructivo: títulos (#, ##, ###),
párrafos, listas (- y 1.), tablas (| ... |), bloques de código indentados,
citas (>), líneas separadoras (---), negrita (**texto**) y cursiva (*texto*).

Autor: Collie Software
"""
import os
import re
import sys

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer,
                                Table, TableStyle, PageBreak, KeepTogether)

CARPETA = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_MD = os.path.join(CARPETA, "INSTRUCTIVO_MODULO_NOMINA.md")
ARCHIVO_PDF = os.path.join(CARPETA, "INSTRUCTIVO_MODULO_NOMINA.pdf")

AZUL = colors.HexColor("#1f538d")
AZUL_CLARO = colors.HexColor("#e8eef6")
GRIS = colors.HexColor("#7f8c8d")
GRIS_CLARO = colors.HexColor("#f4f6f8")
VERDE = colors.HexColor("#27ae60")
ANCHO_UTIL = A4[0] - 3.2 * cm

# Emojis y símbolos que no existen en las fuentes del PDF
REEMPLAZOS = {
    "\u2192": "->", "\u25c0": "<", "\u25b6": ">", "\u2265": ">=", "\u2264": "<=",
    "\u2013": "-", "\u201c": '"', "\u201d": '"', "\u2018": "'", "\u2019": "'",
}


def limpiar_texto(texto):
    """Quita emojis (no se pueden dibujar) y reemplaza símbolos por equivalentes."""
    salida = []
    for caracter in texto:
        codigo = ord(caracter)
        if 0x1F000 <= codigo <= 0x1FAFF or 0x2600 <= codigo <= 0x27BF or codigo in (0xFE0F, 0x2B50, 0x2705, 0x274C):
            continue
        if 0x1F1E6 <= codigo <= 0x1F1FF:      # banderas
            continue
        salida.append(REEMPLAZOS.get(caracter, caracter))
    return "".join(salida)


def registrar_fuentes():
    """Registra una fuente con buen soporte Unicode; si no hay, usa Helvetica."""
    candidatas = [
        (r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\arialbd.ttf"),
        ("/System/Library/Fonts/Supplemental/Arial.ttf", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
    ]
    for normal, negrita in candidatas:
        if os.path.exists(normal) and os.path.exists(negrita):
            try:
                pdfmetrics.registerFont(TTFont("Instructivo", normal))
                pdfmetrics.registerFont(TTFont("Instructivo-Bold", negrita))
                pdfmetrics.registerFontFamily("Instructivo", normal="Instructivo",
                                              bold="Instructivo-Bold", italic="Instructivo",
                                              boldItalic="Instructivo-Bold")
                return "Instructivo", "Instructivo-Bold"
            except Exception as e:
                print("Aviso - no se pudo registrar la fuente %s: %s" % (normal, e))
    return "Helvetica", "Helvetica-Bold"


FUENTE, FUENTE_NEGRITA = registrar_fuentes()


def estilos():
    """Estilos del documento."""
    return {
        "titulo": ParagraphStyle("titulo", fontName=FUENTE_NEGRITA, fontSize=24, leading=30,
                                 textColor=AZUL, alignment=TA_CENTER, spaceAfter=8),
        "subtitulo": ParagraphStyle("subtitulo", fontName=FUENTE_NEGRITA, fontSize=13, leading=19,
                                    textColor=colors.HexColor("#34495e"), alignment=TA_CENTER),
        "h2": ParagraphStyle("h2", fontName=FUENTE_NEGRITA, fontSize=15, leading=20, textColor=AZUL,
                             spaceBefore=14, spaceAfter=6),
        "h3": ParagraphStyle("h3", fontName=FUENTE_NEGRITA, fontSize=12.5, leading=17,
                             textColor=colors.HexColor("#163b65"), spaceBefore=10, spaceAfter=4),
        "cuerpo": ParagraphStyle("cuerpo", fontName=FUENTE, fontSize=10.2, leading=15,
                                 alignment=TA_JUSTIFY, spaceAfter=5),
        "vineta": ParagraphStyle("vineta", fontName=FUENTE, fontSize=10.2, leading=15,
                                 leftIndent=14, bulletIndent=4, spaceAfter=2),
        "numerada": ParagraphStyle("numerada", fontName=FUENTE, fontSize=10.2, leading=15,
                                   leftIndent=18, bulletIndent=4, spaceAfter=2),
        "cita": ParagraphStyle("cita", fontName=FUENTE, fontSize=10.2, leading=15,
                               leftIndent=14, rightIndent=8, textColor=colors.HexColor("#7d6608"),
                               backColor=colors.HexColor("#fef9e7"), borderPadding=6, spaceAfter=6),
        "codigo": ParagraphStyle("codigo", fontName="Courier", fontSize=9, leading=12.5,
                                 leftIndent=16, textColor=colors.HexColor("#1a1a1a"), spaceAfter=2),
        "celda": ParagraphStyle("celda", fontName=FUENTE, fontSize=8.8, leading=11.6),
        "celda_cab": ParagraphStyle("celda_cab", fontName=FUENTE_NEGRITA, fontSize=8.8, leading=11.6,
                                    textColor=colors.white),
        "indice": ParagraphStyle("indice", fontName=FUENTE, fontSize=10, leading=16),
        "indice_sub": ParagraphStyle("indice_sub", fontName=FUENTE, fontSize=9, leading=14,
                                     leftIndent=16, textColor=colors.HexColor("#34495e")),
        "nota": ParagraphStyle("nota", fontName=FUENTE, fontSize=9, leading=13, textColor=GRIS,
                               alignment=TA_CENTER),
    }


def en_linea(texto):
    """Convierte la sintaxis en línea de Markdown al marcado de reportlab."""
    texto = limpiar_texto(texto)
    texto = texto.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    texto = texto.replace("\\_", "_")
    texto = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", texto)
    texto = re.sub(r"(?<!\*)\*([^*]+?)\*(?!\*)", r"<i>\1</i>", texto)
    return texto


def separar_tabla(lineas):
    """Convierte las líneas de una tabla Markdown en un flowable Table."""
    filas = []
    for linea in lineas:
        celdas = [c.strip() for c in linea.strip().strip("|").split("|")]
        if all(re.fullmatch(r":?-{2,}:?", c or "-") for c in celdas):
            continue
        filas.append(celdas)
    if not filas:
        return None
    ancho_columnas = max(len(f) for f in filas)
    filas = [f + [""] * (ancho_columnas - len(f)) for f in filas]
    st = estilos()
    datos = [[Paragraph(en_linea(c), st["celda_cab" if indice == 0 else "celda"])
              for c in fila] for indice, fila in enumerate(filas)]
    # Ancho proporcional al contenido, con un mínimo por columna
    largos = []
    for columna in range(ancho_columnas):
        largo = max(len(re.sub(r"[*|]", "", f[columna])) for f in filas)
        largos.append(min(max(largo, 8), 60))
    total = float(sum(largos))
    anchos = [(largo / total) * ANCHO_UTIL for largo in largos]
    tabla = Table(datos, colWidths=anchos, repeatRows=1, hAlign="LEFT")
    estilos_tabla = [
        ("BACKGROUND", (0, 0), (-1, 0), AZUL),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#d5dbdb")),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
    ]
    for indice in range(1, len(filas)):
        if indice % 2 == 0:
            estilos_tabla.append(("BACKGROUND", (0, indice), (-1, indice), GRIS_CLARO))
    tabla.setStyle(TableStyle(estilos_tabla))
    return tabla


def leer_markdown():
    with open(ARCHIVO_MD, "r", encoding="utf-8") as manejador:
        return manejador.read().splitlines()


def construir_flowables(lineas, st):
    """Recorre el Markdown y devuelve la lista de flowables y el índice de secciones."""
    flowables = []
    indices = []
    pendientes = []
    modo = [None]

    def cerrar():
        if not pendientes:
            return
        if modo[0] == "tabla":
            tabla = separar_tabla(pendientes)
            if tabla is not None:
                flowables.append(tabla)
                flowables.append(Spacer(1, 8))
        elif modo[0] == "codigo":
            for linea in pendientes:
                flowables.append(Paragraph(en_linea(linea.strip()) or "&nbsp;", st["codigo"]))
            flowables.append(Spacer(1, 6))
        elif modo[0] == "cita":
            flowables.append(Paragraph(en_linea(" ".join(pendientes)), st["cita"]))
        elif modo[0] == "vineta":
            for linea in pendientes:
                flowables.append(Paragraph(en_linea(linea), st["vineta"], bulletText="\u2022"))
            flowables.append(Spacer(1, 4))
        elif modo[0] == "numerada":
            for numero, linea in pendientes:
                flowables.append(Paragraph(en_linea(linea), st["numerada"], bulletText="%d." % numero))
            flowables.append(Spacer(1, 4))
        else:
            flowables.append(Paragraph(en_linea(" ".join(pendientes)), st["cuerpo"]))
        del pendientes[:]
        modo[0] = None

    for linea in lineas:
        limpia = linea.rstrip()
        if limpia.strip() in ("---", "***", "___"):
            cerrar()
            flowables.append(Spacer(1, 4))
            continue
        if not limpia.strip():
            cerrar()
            continue
        if limpia.startswith("|"):
            if modo[0] != "tabla":
                cerrar()
                modo[0] = "tabla"
            pendientes.append(limpia)
            continue
        if limpia.startswith("    ") or limpia.startswith("\t"):
            if modo[0] != "codigo":
                cerrar()
                modo[0] = "codigo"
            pendientes.append(limpia)
            continue
        if limpia.startswith("> "):
            if modo[0] != "cita":
                cerrar()
                modo[0] = "cita"
            pendientes.append(limpia[2:])
            continue
        coincidencia = re.match(r"^(\s*)[-*]\s+(.*)$", limpia)
        if coincidencia and not limpia.lstrip().startswith("|"):
            if modo[0] != "vineta":
                cerrar()
                modo[0] = "vineta"
            pendientes.append(coincidencia.group(2))
            continue
        coincidencia = re.match(r"^\s*(\d+)\.\s+(.*)$", limpia)
        if coincidencia:
            if modo[0] != "numerada":
                cerrar()
                modo[0] = "numerada"
            pendientes.append((int(coincidencia.group(1)), coincidencia.group(2)))
            continue
        if limpia.startswith("### "):
            cerrar()
            titulo = limpia[4:].strip()
            indices.append((3, titulo))
            flowables.append(Paragraph(en_linea(titulo), st["h3"]))
            continue
        if limpia.startswith("## "):
            cerrar()
            titulo = limpia[3:].strip()
            indices.append((2, titulo))
            flowables.append(Paragraph(en_linea(titulo), st["h2"]))
            continue
        if limpia.startswith("# "):
            cerrar()
            continue  # el título principal va en la portada
        if modo[0] not in (None, "cuerpo"):
            cerrar()
        modo[0] = "cuerpo"
        pendientes.append(limpia.strip())
    cerrar()
    return flowables, indices


def portada(st, indices):
    """Portada e índice del instructivo."""
    elementos = [Spacer(1, 3.2 * cm)]
    elementos.append(Paragraph("INSTRUCTIVO", st["titulo"]))
    elementos.append(Paragraph("MÓDULO DE NÓMINA Y CONTROL DE ASISTENCIA", st["subtitulo"]))
    elementos.append(Spacer(1, 0.4 * cm))
    elementos.append(Paragraph("Sistema de Control de Flota Automotriz", st["subtitulo"]))
    elementos.append(Spacer(1, 1.6 * cm))
    ficha = [
        ["Área", "Nómina, asistencia y planilla"],
        ["Módulo", "Nómina y Asistencia (menú Finanzas y Reportes)"],
        ["Reloj compatible", "Hikvision, ZKTeco / ZKTime y CSV genérico"],
        ["Versión", "1.0 - setiembre 2026"],
        ["Desarrollado por", "Collie Software"],
    ]
    tabla = Table([[Paragraph("<b>%s</b>" % a, st["celda"]), Paragraph(b, st["celda"])] for a, b in ficha],
                  colWidths=[4.2 * cm, 9.0 * cm], hAlign="CENTER")
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), AZUL_CLARO),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#c8d6e5")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]))
    elementos.append(tabla)
    elementos.append(Spacer(1, 1.4 * cm))
    elementos.append(Paragraph("Documento de uso interno. Ante cualquier duda sobre un cálculo, revise la "
                               "boleta del empleado: cada línea indica la fórmula exacta que se utilizó.",
                               st["nota"]))
    elementos.append(PageBreak())

    elementos.append(Paragraph("Contenido", st["h2"]))
    for nivel, titulo in indices:
        if nivel == 2:
            elementos.append(Paragraph(en_linea(titulo), st["indice"]))
        else:
            elementos.append(Paragraph(en_linea(titulo), st["indice_sub"]))
    elementos.append(PageBreak())
    return elementos


def decorar(canvas, documento):
    """Encabezado y pie de página."""
    canvas.saveState()
    ancho, alto = A4
    canvas.setStrokeColor(AZUL)
    canvas.setLineWidth(1.1)
    canvas.line(1.6 * cm, alto - 1.45 * cm, ancho - 1.6 * cm, alto - 1.45 * cm)
    canvas.setFont(FUENTE, 8)
    canvas.setFillColor(AZUL)
    canvas.drawString(1.6 * cm, alto - 1.25 * cm, "Instructivo del Módulo de Nómina y Asistencia")
    canvas.setFillColor(GRIS)
    canvas.drawRightString(ancho - 1.6 * cm, alto - 1.25 * cm, "Collie Software")
    canvas.setStrokeColor(colors.HexColor("#d5dbdb"))
    canvas.setLineWidth(0.6)
    canvas.line(1.6 * cm, 1.35 * cm, ancho - 1.6 * cm, 1.35 * cm)
    canvas.setFont(FUENTE, 8.5)
    canvas.setFillColor(GRIS)
    canvas.drawCentredString(ancho / 2.0, 1.0 * cm, "Página %d" % documento.page)
    canvas.restoreState()


def main():
    if not os.path.exists(ARCHIVO_MD):
        print("No se encontró el archivo: %s" % ARCHIVO_MD)
        return 1
    st = estilos()
    lineas = leer_markdown()
    cuerpo, indices = construir_flowables(lineas, st)
    documento = BaseDocTemplate(ARCHIVO_PDF, pagesize=A4,
                                leftMargin=1.6 * cm, rightMargin=1.6 * cm,
                                topMargin=1.8 * cm, bottomMargin=1.7 * cm,
                                title="Instructivo del Módulo de Nómina y Asistencia",
                                author="Collie Software", subject="Manual de usuario")
    marco = Frame(documento.leftMargin, documento.bottomMargin, documento.width, documento.height,
                  id="normal")
    documento.addPageTemplates([PageTemplate(id="todas", frames=[marco], onPage=decorar)])
    contenido = portada(st, indices) + cuerpo
    documento.build(contenido)
    tamano = os.path.getsize(ARCHIVO_PDF)
    print("PDF generado: %s (%.1f KB, %d páginas)" % (ARCHIVO_PDF, tamano / 1024.0, documento.page))
    return 0


if __name__ == "__main__":
    sys.exit(main())
