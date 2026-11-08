# -*- coding: utf-8 -*-
"""
CONTRATOS_ACTAS_PDF.PY — Genera los PDF de Contratos y Actas
================================================================
Usa el texto de "plantillas_contratos.py" (idéntico a los documentos del
cliente) y reportlab para armar el PDF con márgenes, títulos y firmas.
"""
import os
import re

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (KeepTogether, Paragraph, SimpleDocTemplate,
                                Spacer, Table, TableStyle)

import plantillas_contratos as PL

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]


# =========================================================
# NÚMEROS EN LETRAS (para la remuneración)
# =========================================================
_UNIDADES = ["cero", "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho",
             "nueve", "diez", "once", "doce", "trece", "catorce", "quince",
             "dieciséis", "diecisiete", "dieciocho", "diecinueve", "veinte"]
_DECENAS = {30: "treinta", 40: "cuarenta", 50: "cincuenta", 60: "sesenta",
            70: "setenta", 80: "ochenta", 90: "noventa"}
_CENTENAS = {100: "cien", 200: "doscientos", 300: "trescientos", 400: "cuatrocientos",
             500: "quinientos", 600: "seiscientos", 700: "setecientos",
             800: "ochocientos", 900: "novecientos"}


def _hasta_999(numero):
    """Número de 1 a 999 en letras."""
    if numero <= 20:
        return _UNIDADES[numero]
    if numero < 30:
        return ("veinti" + _UNIDADES[numero - 20]).replace("veintiseis", "veintiséis")
    if numero < 100:
        decena = (numero // 10) * 10
        resto = numero % 10
        return _DECENAS[decena] + ((" y " + _UNIDADES[resto]) if resto else "")
    if numero == 100:
        return "cien"
    centena = (numero // 100) * 100
    resto = numero % 100
    texto = _CENTENAS.get(centena, "")
    if centena == 100:
        texto = "ciento"
    return texto + ((" " + _hasta_999(resto)) if resto else "")


def numero_a_letras(numero):
    """Convierte un entero a letras (hasta millones)."""
    numero = int(numero)
    if numero == 0:
        return "cero"
    if numero < 1000:
        return _hasta_999(numero)
    if numero < 1000000:
        miles = numero // 1000
        resto = numero % 1000
        prefijo = "mil" if miles == 1 else _hasta_999(miles) + " mil"
        return prefijo + ((" " + _hasta_999(resto)) if resto else "")
    millones = numero // 1000000
    resto = numero % 1000000
    prefijo = ("un millón" if millones == 1 else _hasta_999(millones) + " millones")
    return prefijo + ((" " + numero_a_letras(resto)) if resto else "")


def remuneracion_en_letras(monto):
    """'mil doscientos treinta con 00/100 soles'."""
    try:
        valor = float(monto or 0)
    except (TypeError, ValueError):
        valor = 0.0
    enteros = int(valor)
    centavos = int(round((valor - enteros) * 100))
    if centavos >= 100:
        enteros += 1
        centavos -= 100
    return "%s con %02d/100 soles" % (numero_a_letras(enteros), centavos)


# =========================================================
# ESTILOS
# =========================================================
def _estilos(margen_superior=2.0 * cm):
    base = dict(fontName="Helvetica", fontSize=10.5, leading=14.5,
                alignment=TA_JUSTIFY, spaceAfter=7)
    return {
        "titulo": ParagraphStyle("titulo", fontName="Helvetica-Bold", fontSize=14,
                                 leading=18, alignment=TA_CENTER, spaceAfter=14),
        "subtitulo": ParagraphStyle("subtitulo", fontName="Helvetica-Bold", fontSize=10.5,
                                    leading=14, alignment=TA_LEFT,
                                    spaceBefore=8, spaceAfter=4),
        "subsubtitulo": ParagraphStyle("subsubtitulo", fontName="Helvetica-Bold",
                                       fontSize=10.5, leading=14, alignment=TA_LEFT,
                                       spaceBefore=6, spaceAfter=3),
        "clausula": ParagraphStyle("clausula", **base),
        "parrafo": ParagraphStyle("parrafo", **base),
        "vineta": ParagraphStyle("vineta", fontName="Helvetica", fontSize=10.5,
                                 leading=14, leftIndent=14, bulletIndent=4,
                                 alignment=TA_LEFT, spaceAfter=2),
        "vineta_vehiculo": ParagraphStyle("vineta_vehiculo", fontName="Helvetica",
                                          fontSize=10.5, leading=14, leftIndent=14,
                                          bulletIndent=4, alignment=TA_LEFT, spaceAfter=2),
        "firma": ParagraphStyle("firma", fontName="Helvetica", fontSize=10.5,
                                leading=15, alignment=TA_CENTER),
    }


def _escapar(texto):
    return (str(texto or "").replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;"))


# =========================================================
# MARCADORES
# =========================================================
def _fecha_partes(fecha):
    """'05/11/2026' -> ('05', 'noviembre', '2026')."""
    texto = str(fecha or "").strip()
    partes = re.findall(r"\d+", texto)
    if len(partes) >= 3:
        dia, mes, anio = partes[0], int(partes[1]), partes[2]
        nombre = MESES[mes - 1] if 1 <= mes <= 12 else ""
        return dia, nombre, anio
    return "", "", ""


def marcadores(datos):
    """Diccionario completo de valores para las plantillas."""
    datos = dict(datos or {})
    dia, mes, anio = _fecha_partes(datos.get("fecha_firma"))
    base = {
        "empresa": datos.get("empresa") or "",
        "ruc_empresa": datos.get("ruc_empresa") or "",
        "domicilio_empresa": datos.get("domicilio_empresa") or "",
        "trabajador": datos.get("trabajador") or "",
        "dni": datos.get("dni") or "",
        "direccion_trabajador": datos.get("direccion_trabajador") or "",
        "cargo": datos.get("cargo") or "",
        "remuneracion": datos.get("remuneracion_texto") or "",
        "remuneracion_letras": datos.get("remuneracion_letras") or "",
        "fecha_inicio": datos.get("fecha_inicio") or "",
        "fecha_fin": datos.get("fecha_fin") or "",
        "fecha_inicio_original": datos.get("fecha_inicio_original") or "",
        "vencimiento_anterior": datos.get("vencimiento_anterior") or "",
        "ciudad_firma": datos.get("ciudad_firma") or "",
        "dia_firma": dia,
        "mes_firma": mes,
        "anio_firma": anio,
        "vehiculo_marca": datos.get("vehiculo_marca") or "",
        "vehiculo_modelo": datos.get("vehiculo_modelo") or "",
        "vehiculo_anio": datos.get("vehiculo_anio") or "",
        "vehiculo_placa": datos.get("vehiculo_placa") or "",
        "celular_marca_modelo": datos.get("celular_marca_modelo") or "",
    }
    return base


def reemplazar(texto, valores):
    """Sustituye {marcadores} evitando fallos si falta alguno."""
    def _cambio(coincidencia):
        clave = coincidencia.group(1)
        return str(valores.get(clave, coincidencia.group(0)))
    return re.sub(r"\{([a-z_]+)\}", _cambio, str(texto or ""))


# =========================================================
# CONSTRUCCIÓN DEL PDF
# =========================================================
def _parrafo_con_negritas(texto, estilo, valores):
    """Aplica negritas a las partes habituales del contrato."""
    seguro = _escapar(texto)
    for clave in ("empresa", "trabajador", "dni", "ruc_empresa", "fecha_inicio",
                  "fecha_fin", "fecha_inicio_original", "vencimiento_anterior",
                  "remuneracion", "cargo", "domicilio_empresa",
                  "direccion_trabajador", "vehiculo_placa", "vehiculo_marca",
                  "vehiculo_modelo"):
        valor = _escapar(valores.get(clave) or "")
        if not valor:
            continue
        seguro = seguro.replace(valor, "<b>" + valor + "</b>", 1) if clave in (
            "empresa", "trabajador") else seguro.replace(valor, "<b>" + valor + "</b>")
    for etiqueta in ("EL EMPLEADOR", "EL TRABAJADOR"):
        seguro = seguro.replace(etiqueta, "<b>" + etiqueta + "</b>")
    return Paragraph(seguro, estilo)



def _bloque_firmas(valores, estilos):
    """Bloque de firmas: empleador (razón social + RUC) y trabajador (nombre + DNI)."""
    lineas = Table([
        [Paragraph("______________________________", estilos["firma"]),
         Paragraph("______________________________", estilos["firma"])],
        [Paragraph("<b>" + _escapar(valores.get("empresa") or "") + "</b>",
                   estilos["firma"]),
         Paragraph("<b>" + _escapar(valores.get("trabajador") or "") + "</b>",
                   estilos["firma"])],
        [Paragraph("R.U.C. N° " + _escapar(valores.get("ruc_empresa") or ""),
                   estilos["firma"]),
         Paragraph("DNI " + _escapar(valores.get("dni") or ""), estilos["firma"])],
    ], colWidths=[7.0 * cm, 7.0 * cm])
    lineas.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"),
                                ("TOPPADDING", (0, 0), (-1, -1), 1),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
    return lineas


def _historia(bloques, valores, estilos):
    """Convierte los bloques de la plantilla en elementos del PDF."""
    historia = []
    for tipo, contenido in bloques:
        texto = reemplazar(contenido, valores)
        if tipo == "titulo":
            historia.append(Paragraph(_escapar(texto), estilos["titulo"]))
        elif tipo == "subtitulo":
            historia.append(Paragraph(_escapar(texto), estilos["subtitulo"]))
        elif tipo == "subsubtitulo":
            historia.append(Paragraph(_escapar(texto), estilos["subsubtitulo"]))
        elif tipo == "vineta":
            historia.append(Paragraph(_escapar(texto), estilos["vineta"], bulletText="•"))
        elif tipo == "vineta_vehiculo":
            historia.append(Paragraph(_escapar(texto), estilos["vineta_vehiculo"],
                                      bulletText="•"))
        elif tipo == "firma":
            historia.append(Spacer(1, 1.4 * cm))
            historia.append(_bloque_firmas(valores, estilos))
        else:
            historia.append(_parrafo_con_negritas(texto, estilos["parrafo"], valores))
    return historia


def _crear_pdf(ruta, historia, margenes):
    """Escribe el PDF con los márgenes indicados (en cm)."""
    izquierda, derecha, superior, inferior = margenes
    carpeta = os.path.dirname(ruta)
    if carpeta and not os.path.isdir(carpeta):
        os.makedirs(carpeta, exist_ok=True)
    documento = SimpleDocTemplate(
        ruta, pagesize=A4,
        leftMargin=izquierda * cm, rightMargin=derecha * cm,
        topMargin=superior * cm, bottomMargin=inferior * cm,
        title=os.path.basename(ruta), author="Sistema de Control de Flota")
    documento.build(historia)
    return ruta


def generar_contrato(datos, ruta, renovacion=False):
    """Genera el contrato inicial o el contrato de renovación."""
    valores = marcadores(datos)
    estilos = _estilos()
    bloques = PL.RENOVACION_BLOQUES if renovacion else PL.CONTRATO_BLOQUES
    historia = _historia(bloques, valores, estilos)
    return _crear_pdf(ruta, historia, (3.0, 3.0, 2.5, 2.5))



def _tabla_bienes(datos, estilos):
    """Tabla con los bienes entregados (lista editable de cada acta)."""
    bienes = list(datos.get("bienes") or PL.ACTA_BIENES_POR_DEFECTO)
    filas = [[Paragraph("<b>Bien / equipo entregado</b>", estilos["vineta"]),
              Paragraph("<b>Cantidad</b>", estilos["vineta"])]]
    for bien in bienes:
        texto = reemplazar(str(bien).strip(), marcadores(datos))
        cantidad, detalle = "", texto
        partes = re.match(r"^(\d+)\s+(.*)$", texto)
        if partes:
            cantidad, detalle = partes.group(1), partes.group(2)
        filas.append([Paragraph(_escapar(detalle), estilos["vineta"]),
                      Paragraph(_escapar(cantidad), estilos["vineta"])])
    tabla = Table(filas, colWidths=[11.0 * cm, 3.0 * cm])
    tabla.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#b0b0b0")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eaeaea")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return tabla


def generar_acta(datos, ruta):
    """Genera el acta de entrega de bienes, equipos y vehículo."""
    valores = marcadores(datos)
    estilos = _estilos()
    historia = []
    for tipo, contenido in PL.ACTA_BLOQUES:
        texto = reemplazar(contenido, valores)
        if tipo == "titulo":
            historia.append(Paragraph(_escapar(texto), estilos["titulo"]))
        elif tipo == "subtitulo":
            historia.append(Paragraph(_escapar(texto), estilos["subtitulo"]))
        elif tipo == "subsubtitulo":
            historia.append(Paragraph(_escapar(texto), estilos["subsubtitulo"]))
        elif tipo == "vineta_vehiculo":
            historia.append(Paragraph(_escapar(texto), estilos["vineta_vehiculo"],
                                      bulletText="•"))
        elif tipo == "parrafo":
            # los párrafos del acta traen viñetas internas ("- ...")
            lineas = str(texto).split("\n")
            for linea in lineas:
                limpia = linea.strip()
                if not limpia:
                    continue
                if limpia.startswith("- "):
                    historia.append(Paragraph(_escapar(limpia[2:]),
                                              estilos["vineta"], bulletText="\u2022"))
                else:
                    historia.append(_parrafo_con_negritas(limpia, estilos["parrafo"], valores))
                    if limpia.lower().startswith("los siguientes bienes"):
                        historia.append(Spacer(1, 0.2 * cm))
                        historia.append(_tabla_bienes(datos, estilos))
                        historia.append(Spacer(1, 0.3 * cm))
        else:
            historia.append(_parrafo_con_negritas(texto, estilos["parrafo"], valores))

    # 🖊️ firmas
    historia.append(Spacer(1, 1.6 * cm))
    historia.append(_bloque_firmas(valores, estilos))

    return _crear_pdf(ruta, historia, (3.17, 3.17, 2.54, 2.54))
