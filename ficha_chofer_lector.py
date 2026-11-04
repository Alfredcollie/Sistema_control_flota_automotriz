# -*- coding: utf-8 -*-
"""Lectura de la FICHA DE REGISTRO DE CHOFERES en PDF.

Acepta dos casos:

  1. PDF con formulario interactivo (AcroForm) -> se leen los campos (lo normal).
  2. PDF SIN formulario (aplanado, impreso o convertido) -> se leen los textos
     por su POSICIÓN, con la misma geometría del generador oficial.

🔑 La geometría NO se copia aquí: se toma de 'crear_ficha_chofer_pdf.CAMPOS_FICHA',
   así si la ficha cambia, el lector cambia solo.
"""

import os

try:                                     # geometría oficial
    from crear_ficha_chofer_pdf import (CAMPOS_FICHA as _CAMPOS, ALTO_CARTA as _ALTO,
                                        ETIQUETAS_FICHA as _ETIQUETAS)
except Exception:                        # respaldo (no debería usarse)
    _CAMPOS = []
    _ALTO = 792.0
    _ETIQUETAS = []

CAMPOS = [(campo[0], campo[3], campo[2], campo[4]) for campo in _CAMPOS]   # nombre, x, y, ancho
# Etiquetas dibujadas: (x_etiqueta, y, texto) — se usan para NO confundirlas con datos
ETIQUETAS = []
for _campo in _CAMPOS:
    _nombre, _etiqueta, _y, _x, _ancho = _campo[0], _campo[1], _campo[2], _campo[3], _campo[4]
    ETIQUETAS.append((40 if _x <= 205 else _x - 100, _y, str(_etiqueta)))
ALTO_CARTA = _ALTO
CAMPOS_IMPORTANTES = ("dni", "nombres", "licencia")

# 🗂️ Geometría del formato ANTERIOR de la ficha (antes de agregar los datos
#    bancarios). Sirve para poder leer fichas que ya estaban impresas/llenadas.
CAMPOS_FORMATO_ANTERIOR = [
    ("dni", 205, 636, 375), ("nombres", 205, 612, 375), ("ruc", 205, 588, 375),
    ("direccion", 205, 564, 375), ("fecha_nacimiento", 205, 540, 375),
    ("sexo", 205, 516, 375), ("numero_hijos", 205, 492, 375),
    ("telefono", 205, 468, 375), ("telefono_emergencia", 205, 444, 375),
    ("contacto_emergencia_nombre", 205, 420, 375),
    ("talla_polo", 205, 396, 180), ("talla_casaca", 480, 396, 100),
    ("correo", 205, 372, 375), ("estado_laboral", 205, 348, 375),
    ("licencia", 205, 290, 375), ("categoria_licencia", 205, 266, 375),
    ("venc_licencia", 205, 242, 375), ("inicio_contrato", 205, 184, 375),
    ("fin_contrato", 205, 160, 375), ("carnet_sanidad", 205, 102, 145),
    ("sanidad_emision", 470, 102, 115), ("venc_sanidad", 205, 78, 375),
]


def _leer_formulario(ruta):
    """Valores del formulario interactivo ({} si el PDF no tiene campos)."""
    try:
        from pypdf import PdfReader
    except Exception:
        return {}
    try:
        campos = PdfReader(ruta).get_fields() or {}
    except Exception:
        return {}
    valores = {}
    for nombre, datos in campos.items():
        try:
            valor = str((datos or {}).get("/V", "") or "").strip()
        except Exception:
            valor = ""
        if valor and valor != " ":
            valores[nombre] = valor
    return valores


def _leer_por_posicion(ruta, campos=None):
    """Valores leídos del TEXTO del PDF, ubicando cada dato en su casilla."""
    try:
        import fitz
    except Exception:
        return {}
    try:
        documento = fitz.open(ruta)
    except Exception:
        return {}
    try:
        if documento.page_count == 0:
            return {}
        pagina = documento[0]
        alto_pagina = float(pagina.rect.height) or ALTO_CARTA
        palabras = pagina.get_text("words") or []
    finally:
        documento.close()

    tabla = campos if campos is not None else CAMPOS
    # cajas de las etiquetas del formato que se está leyendo (para descartarlas)
    cajas_etiqueta = []
    for nombre_campo, x_campo_e, y_etiqueta_e, ancho_e in tabla:
        texto_etiqueta = ""
        for campo_oficial in _CAMPOS:
            if campo_oficial[0] == nombre_campo:
                texto_etiqueta = str(campo_oficial[1])
                break
        if texto_etiqueta:
            x_etiqueta = 40 if x_campo_e <= 205 else x_campo_e - 100
            ancho_real = None
            for x_oficial, y_oficial, ancho_oficial in _ETIQUETAS:
                if abs(y_oficial - y_etiqueta_e) < 0.5 and abs(x_oficial - x_etiqueta) < 0.5:
                    ancho_real = ancho_oficial
                    break
            cajas_etiqueta.append((x_etiqueta, y_etiqueta_e,
                                   ancho_real if ancho_real else len(texto_etiqueta) * 5.3))
    valores = {}
    for nombre, x_campo, y_etiqueta, ancho in tabla:
        izquierda = x_campo
        derecha = x_campo + ancho
        # la ficha se dibuja con origen abajo-izquierda; fitz usa arriba-izquierda
        arriba = alto_pagina - (y_etiqueta + 11)
        abajo = alto_pagina - (y_etiqueta - 5)
        def _es_etiqueta(w):
            """¿La palabra es parte del rótulo de algún campo de esta línea?"""
            for x_etiq, y_etiq, largo in cajas_etiqueta:
                if abs(y_etiq - y_etiqueta) > 9:
                    continue
                fin = x_etiq + largo + 4                # ancho del rótulo + margen
                if w[0] >= x_etiq - 3 and w[2] <= fin:
                    return True
            return False

        dentro = [w for w in palabras
                  if w[0] >= izquierda - 8 and w[2] <= derecha + 40
                  and w[1] >= arriba - 5 and w[3] <= abajo + 7
                  and not _es_etiqueta(w)]
        dentro.sort(key=lambda w: (round(w[1], 1), w[0]))
        texto = " ".join(str(w[4]) for w in dentro).strip()
        if texto:
            valores[nombre] = texto
    return valores


def leer_ficha(ruta):
    """Devuelve (valores, origen, aviso).

    origen: "formulario" (tenía campos) o "texto" (se recuperó por posición).
    aviso : texto para mostrar al usuario ('' si todo normal).
    """
    if not ruta or not os.path.isfile(ruta):
        return {}, "", "No se encontró el archivo PDF."

    valores = _leer_formulario(ruta)
    if valores:
        return valores, "formulario", ""

    valores = _leer_por_posicion(ruta)
    formato = "actual"
    # 🗂️ También se prueba el formato ANTERIOR (fichas ya impresas): gana el que
    #    reconozca más datos.
    alternos = _leer_por_posicion(ruta, CAMPOS_FORMATO_ANTERIOR)
    if len(alternos) > len(valores):
        valores = alternos
        formato = "anterior"
    if not valores:
        return {}, "", ("El PDF no tiene formulario y tampoco se pudo leer texto.\n\n"
                        "Si es una FOTO o un escaneo, hay que escribir los datos a mano "
                        "(o pasar el PDF por un OCR).")
    aviso = ("El PDF venía SIN formulario interactivo (aplanado o convertido).\n\n"
             "Se recuperaron %d dato(s) leyéndolos por su posición en la ficha "
             "(formato %s).\nRevisá el formulario antes de guardar."
             % (len(valores), formato))
    faltan = [n for n in CAMPOS_IMPORTANTES if not valores.get(n)]
    if faltan:
        aviso += "\n\nFaltan datos clave: " + ", ".join(faltan)
    return valores, "texto", aviso
