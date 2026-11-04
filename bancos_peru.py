# -*- coding: utf-8 -*-
"""Bancos y entidades financieras del Perú (lista única).

La usan el módulo de Choferes (desplegable) y la Ficha PDF del chofer,
así ambos muestran exactamente lo mismo.
"""

BANCOS_PERU = [
    "Seleccione banco…",
    "BCP - Banco de Crédito del Perú",
    "BBVA Continental",
    "Interbank",
    "Scotiabank Perú",
    "Banco de la Nación",
    "BanBif",
    "Banco Pichincha",
    "Mibanco",
    "Banco Falabella",
    "Banco Ripley",
    "Banco GNB",
    "Banco Santander Perú",
    "Banco Itaú Perú",
    "Banco de Comercio",
    "Banco Cencosud",
    "Banco Pólen",
    "Caja Arequipa",
    "Caja Huancayo",
    "Caja Piura",
    "Caja Cusco",
    "Caja Trujillo",
    "Caja Sullana",
    "Caja Ica",
    "Caja Tacna",
    "Caja Paita",
    "Caja del Santa",
    "Caja Los Andes",
    "Cooperativa Abaco",
    "Cooperativa Pacífico",
    "Otro / Efectivo",
]

# Opciones sin la fila "Seleccione banco…" (para el PDF)
BANCOS_PDF = [b for b in BANCOS_PERU if not b.startswith("Seleccione")]
