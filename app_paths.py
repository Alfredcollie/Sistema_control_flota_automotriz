# -*- coding: utf-8 -*-
"""
APP_PATHS.PY - Rutas absolutas de configuración y recursos.
Funciona igual en desarrollo (python) y compilado (.exe).
"""
import os
import sys
from pathlib import Path


def _base_dir():
    """Carpeta base: la del .exe si está compilado, o la del script."""
    if getattr(sys, "frozen", False):
        return Path(os.path.dirname(sys.executable))
    return Path(__file__).resolve().parent


BASE_DIR = _base_dir()

# Archivo de configuración local de ESTE programa
CONFIG_FILE = BASE_DIR / "config_local.json"

# Carpeta para recursos (logos, etc.)
RESOURCES_DIR = BASE_DIR