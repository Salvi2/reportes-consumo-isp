"""
Sistema de Reportes de Consumo ISP
Módulo de Constantes
"""

import os
import sys
from dotenv import load_dotenv


def obtener_ruta_env():
    # Buscar primero junto al .exe (por si el usuario lo puso ahí)
    if getattr(sys, 'frozen', False):
        ruta_exe = os.path.join(os.path.dirname(sys.executable), '.env')
        ruta_interna = os.path.join(sys._MEIPASS, '.env')

        if os.path.exists(ruta_exe):
            return ruta_exe
        elif os.path.exists(ruta_interna):
            return ruta_interna
    else:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')

    return '.env'


load_dotenv(obtener_ruta_env())

PROVEEDORES = [
    os.getenv("PROVEEDOR_1", "Proveedor 1"),
    os.getenv("PROVEEDOR_2", "Proveedor 2"),
    os.getenv("PROVEEDOR_3", "Proveedor 3"),
    os.getenv("PROVEEDOR_4", "Proveedor 4"),
    os.getenv("PROVEEDOR_5", "Proveedor 5"),
]

CONTRATADO = [
    int(os.getenv("CONTRATADO_1", 1000)),
    int(os.getenv("CONTRATADO_2", 1000)),
    int(os.getenv("CONTRATADO_3", 1000)),
    int(os.getenv("CONTRATADO_4", 700)),
    int(os.getenv("CONTRATADO_5", 300)),
]

TOTAL_CONTRATADO = sum(CONTRATADO)

NOMBRE_ARCHIVO = os.getenv("NOMBRE_ARCHIVO", "Consumo_ISP")

MESES = [
    "Enero", "Febrero", "Marzo", "Abril",
    "Mayo", "Junio", "Julio", "Agosto",
    "Septiembre", "Octubre", "Noviembre", "Diciembre"
]

UMBRAL_VERDE = int(os.getenv("UMBRAL_VERDE", 65))
UMBRAL_AMARILLO = int(os.getenv("UMBRAL_AMARILLO", 85))

COLOR_VERDE = "00B050"
COLOR_AMARILLO = "FFFF00"
COLOR_ROJO = "FF0000"
COLOR_BLANCO = "FFFFFF"
COLOR_NEGRO = "000000"

COLOR_PRINCIPAL = "#FD8204"
COLOR_SECUNDARIO = "#033087"
COLOR_PRINCIPAL_HEX = "FD8204"
COLOR_SECUNDARIO_HEX = "033087"

COLOR_ENCABEZADO = "033087"
COLOR_FILA_TOTAL = "FDE8D0"