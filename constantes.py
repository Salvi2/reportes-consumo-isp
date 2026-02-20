"""
Sistema de Reportes de Consumo ISP
Módulo de Constantes
"""

import os
from dotenv import load_dotenv

# Cargar el archivo .env
load_dotenv()

# Proveedores (lee del .env)
PROVEEDORES = [
    os.getenv("PROVEEDOR_1", "Proveedor 1"),
    os.getenv("PROVEEDOR_2", "Proveedor 2"),
    os.getenv("PROVEEDOR_3", "Proveedor 3"),
    os.getenv("PROVEEDOR_4", "Proveedor 4"),
    os.getenv("PROVEEDOR_5", "Proveedor 5"),
]

# Valores contratados (lee del .env)
CONTRATADO = [
    int(os.getenv("CONTRATADO_1", 1000)),
    int(os.getenv("CONTRATADO_2", 1000)),
    int(os.getenv("CONTRATADO_3", 1000)),
    int(os.getenv("CONTRATADO_4", 700)),
    int(os.getenv("CONTRATADO_5", 300)),
]

TOTAL_CONTRATADO = sum(CONTRATADO)

# Nombre del archivo de salida
NOMBRE_ARCHIVO = os.getenv("NOMBRE_ARCHIVO", "Reporte_ISP")

# Meses
MESES = [
    "Enero", "Febrero", "Marzo", "Abril",
    "Mayo", "Junio", "Julio", "Agosto",
    "Septiembre", "Octubre", "Noviembre", "Diciembre"
]

# Umbrales del semáforo (lee del .env)
UMBRAL_VERDE = int(os.getenv("UMBRAL_VERDE", 65))
UMBRAL_AMARILLO = int(os.getenv("UMBRAL_AMARILLO", 85))

# Colores formato condicional (fijos, no sensibles)
COLOR_VERDE = "00B050"
COLOR_AMARILLO = "FFFF00"
COLOR_ROJO = "FF0000"
COLOR_BLANCO = "FFFFFF"
COLOR_NEGRO = "000000"

# Colores institucionales (fijos, no sensibles)
COLOR_PRINCIPAL = "#FD8204"
COLOR_SECUNDARIO = "#033087"
COLOR_PRINCIPAL_HEX = "FD8204"
COLOR_SECUNDARIO_HEX = "033087"

# Colores Excel
COLOR_ENCABEZADO = "033087"
COLOR_FILA_TOTAL = "FDE8D0"