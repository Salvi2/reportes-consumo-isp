"""
Sistema de Reportes de Consumo ISP
Módulo de Historial
"""

import json
import os


ARCHIVO_HISTORIAL = "historial.json"


def cargar_historial():
    if os.path.exists(ARCHIVO_HISTORIAL):
        with open(ARCHIVO_HISTORIAL, "r") as archivo:
            return json.load(archivo)
    return {}


def mes_existe(anio, mes):
    """
    Verifica si ya existe un reporte para ese mes y año.
    """
    historial = cargar_historial()

    if anio in historial and mes in historial[anio]:
        return True
    return False


def guardar_mes(anio, mes, consumos, porcentajes, total_consumo, porcentaje_general):
    historial = cargar_historial()

    if anio not in historial:
        historial[anio] = {}

    historial[anio][mes] = {
        "consumos": consumos,
        "porcentajes": porcentajes,
        "total_consumo": total_consumo,
        "porcentaje_general": porcentaje_general
    }

    with open(ARCHIVO_HISTORIAL, "w") as archivo:
        json.dump(historial, archivo, indent=4)


def obtener_datos_anio(anio):
    historial = cargar_historial()

    if anio in historial:
        return historial[anio]
    return None