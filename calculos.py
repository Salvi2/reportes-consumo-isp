"""
Sistema de Reportes de Consumo ISP
Módulo de Cálculos
Implementa RF-08, RF-09 y RF-10 del SRS
"""

from constantes import CONTRATADO, TOTAL_CONTRATADO


def calcular_porcentajes(consumos):
    """
    Calcula el porcentaje de uso por proveedor.
    Fórmula: (Consumo / Contratado) × 100
    Redondeado a 2 decimales.
    
    Parámetros:
        consumos (list): Lista de 5 valores de consumo en Mbps
    
    Retorna:
        list: Lista de 5 porcentajes redondeados a 2 decimales
    """
    porcentajes = []

    for i in range(len(consumos)):
        porcentaje = (consumos[i] / CONTRATADO[i]) * 100
        porcentaje = round(porcentaje, 2)
        porcentajes.append(porcentaje)

    return porcentajes


def calcular_total_consumo(consumos):
    """
    Calcula la sumatoria total de consumo.
    Fórmula: Total = Σ consumos
    
    Parámetros:
        consumos (list): Lista de 5 valores de consumo en Mbps
    
    Retorna:
        float: Suma total de consumos
    """
    total = 0

    for consumo in consumos:
        total = total + consumo

    return round(total, 2)


def calcular_porcentaje_general(total_consumo):
    """
    Calcula el porcentaje general de uso.
    Fórmula: (Total Consumo / 4000) × 100
    Redondeado a 2 decimales.
    
    Parámetros:
        total_consumo (float): Suma total de consumos
    
    Retorna:
        float: Porcentaje general redondeado a 2 decimales
    """
    porcentaje = (total_consumo / TOTAL_CONTRATADO) * 100
    return round(porcentaje, 2)