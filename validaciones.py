"""
Sistema de Reportes de Consumo ISP
Módulo de Validaciones
Implementa RF-04 del SRS
"""

from constantes import PROVEEDORES, CONTRATADO


def validar_datos(valores_texto):
    """
    Valida todos los campos de consumo ingresados por el usuario.
    
    Orden de validación según el SRS:
    1. Campo completo (no vacío)
    2. Valor numérico
    3. Valor no negativo (>= 0)
    4. Valor dentro del límite contractual (<= contratado)
    
    Parámetros:
        valores_texto (list): Lista de 5 strings ingresados por el usuario
    
    Retorna:
        tuple: (exito: bool, mensaje: str, valores: list)
            - Si es válido: (True, "OK", [lista de floats])
            - Si hay error: (False, "mensaje de error", [])
    """
    valores = []

    for i in range(len(PROVEEDORES)):
        nombre = PROVEEDORES[i]
        valor = valores_texto[i]

        # Validación 1: Campo completo
        if valor.strip() == "":
            return (
                False,
                f"Error: El campo de consumo de {nombre} se encuentra vacío.\n"
                f"Por favor ingrese un valor.",
                []
            )

        # Validación 2: Valor numérico
        try:
            numero = float(valor)
        except ValueError:
            return (
                False,
                f"Error: El valor ingresado en {nombre} no es numérico.\n"
                f"Por favor ingrese solo números.",
                []
            )

        # Validación 3: Valor no negativo
        if numero < 0:
            return (
                False,
                f"Error: El valor de consumo de {nombre} no puede ser negativo.",
                []
            )

        # Validación 4: No exceder contratado
        if numero > CONTRATADO[i]:
            return (
                False,
                f"Error: El consumo de {nombre} ({numero} Mbps) excede\n"
                f"el máximo contratado de {CONTRATADO[i]} Mbps.",
                []
            )

        valores.append(numero)

    return (True, "OK", valores)