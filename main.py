"""
Sistema de Reportes de Consumo ISP
Punto de entrada principal de la aplicación

Desarrollado para el Departamento de Telecomunicaciones
Universidad Metropolitana (UNIMET)

Elaborado por: Yosmary Lopez
Revisado por: Gerente de Redes
Fecha: 18/02/2026
"""

from interfaz import AplicacionSRCT


def main():
    """
    Función principal que inicia la aplicación.
    """
    app = AplicacionSRCT()
    app.ejecutar()


if __name__ == "__main__":
    main()