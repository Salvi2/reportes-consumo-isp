"""
Sistema de Reportes de Consumo ISP
Módulo de Indicador de Ancho de Banda (PDF)
"""

from reportlab.lib.pagesizes import letter
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle,
    Paragraph, Spacer
)
from reportlab.graphics.shapes import Drawing, String
from reportlab.graphics.charts.piecharts import Pie

from constantes import (
    NOMBRE_ARCHIVO, TOTAL_CONTRATADO,
    UMBRAL_VERDE, UMBRAL_AMARILLO
)


def generar_indicador(mes, anio, total_consumo, porcentaje_general):

    nombre_archivo = f"{NOMBRE_ARCHIVO}_Indicador_{mes}_{anio}.pdf"

    doc = SimpleDocTemplate(
        nombre_archivo,
        pagesize=letter,
        rightMargin=50,
        leftMargin=50,
        topMargin=40,
        bottomMargin=40
    )

    NARANJA = HexColor("#FD8204")
    AZUL = HexColor("#033087")
    VERDE = HexColor("#00B050")
    AMARILLO = HexColor("#FFD700")
    ROJO = HexColor("#FF0000")
    GRIS = HexColor("#E0E0E0")

    elementos = []
    estilos = getSampleStyleSheet()

    # ═══════════════════════════════════════
    # ENCABEZADO
    # ═══════════════════════════════════════
    estilo_titulo = ParagraphStyle(
        'Titulo',
        parent=estilos['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=20,
        textColor=AZUL,
        alignment=1,
        spaceAfter=5
    )

    estilo_subtitulo = ParagraphStyle(
        'Subtitulo',
        parent=estilos['Normal'],
        fontName='Helvetica',
        fontSize=10,
        textColor=HexColor("#888888"),
        alignment=1,
        spaceAfter=5
    )

    estilo_anio = ParagraphStyle(
        'Anio',
        parent=estilos['Normal'],
        fontName='Helvetica-Bold',
        fontSize=14,
        textColor=NARANJA,
        alignment=1,
        spaceAfter=25
    )

    elementos.append(Paragraph("Indicador de Ancho de Banda", estilo_titulo))
    elementos.append(Paragraph(
        "CETIC — UNIMET",
        estilo_subtitulo
    ))
    elementos.append(Spacer(1, 5))
    elementos.append(Paragraph(f"Año {anio}", estilo_anio))

    # ═══════════════════════════════════════
    # TABLA (una sola fila)
    # ═══════════════════════════════════════
    disponible = TOTAL_CONTRATADO - total_consumo
    disponible = round(disponible, 2)

    datos_tabla = [
        [
            "Mes",
            "Total Consumo\n(Mbps)",
            "Total Contratado\n(Mbps)",
            "Promedio\nUsado"
        ],
        [
            mes,
            f"{total_consumo:,.2f}",
            f"{TOTAL_CONTRATADO:,}",
            f"{porcentaje_general:.2f}%"
        ]
    ]

    tabla = Table(
        datos_tabla,
        colWidths=[100, 130, 130, 110]
    )

    # Determinar color del porcentaje
    if porcentaje_general < UMBRAL_VERDE:
        color_porcentaje = VERDE
    elif porcentaje_general <= UMBRAL_AMARILLO:
        color_porcentaje = AMARILLO
    else:
        color_porcentaje = ROJO

    texto_porcentaje = white if porcentaje_general > UMBRAL_AMARILLO else black

    estilo_tabla = TableStyle([
        # Encabezado
        ('BACKGROUND', (0, 0), (-1, 0), AZUL),
        ('TEXTCOLOR', (0, 0), (-1, 0), white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),

        # Datos
        ('FONTNAME', (0, 1), (-1, 1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, 1), 12),
        ('ALIGN', (0, 1), (-1, 1), 'CENTER'),
        ('VALIGN', (0, 1), (-1, 1), 'MIDDLE'),

        # Mes en negrita
        ('FONTNAME', (0, 1), (0, 1), 'Helvetica-Bold'),

        # Color del porcentaje
        ('BACKGROUND', (3, 1), (3, 1), color_porcentaje),
        ('TEXTCOLOR', (3, 1), (3, 1), texto_porcentaje),
        ('FONTNAME', (3, 1), (3, 1), 'Helvetica-Bold'),
        ('FONTSIZE', (3, 1), (3, 1), 14),

        # Bordes
        ('GRID', (0, 0), (-1, -1), 0.5, HexColor("#D0D0D0")),

        # Padding
        ('TOPPADDING', (0, 0), (-1, -1), 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
    ])

    tabla.setStyle(estilo_tabla)
    elementos.append(tabla)
    elementos.append(Spacer(1, 40))

    # ═══════════════════════════════════════
    # GRÁFICO DE TORTA
    # ═══════════════════════════════════════
    estilo_grafico_titulo = ParagraphStyle(
        'GraficoTitulo',
        parent=estilos['Normal'],
        fontName='Helvetica-Bold',
        fontSize=12,
        textColor=AZUL,
        alignment=1,
        spaceAfter=15
    )

    elementos.append(Paragraph(
        f"Distribución de Ancho de Banda — {mes} {anio}",
        estilo_grafico_titulo
    ))

    drawing = Drawing(500, 250)

    torta = Pie()
    torta.x = 125
    torta.y = 25
    torta.width = 200
    torta.height = 200

    torta.data = [total_consumo, disponible]
    torta.labels = [
        f"Consumido{porcentaje_general:.2f}%",
       f"Contratado:{100 - porcentaje_general:.2f}%"
    ]

    # Colores de la torta
    torta.slices[0].fillColor = NARANJA
    torta.slices[0].strokeColor = white
    torta.slices[0].strokeWidth = 2

    torta.slices[1].fillColor = AZUL
    torta.slices[1].strokeColor = white
    torta.slices[1].strokeWidth = 2

    # Estilo de las etiquetas
    torta.slices[0].labelRadius = 1.3
    torta.slices[1].labelRadius = 1.3
    torta.slices[0].fontName = 'Helvetica-Bold'
    torta.slices[0].fontSize = 9
    torta.slices[1].fontName = 'Helvetica-Bold'
    torta.slices[1].fontSize = 9

    drawing.add(torta)
    elementos.append(drawing)
    elementos.append(Spacer(1, 20))

    # Leyenda
    estilo_leyenda = ParagraphStyle(
        'Leyenda',
        parent=estilos['Normal'],
        fontName='Helvetica',
        fontSize=9,
        textColor=HexColor("#666666"),
        alignment=1
    )

    elementos.append(Paragraph(
        '<font color="#FD8204">■</font> Consumido&nbsp;&nbsp;&nbsp;&nbsp;'
        '<font color="#033087">■</font> Contratado,
        estilo_leyenda
    ))

    # ═══════════════════════════════════════
    # SEMÁFORO
    # ═══════════════════════════════════════
    elementos.append(Spacer(1, 25))

    if porcentaje_general < UMBRAL_VERDE:
        estado = "Consumo normal (<65%)"
        color_estado = "#00B050"
    elif porcentaje_general <= UMBRAL_AMARILLO:
        estado = "PRECAUCION — Consumo moderado (65%-80%)"
        color_estado = "#FFD700"
    else:
        estado = "ALERTA — Consumo alto (>80%)"
        color_estado = "#FF0000"

    estilo_estado = ParagraphStyle(
        'Estado',
        parent=estilos['Normal'],
        fontName='Helvetica-Bold',
        fontSize=12,
        textColor=HexColor(color_estado),
        alignment=1,
        spaceAfter=30
    )

    elementos.append(Paragraph(f"Estado: {estado}", estilo_estado))

    # ═══════════════════════════════════════
    # PIE DE PÁGINA
    # ═══════════════════════════════════════
    estilo_pie = ParagraphStyle(
        'Pie',
        parent=estilos['Normal'],
        fontName='Helvetica',
        fontSize=8,
        textColor=HexColor("#AAAAAA"),
        alignment=1
    )

    doc.build(elementos)

    return nombre_archivo