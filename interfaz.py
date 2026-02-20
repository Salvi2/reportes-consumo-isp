"""
Sistema de Reportes de Consumo ISP
Módulo de Interfaz Gráfica
Diseño moderno y minimalista
"""

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import os
import sys

from constantes import PROVEEDORES, CONTRATADO, MESES
from validaciones import validar_datos
from calculos import (
    calcular_porcentajes,
    calcular_total_consumo,
    calcular_porcentaje_general
)
from generador_excel import generar_reporte


def obtener_ruta(nombre_archivo):
    if getattr(sys, 'frozen', False):
        ruta = os.path.join(sys._MEIPASS, nombre_archivo)
    else:
        ruta = nombre_archivo
    return ruta


class AplicacionSRCT:

    def __init__(self):

        # ═══════════════════════════════════════
        # COLORES
        # ═══════════════════════════════════════
        self.NARANJA = "#FD8204"
        self.AZUL = "#033087"
        self.BLANCO = "#FFFFFF"
        self.GRIS_CLARO = "#F7F7F7"
        self.GRIS_TEXTO = "#333333"
        self.GRIS_SECUNDARIO = "#888888"
        self.GRIS_BORDE = "#E0E0E0"
        self.VERDE_EXITO = "#28A745"
        self.ROJO_ERROR = "#DC3545"

        # ═══════════════════════════════════════
        # VENTANA PRINCIPAL
        # ═══════════════════════════════════════
        self.ventana = tk.Tk()
        self.ventana.title("Reportes de Consumo ISP")
        self.ventana.geometry("520x560")
        self.ventana.resizable(False, False)
        self.ventana.configure(bg=self.BLANCO)

        # Tema moderno
        self.estilo = ttk.Style()
        self.estilo.theme_use("clam")

        # Configurar estilos del combobox
        self.estilo.configure(
            "TCombobox",
            fieldbackground=self.BLANCO,
            background=self.BLANCO
        )

        # Icono
        try:
            self.ventana.iconbitmap(obtener_ruta("image.ico"))
        except:
            pass

        self.centrar_ventana()

        # ═══════════════════════════════════════
        # ACENTO SUPERIOR (línea naranja delgada)
        # ═══════════════════════════════════════
        tk.Frame(
            self.ventana, bg=self.NARANJA, height=4
        ).pack(fill="x")

        # ═══════════════════════════════════════
        # ENCABEZADO
        # ═══════════════════════════════════════
        frame_header = tk.Frame(self.ventana, bg=self.BLANCO, pady=20)
        frame_header.pack(fill="x")

        tk.Label(
            frame_header,
            text="Reportes de Consumo ISP",
            font=("Segoe UI", 20, "bold"),
            fg=self.AZUL,
            bg=self.BLANCO
        ).pack()

        tk.Label(
            frame_header,
            text="CETIC — UNIMET",
            font=("Segoe UI", 9),
            fg=self.GRIS_SECUNDARIO,
            bg=self.BLANCO
        ).pack()

        # Separador
        tk.Frame(
            self.ventana, bg=self.GRIS_BORDE, height=1
        ).pack(fill="x", padx=30)

        # ═══════════════════════════════════════
        # PERÍODO
        # ═══════════════════════════════════════
        frame_periodo = tk.Frame(self.ventana, bg=self.BLANCO, pady=15)
        frame_periodo.pack(fill="x", padx=30)

        tk.Label(
            frame_periodo,
            text="Periodo",
            font=("Segoe UI", 10, "bold"),
            fg=self.GRIS_TEXTO,
            bg=self.BLANCO
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 8))

        # Mes
        tk.Label(
            frame_periodo,
            text="Mes",
            font=("Segoe UI", 9),
            fg=self.GRIS_SECUNDARIO,
            bg=self.BLANCO
        ).grid(row=1, column=0, sticky="w", padx=(0, 8))

        mes_actual = datetime.now().month - 1
        self.combo_mes = ttk.Combobox(
            frame_periodo,
            values=MESES,
            state="readonly",
            width=14,
            font=("Segoe UI", 10)
        )
        self.combo_mes.current(mes_actual)
        self.combo_mes.grid(row=1, column=1, padx=(0, 25))

        # Año
        tk.Label(
            frame_periodo,
            text="Año",
            font=("Segoe UI", 9),
            fg=self.GRIS_SECUNDARIO,
            bg=self.BLANCO
        ).grid(row=1, column=2, sticky="w", padx=(0, 8))

        anio_actual = str(datetime.now().year)
        self.entry_anio = tk.Entry(
            frame_periodo,
            width=7,
            font=("Segoe UI", 10),
            justify="center",
            relief="solid",
            bd=1,
            highlightthickness=0
        )
        self.entry_anio.insert(0, anio_actual)
        self.entry_anio.grid(row=1, column=3)

        # ═══════════════════════════════════════
        # TABLA DE PROVEEDORES
        # ═══════════════════════════════════════
        frame_tabla = tk.Frame(self.ventana, bg=self.BLANCO, pady=10)
        frame_tabla.pack(fill="x", padx=30)

        tk.Label(
            frame_tabla,
            text="Consumo por proveedor",
            font=("Segoe UI", 10, "bold"),
            fg=self.GRIS_TEXTO,
            bg=self.BLANCO
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

        # Encabezados
        headers = [
            ("Proveedor", 0, "w"),
            ("Contratado", 1, ""),
            ("Consumo (Mbps)", 2, "")
        ]

        for texto, col, anchor in headers:
            tk.Label(
                frame_tabla,
                text=texto,
                font=("Segoe UI", 9),
                fg=self.GRIS_SECUNDARIO,
                bg=self.BLANCO
            ).grid(row=1, column=col, sticky=anchor, padx=(0, 15), pady=(0, 4))

        # Separador bajo encabezados
        separador_header = tk.Frame(frame_tabla, bg=self.GRIS_BORDE, height=1)
        separador_header.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 5))

        # Campos
        self.entries_consumo = []

        for i in range(len(PROVEEDORES)):
            fila = i + 3

            # Fondo alternado
            color_fila = self.BLANCO if i % 2 == 0 else self.GRIS_CLARO

            # Nombre
            tk.Label(
                frame_tabla,
                text=PROVEEDORES[i],
                font=("Segoe UI", 10),
                fg=self.GRIS_TEXTO,
                bg=self.BLANCO
            ).grid(row=fila, column=0, sticky="w", pady=4, padx=(0, 15))

            # Contratado
            tk.Label(
                frame_tabla,
                text=f"{CONTRATADO[i]:,} Mbps",
                font=("Segoe UI", 10),
                fg=self.GRIS_SECUNDARIO,
                bg=self.BLANCO
            ).grid(row=fila, column=1, pady=4, padx=(0, 15))

            # Input
            entry = tk.Entry(
                frame_tabla,
                width=14,
                font=("Segoe UI", 10),
                justify="center",
                relief="solid",
                bd=1,
                highlightcolor=self.NARANJA,
                highlightthickness=1
            )
            entry.grid(row=fila, column=2, pady=4)
            self.entries_consumo.append(entry)

        # Separador bajo tabla
        tk.Frame(
            self.ventana, bg=self.GRIS_BORDE, height=1
        ).pack(fill="x", padx=30, pady=(5, 0))

        # ═══════════════════════════════════════
        # BOTONES
        # ═══════════════════════════════════════
        frame_botones = tk.Frame(self.ventana, bg=self.BLANCO, pady=20)
        frame_botones.pack(fill="x", padx=30)

        # Botón Generar (naranja)
        self.boton_generar = tk.Button(
            frame_botones,
            text="Generar Reporte",
            font=("Segoe UI", 11, "bold"),
            bg=self.NARANJA,
            fg=self.BLANCO,
            activebackground="#E57303",
            activeforeground=self.BLANCO,
            cursor="hand2",
            pady=10,
            relief="flat",
            bd=0,
            command=self.generar_reporte
        )
        self.boton_generar.pack(fill="x", pady=(0, 8))

        # Hover generar
        self.boton_generar.bind(
            "<Enter>", lambda e: self.boton_generar.config(bg="#E57303")
        )
        self.boton_generar.bind(
            "<Leave>", lambda e: self.boton_generar.config(bg=self.NARANJA)
        )

        # Botón Limpiar (borde, sin relleno)
        self.boton_limpiar = tk.Button(
            frame_botones,
            text="Limpiar campos",
            font=("Segoe UI", 10),
            bg=self.BLANCO,
            fg=self.GRIS_SECUNDARIO,
            activebackground=self.GRIS_CLARO,
            activeforeground=self.GRIS_TEXTO,
            cursor="hand2",
            pady=6,
            relief="solid",
            bd=1,
            command=self.limpiar_campos
        )
        self.boton_limpiar.pack(fill="x")

        # Hover limpiar
        self.boton_limpiar.bind(
            "<Enter>", lambda e: self.boton_limpiar.config(
                bg=self.GRIS_CLARO, fg=self.GRIS_TEXTO
            )
        )
        self.boton_limpiar.bind(
            "<Leave>", lambda e: self.boton_limpiar.config(
                bg=self.BLANCO, fg=self.GRIS_SECUNDARIO
            )
        )

        # ═══════════════════════════════════════
        # BARRA DE ESTADO (minimalista)
        # ═══════════════════════════════════════
        frame_estado = tk.Frame(self.ventana, bg=self.GRIS_CLARO)
        frame_estado.pack(fill="x", side="bottom")

        self.label_estado = tk.Label(
            frame_estado,
            text="Listo",
            font=("Segoe UI", 9),
            bg=self.GRIS_CLARO,
            fg=self.GRIS_SECUNDARIO,
            anchor="w",
            padx=15,
            pady=6
        )
        self.label_estado.pack(fill="x")

    def centrar_ventana(self):
        self.ventana.update_idletasks()
        ancho = 520
        alto = 560
        x = (self.ventana.winfo_screenwidth() // 2) - (ancho // 2)
        y = (self.ventana.winfo_screenheight() // 2) - (alto // 2)
        self.ventana.geometry(f"{ancho}x{alto}+{x}+{y}")

    def generar_reporte(self):

        self.label_estado.config(text="Validando...", fg=self.GRIS_SECUNDARIO)
        self.ventana.update()

        # Obtener valores
        valores_texto = []
        for entry in self.entries_consumo:
            valores_texto.append(entry.get())

        mes = self.combo_mes.get()
        anio = self.entry_anio.get()

        if mes == "":
            messagebox.showerror("Error", "Seleccione un mes.")
            self.label_estado.config(text="Error en los datos", fg=self.ROJO_ERROR)
            return

        if anio.strip() == "" or not anio.isdigit():
            messagebox.showerror("Error", "Ingrese un año valido.")
            self.label_estado.config(text="Error en los datos", fg=self.ROJO_ERROR)
            return

        # Validar
        exito, mensaje, consumos = validar_datos(valores_texto)

        if not exito:
            messagebox.showerror("Error de validacion", mensaje)
            self.label_estado.config(text="Error en los datos", fg=self.ROJO_ERROR)
            return

        # Calcular
        self.label_estado.config(text="Generando...", fg=self.GRIS_SECUNDARIO)
        self.ventana.update()

        porcentajes = calcular_porcentajes(consumos)
        total_consumo = calcular_total_consumo(consumos)
        porcentaje_general = calcular_porcentaje_general(total_consumo)

        # Generar Excel
        try:
            nombre_archivo = generar_reporte(
                mes, anio,
                consumos, porcentajes,
                total_consumo, porcentaje_general
            )

            ruta_completa = os.path.abspath(nombre_archivo)

            messagebox.showinfo(
                "Reporte generado",
                f"Archivo: {nombre_archivo}\n"
                f"Ubicacion: {ruta_completa}\n\n"
                f"Consumo total: {total_consumo:,.2f} Mbps\n"
                f"Uso general: {porcentaje_general}%"
            )

            self.label_estado.config(
                text=f"Generado: {nombre_archivo}",
                fg=self.VERDE_EXITO
            )
            # Abrir el Excel automáticamente
            os.startfile(ruta_completa)

        except PermissionError:
            messagebox.showerror(
                "Error",
                "No se pudo guardar el archivo.\n"
                "Verifique que no este abierto en Excel."
            )
            self.label_estado.config(text="Error al guardar", fg=self.ROJO_ERROR)

        except Exception as e:
            messagebox.showerror("Error", f"Error inesperado:\n{str(e)}")
            self.label_estado.config(text="Error inesperado", fg=self.ROJO_ERROR)

    def limpiar_campos(self):
        for entry in self.entries_consumo:
            entry.delete(0, tk.END)
        self.label_estado.config(text="Listo", fg=self.GRIS_SECUNDARIO)

    def ejecutar(self):
        self.ventana.mainloop()