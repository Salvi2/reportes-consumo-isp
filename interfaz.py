"""
Sistema de Reportes de Consumo ISP
Módulo de Interfaz Gráfica
"""
from tkinter import ttk, messagebox, simpledialog
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import os
import sys
from correo import enviar_correo
from indicador_pdf import generar_indicador
from constantes import SMTP_EMAIL, SMTP_PASS
from constantes import PROVEEDORES, CONTRATADO, MESES, TOTAL_CONTRATADO
from constantes import PROVEEDORES, CONTRATADO, MESES
from validaciones import validar_datos
from calculos import (
    calcular_porcentajes,
    calcular_total_consumo,
    calcular_porcentaje_general
)
from generador_excel import generar_reporte
from historial import guardar_mes, obtener_datos_anio, mes_existe
from reporte_anual import generar_reporte_anual


def obtener_ruta(nombre_archivo):
    if getattr(sys, 'frozen', False):
        ruta = os.path.join(sys._MEIPASS, nombre_archivo)
    else:
        ruta = nombre_archivo
    return ruta


class AplicacionSRCT:

    def __init__(self):

        self.NARANJA = "#FD8204"
        self.AZUL = "#033087"
        self.BLANCO = "#FFFFFF"
        self.GRIS_CLARO = "#F7F7F7"
        self.GRIS_TEXTO = "#333333"
        self.GRIS_SECUNDARIO = "#888888"
        self.GRIS_BORDE = "#E0E0E0"
        self.VERDE_EXITO = "#28A745"
        self.ROJO_ERROR = "#DC3545"

        self.ventana = tk.Tk()
        self.ventana.title("Reportes de Consumo ISP")
        self.ventana.geometry("520x720")
        self.ventana.resizable(False, False)
        self.ventana.configure(bg=self.BLANCO)

        self.estilo = ttk.Style()
        self.estilo.theme_use("clam")
        self.estilo.configure(
            "TCombobox",
            fieldbackground=self.BLANCO,
            background=self.BLANCO
        )

        try:
            self.ventana.iconbitmap(obtener_ruta("image.ico"))
        except:
            pass

        self.centrar_ventana()

        # Acento superior
        tk.Frame(self.ventana, bg=self.NARANJA, height=4).pack(fill="x")

        # Encabezado
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

        tk.Frame(self.ventana, bg=self.GRIS_BORDE, height=1).pack(fill="x", padx=30)

        # Periodo
        frame_periodo = tk.Frame(self.ventana, bg=self.BLANCO, pady=15)
        frame_periodo.pack(fill="x", padx=30)

        tk.Label(
            frame_periodo, text="Periodo",
            font=("Segoe UI", 10, "bold"),
            fg=self.GRIS_TEXTO, bg=self.BLANCO
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 8))

        tk.Label(
            frame_periodo, text="Mes",
            font=("Segoe UI", 9),
            fg=self.GRIS_SECUNDARIO, bg=self.BLANCO
        ).grid(row=1, column=0, sticky="w", padx=(0, 8))

        mes_actual = datetime.now().month - 1
        self.combo_mes = ttk.Combobox(
            frame_periodo, values=MESES,
            state="readonly", width=14,
            font=("Segoe UI", 10)
        )
        self.combo_mes.current(mes_actual)
        self.combo_mes.grid(row=1, column=1, padx=(0, 25))

        tk.Label(
            frame_periodo, text="Año",
            font=("Segoe UI", 9),
            fg=self.GRIS_SECUNDARIO, bg=self.BLANCO
        ).grid(row=1, column=2, sticky="w", padx=(0, 8))

        anio_actual = str(datetime.now().year)
        self.entry_anio = tk.Entry(
            frame_periodo, width=7,
            font=("Segoe UI", 10),
            justify="center", relief="solid",
            bd=1, highlightthickness=0
        )
        self.entry_anio.insert(0, anio_actual)
        self.entry_anio.grid(row=1, column=3)

        # Tabla de proveedores
        frame_tabla = tk.Frame(self.ventana, bg=self.BLANCO, pady=10)
        frame_tabla.pack(fill="x", padx=30)

        tk.Label(
            frame_tabla, text="Consumo por proveedor",
            font=("Segoe UI", 10, "bold"),
            fg=self.GRIS_TEXTO, bg=self.BLANCO
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

        headers = [
            ("Proveedor", 0, "w"),
            ("Contratado", 1, ""),
            ("Consumo (Mbps)", 2, "")
        ]

        for texto, col, anchor in headers:
            tk.Label(
                frame_tabla, text=texto,
                font=("Segoe UI", 9),
                fg=self.GRIS_SECUNDARIO, bg=self.BLANCO
            ).grid(row=1, column=col, sticky=anchor, padx=(0, 15), pady=(0, 4))

        tk.Frame(
            frame_tabla, bg=self.GRIS_BORDE, height=1
        ).grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 5))

        self.entries_consumo = []

        for i in range(len(PROVEEDORES)):
            fila = i + 3

            tk.Label(
                frame_tabla, text=PROVEEDORES[i],
                font=("Segoe UI", 10),
                fg=self.GRIS_TEXTO, bg=self.BLANCO
            ).grid(row=fila, column=0, sticky="w", pady=4, padx=(0, 15))

            tk.Label(
                frame_tabla, text=f"{CONTRATADO[i]:,} Mbps",
                font=("Segoe UI", 10),
                fg=self.GRIS_SECUNDARIO, bg=self.BLANCO
            ).grid(row=fila, column=1, pady=4, padx=(0, 15))

            entry = tk.Entry(
                frame_tabla, width=14,
                font=("Segoe UI", 10),
                justify="center", relief="solid",
                bd=1, highlightcolor=self.NARANJA,
                highlightthickness=1
            )
            entry.grid(row=fila, column=2, pady=4)
            self.entries_consumo.append(entry)

        # Enter para navegar campos
        for i in range(len(self.entries_consumo) - 1):
            self.entries_consumo[i].bind(
                "<Return>",
                lambda e, siguiente=i+1: self.entries_consumo[siguiente].focus()
            )

        self.entries_consumo[-1].bind(
            "<Return>",
            lambda e: self.generar_reporte()
        )

        # Separador
        tk.Frame(
            self.ventana, bg=self.GRIS_BORDE, height=1
        ).pack(fill="x", padx=30, pady=(5, 0))

        # Botones
        frame_botones = tk.Frame(self.ventana, bg=self.BLANCO, pady=15)
        frame_botones.pack(fill="x", padx=30)

        # Boton generar
        self.boton_generar = tk.Button(
            frame_botones, text="generar reporte por proveeedor",
            font=("Segoe UI", 11, "bold"),
            bg=self.NARANJA, fg=self.BLANCO,
            activebackground="#E57303",
            activeforeground=self.BLANCO,
            cursor="hand2", pady=10,
            relief="flat", bd=0,
            command=self.generar_reporte
        )
        self.boton_generar.pack(fill="x", pady=(0, 8))

        self.boton_generar.bind(
            "<Enter>", lambda e: self.boton_generar.config(bg="#E57303")
        )
        self.boton_generar.bind(
            "<Leave>", lambda e: self.boton_generar.config(bg=self.NARANJA)
        )
        # Boton indicador
        self.boton_indicador = tk.Button(
            frame_botones, text="Indicador de Ancho de Banda",
            font=("Segoe UI", 10),
            bg=self.BLANCO, fg=self.AZUL,
            activebackground=self.GRIS_CLARO,
            activeforeground=self.AZUL,
            cursor="hand2", pady=6,
            relief="solid", bd=1,
            command=self.generar_indicador
        )
        self.boton_indicador.pack(fill="x", pady=(0, 8))

        self.boton_indicador.bind(
            "<Enter>", lambda e: self.boton_indicador.config(
                bg=self.GRIS_CLARO
            )
        )
        self.boton_indicador.bind(
            "<Leave>", lambda e: self.boton_indicador.config(
                bg=self.BLANCO
            )
        )
        # Boton reporte anual
        self.boton_anual = tk.Button(
            frame_botones, text="Reporte Anual",
            font=("Segoe UI", 10),
            bg=self.AZUL, fg=self.BLANCO,
            activebackground="#0A4DB3",
            activeforeground=self.BLANCO,
            cursor="hand2", pady=6,
            relief="flat", bd=0,
            command=self.generar_anual
        )
        self.boton_anual.pack(fill="x", pady=(0, 8))

        self.boton_anual.bind(
            "<Enter>", lambda e: self.boton_anual.config(bg="#0A4DB3")
        )
        self.boton_anual.bind(
            "<Leave>", lambda e: self.boton_anual.config(bg=self.AZUL)
        )

        # Boton limpiar
        self.boton_limpiar = tk.Button(
            frame_botones, text="Limpiar campos",
            font=("Segoe UI", 10),
            bg=self.BLANCO, fg=self.GRIS_SECUNDARIO,
            activebackground=self.GRIS_CLARO,
            activeforeground=self.GRIS_TEXTO,
            cursor="hand2", pady=6,
            relief="solid", bd=1,
            command=self.limpiar_campos
        )
        self.boton_limpiar.pack(fill="x")

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

        # Barra de estado
        frame_estado = tk.Frame(self.ventana, bg=self.GRIS_CLARO)
        frame_estado.pack(fill="x", side="bottom")

        self.label_estado = tk.Label(
            frame_estado, text="Listo",
            font=("Segoe UI", 9),
            bg=self.GRIS_CLARO,
            fg=self.GRIS_SECUNDARIO,
            anchor="w", padx=15, pady=6
        )
        self.label_estado.pack(fill="x")

    def centrar_ventana(self):
        self.ventana.update_idletasks()
        ancho = 520
        alto = 720
        x = (self.ventana.winfo_screenwidth() // 2) - (ancho // 2)
        y = (self.ventana.winfo_screenheight() // 2) - (alto // 2)
        self.ventana.geometry(f"{ancho}x{alto}+{x}+{y}")
    def generar_indicador(self):

        from indicador_pdf import generar_indicador
        from correo import enviar_correo
        from constantes import SMTP_EMAIL, SMTP_PASS

        self.label_estado.config(text="Generando indicador...", fg=self.GRIS_SECUNDARIO)
        self.ventana.update()

        valores_texto = []
        for entry in self.entries_consumo:
            valores_texto.append(entry.get())

        mes = self.combo_mes.get()
        anio = self.entry_anio.get()

        exito, mensaje, consumos = validar_datos(valores_texto)

        if not exito:
            messagebox.showerror("Error", mensaje)
            return

        total_consumo = calcular_total_consumo(consumos)
        porcentaje_general = calcular_porcentaje_general(total_consumo)

        # Generar PDF
        nombre_pdf = generar_indicador(
            mes, anio,
            total_consumo, porcentaje_general
        )

        ruta_pdf = os.path.abspath(nombre_pdf)

        messagebox.showinfo(
            "Indicador generado",
            f"PDF generado correctamente:\n{nombre_pdf}"
        )

        # Preguntar si desea enviarlo
        enviar = messagebox.askyesno(
            "Enviar por correo",
            "¿Desea enviar el indicador por correo?"
        )

        if not enviar:
            os.startfile(ruta_pdf)
            return

        # Pedir correo destino
        destinatario = self.ventana_correo()
        

        if not destinatario:
            return

        asunto = f"Indicador de Ancho de Banda - {mes} {anio}"
        mensaje_correo = (
            f"Adjunto se encuentra el indicador de ancho de banda.\n\n"
            f"Mes: {mes}\n"
            f"Año: {anio}\n"
            f"Consumo total: {total_consumo:,.2f} Mbps\n"
            f"Uso general: {porcentaje_general}%"
        )

        enviado, respuesta = enviar_correo(
            SMTP_EMAIL,
            SMTP_PASS,
            destinatario,
            asunto,
            mensaje_correo,
            ruta_pdf
        )

        if enviado:
            messagebox.showinfo("Correo enviado", respuesta)
        else:
            messagebox.showerror("Error", respuesta)

    def generar_reporte(self):

        self.label_estado.config(text="Validando...", fg=self.GRIS_SECUNDARIO)
        self.ventana.update()

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

        exito, mensaje, consumos = validar_datos(valores_texto)

        if not exito:
            messagebox.showerror("Error de validacion", mensaje)
            self.label_estado.config(text="Error en los datos", fg=self.ROJO_ERROR)
            return

        self.label_estado.config(text="Generando...", fg=self.GRIS_SECUNDARIO)
        self.ventana.update()

        porcentajes = calcular_porcentajes(consumos)
        total_consumo = calcular_total_consumo(consumos)
        porcentaje_general = calcular_porcentaje_general(total_consumo)

        try:
            nombre_archivo = generar_reporte(
                mes, anio,
                consumos, porcentajes,
                total_consumo, porcentaje_general
            )

            # Verificar si el mes ya existe en el historial
            guardar = True

            if mes_existe(anio, mes):
                respuesta = messagebox.askyesno(
                    "Mes existente",
                    f"Ya existe un reporte de {mes} {anio}.\n"
                    f"Desea reemplazarlo?"
                )
                if not respuesta:
                    guardar = False

            if guardar:
                guardar_mes(
                    anio, mes,
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

            os.startfile(ruta_completa)

        except PermissionError:
            messagebox.showerror(
                "Error",
                "No se pudo guardar el archivo.\n"
                "Verifique que no este abierto en Excel."
            )
            self.label_estado.config(text="Error al guardar", fg=self.ROJO_ERROR)

        except Exception as e:
            messagebox.showerror("Error", f"Error inesperado(verifique si no está abierto en excel, cerrar primero):\n{str(e)}")
            self.label_estado.config(text="Error inesperado", fg=self.ROJO_ERROR)
    def ventana_correo(self):
        """
        Ventana moderna para solicitar correo destino.
        """

        ventana = tk.Toplevel(self.ventana)
        ventana.title("Enviar Indicador")
        ventana.geometry("400x200")
        ventana.resizable(False, False)
        ventana.configure(bg=self.BLANCO)

        ventana.grab_set()  # Bloquea la principal hasta cerrar

        tk.Label(
            ventana,
            text="Enviar Indicador por Correo",
            font=("Segoe UI", 14, "bold"),
            fg=self.AZUL,
            bg=self.BLANCO
        ).pack(pady=(20,10))

        tk.Label(
            ventana,
            text="Correo destino:",
            font=("Segoe UI", 10),
            fg=self.GRIS_TEXTO,
            bg=self.BLANCO
        ).pack()

        entry_correo = tk.Entry(
            ventana,
            font=("Segoe UI", 10),
            width=35,
            relief="solid",
            bd=1
        )
        entry_correo.pack(pady=10)

        resultado = {"correo": None}

        def confirmar():
            correo = entry_correo.get().strip()
            if correo == "":
                messagebox.showerror("Error", "Ingrese un correo válido.")
                return
            resultado["correo"] = correo
            ventana.destroy()

        boton = tk.Button(
            ventana,
            text="Enviar",
            font=("Segoe UI", 10, "bold"),
            bg=self.NARANJA,
            fg=self.BLANCO,
            activebackground="#E57303",
            activeforeground=self.BLANCO,
            relief="flat",
            command=confirmar
        )
        boton.pack(pady=10)

        ventana.wait_window()

        return resultado["correo"]
    def generar_anual(self):

        anio = self.entry_anio.get()

        if anio.strip() == "" or not anio.isdigit():
            messagebox.showerror("Error", "Ingrese un año valido.")
            return

        datos = obtener_datos_anio(anio)

        if datos is None or len(datos) == 0:
            messagebox.showerror(
                "Sin datos",
                f"No hay reportes generados para el año {anio}.\n"
                f"Genere al menos un reporte mensual primero."
            )
            return

        try:
            nombre = generar_reporte_anual(anio, datos)
            ruta = os.path.abspath(nombre)

            messagebox.showinfo(
                "Reporte Anual generado",
                f"Archivo: {nombre}\n"
                f"Meses registrados: {len(datos)}\n"
                f"Ubicacion: {ruta}"
            )

            self.label_estado.config(
                text=f"Generado: {nombre}",
                fg=self.VERDE_EXITO
            )

            os.startfile(ruta)

        except Exception as e:
            messagebox.showerror("Error", f"Error inesperado:\n{str(e)}")
            self.label_estado.config(text="Error inesperado", fg=self.ROJO_ERROR)

    def limpiar_campos(self):
        for entry in self.entries_consumo:
            entry.delete(0, tk.END)
        self.label_estado.config(text="Listo", fg=self.GRIS_SECUNDARIO)

    def ejecutar(self):
        self.ventana.mainloop()