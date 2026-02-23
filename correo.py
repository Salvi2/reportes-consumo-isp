"""
Sistema de Reportes de Consumo ISP
Módulo de Envío de Correo (SMTP Gmail)
"""

import smtplib
from email.message import EmailMessage
import ssl
import os


def enviar_correo(remitente, password_app, destinatario, asunto, mensaje, archivo_adjunto):

    try:
        msg = EmailMessage()
        msg["From"] = remitente
        msg["To"] = destinatario
        msg["Subject"] = asunto
        msg.set_content(mensaje)

        # Adjuntar archivo PDF
        with open(archivo_adjunto, "rb") as f:
            archivo = f.read()
            nombre_archivo = os.path.basename(archivo_adjunto)

        msg.add_attachment(
            archivo,
            maintype="application",
            subtype="pdf",
            filename=nombre_archivo
        )

        contexto = ssl.create_default_context()

        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=contexto) as servidor:
            servidor.login(remitente, password_app)
            servidor.send_message(msg)

        return True, "Correo enviado correctamente."

    except Exception as e:
        return False, f"Error al enviar correo:\n{str(e)}"