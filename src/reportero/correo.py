"""Envío del reporte por correo (SMTP). Las credenciales se leen de un archivo .env."""
from __future__ import annotations

import os
import smtplib
from email.message import EmailMessage
from pathlib import Path


def cargar_env(ruta: Path = Path(".env")) -> None:
    """Carga variables CLAVE=valor desde un archivo .env (sin dependencias externas)."""
    if not ruta.exists():
        return
    for linea in ruta.read_text(encoding="utf-8").splitlines():
        linea = linea.strip()
        if linea and not linea.startswith("#") and "=" in linea:
            clave, valor = linea.split("=", 1)
            os.environ.setdefault(clave.strip(), valor.strip().strip('"').strip("'"))


def enviar_reporte(destinatarios: list[str], adjuntos: list[Path], resumen: str) -> None:
    cargar_env()
    requeridas = ["SMTP_HOST", "SMTP_PORT", "SMTP_USUARIO", "SMTP_CLAVE"]
    faltan = [v for v in requeridas if not os.getenv(v)]
    if faltan:
        raise RuntimeError(f"Faltan variables en .env: {', '.join(faltan)} (ver .env.ejemplo)")

    msg = EmailMessage()
    msg["Subject"] = "Reporte de ventas consolidado"
    msg["From"] = os.getenv("SMTP_REMITENTE", os.environ["SMTP_USUARIO"])
    msg["To"] = ", ".join(destinatarios)
    msg.set_content(
        "Hola,\n\nAdjunto el reporte de ventas generado automáticamente.\n\n"
        f"{resumen}\n\nSaludos."
    )
    for archivo in adjuntos:
        tipo = "pdf" if archivo.suffix == ".pdf" else "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        msg.add_attachment(archivo.read_bytes(), maintype="application", subtype=tipo, filename=archivo.name)

    with smtplib.SMTP(os.environ["SMTP_HOST"], int(os.environ["SMTP_PORT"])) as smtp:
        smtp.starttls()
        smtp.login(os.environ["SMTP_USUARIO"], os.environ["SMTP_CLAVE"])
        smtp.send_message(msg)
