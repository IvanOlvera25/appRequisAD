# comisiones/mail.py
import smtplib, mimetypes, os
from email.message import EmailMessage
from flask import current_app

def get_recipients_from_string(s: str):
    if not s:
        return []
    raw = [p.strip() for p in s.replace(";", ",").split(",")]
    return [p for p in raw if p]

def send_email_basic(subject, body, to_emails, attachment_path=None):
    """Envía correo usando credenciales definidas en config: EMAIL_USER / EMAIL_PASS."""
    if not to_emails:
        return
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = current_app.config.get("EMAIL_USER", "ad17solutionsbot@gmail.com")
    msg['To'] = ", ".join(to_emails)
    msg.set_content(body)

    if attachment_path and os.path.isfile(attachment_path):
        mime_type, _ = mimetypes.guess_type(attachment_path)
        if not mime_type:
            mime_type = "application/octet-stream"
        maintype, subtype = mime_type.split("/", 1)
        with open(attachment_path, "rb") as f:
            msg.add_attachment(f.read(), maintype=maintype, subtype=subtype, filename=os.path.basename(attachment_path))

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(
                current_app.config.get("EMAIL_USER", "ad17solutionsbot@gmail.com"),
                current_app.config.get("EMAIL_PASS", "misvtfhrnwbmiptb"),
            )
            smtp.send_message(msg)
    except Exception as e:
        print(f"[Comisiones] Error al enviar correo: {e}")

def build_comision_body(c: dict) -> str:
    base = f"""Hola {c.get('trabajador','')},

Te informamos un cambio en el estado de tu comisión.

Proyecto: {c.get('proyecto','')}
Monto del proyecto: ${float(c.get('monto_proyecto',0) or 0):.2f}
Tipo de cálculo: {c.get('tipo_calculo','')}
Parámetro: {c.get('porcentaje','0')}% / ${float(c.get('monto_comision',0) or 0):.2f}
Comisión calculada: ${float(c.get('monto_calculado',0) or 0):.2f}

Estado actual: {c.get('estado','')}
Fecha de registro: {c.get('fecha','')}
"""
    return base