import smtplib
from email.mime.text import MIMEText

SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587
USERNAME = "josue.romero@spradling.group"
PASSWORD = "kyszjdvrjmngphkt"


def enviar_correo():
    msg = MIMEText("Prueba de envío SMTP desde Python")
    msg["Subject"] = "Correo automático"
    msg["From"] = USERNAME
    msg["To"] = "josue.romero@spradling.group"

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(USERNAME, PASSWORD)
        server.send_message(msg)
        print("Correo enviado!")


enviar_correo()
