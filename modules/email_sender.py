import smtplib
import time
import urllib.parse

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


GMAIL = "robotmercadopublicov4@gmail.com"

PASSWORD = "fkejbgvxrgnunbbq"


def enviar_correo(destinatario, empresa):

    asunto = "Propuesta de Factoring para Finanzas"

    # URL TRACKING APERTURA
    tracking_url = f"https://banproyecta-tracking.onrender.com/track?email={destinatario}&t={time.time()}"

    # LINK REAL
    link_real = "https://www.banproyecta.cl/"

    # CODIFICAR LINK
    destino_codificado = urllib.parse.quote(link_real)

    # LINK TRACKEADO
    click_url = f"https://banproyecta-tracking.onrender.com/click?email={destinatario}&destino={destino_codificado}"

    html = f"""
    <html>
        <body>

            <p>
               Hola. Somos Factoring Banproyecta.
               Anticipo del 100% del pago de sus facturas a tasas bajas.
            </p>

            <p>
                Quedamos atentos.
            </p>

            <p>
                <a href="{click_url}">
                    Ver propuesta comercial
                </a>
            </p>

            <!-- PIXEL INVISIBLE -->
            <img src="{tracking_url}"
                 width="1"
                 height="1"
                 style="display:none;" />

        </body>
    </html>
    """

    msg = MIMEMultipart("alternative")

    msg["Subject"] = asunto

    msg["From"] = GMAIL

    msg["To"] = destinatario

    parte_html = MIMEText(html, "html")

    msg.attach(parte_html)

    servidor = smtplib.SMTP("smtp.gmail.com", 587)

    servidor.starttls()

    servidor.login(GMAIL, PASSWORD)

    servidor.sendmail(
        GMAIL,
        destinatario,
        msg.as_string()
    )

    servidor.quit()

    print(tracking_url)

    print(click_url)

    print("Correo enviado")