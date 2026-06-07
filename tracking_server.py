from flask import Flask, request, send_file, redirect
from datetime import datetime
import pandas as pd
import os

app = Flask(__name__)

PIXEL_PATH = "pixel.png"

# crear pixel si no existe
if not os.path.exists(PIXEL_PATH):

    from PIL import Image

    img = Image.new("RGBA", (1, 1), (255, 255, 255, 0))

    img.save(PIXEL_PATH)


# =========================================
# TRACKING APERTURA
# =========================================

@app.route("/track")
def track():

    email = request.args.get("email", "")

    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    ip = request.remote_addr

    user_agent = request.headers.get("User-Agent")

    print("\nCORREO ABIERTO")
    print(email)
    print(fecha)
    print(ip)
    print(user_agent)

    archivo = "tracking.xlsx"

    nuevo = pd.DataFrame([{
        "email": email,
        "fecha_apertura": fecha,
        "ip": ip,
        "user_agent": user_agent
    }])

    if os.path.exists(archivo):

        existente = pd.read_excel(archivo)

        final = pd.concat([existente, nuevo], ignore_index=True)

    else:

        final = nuevo

    final.to_excel(archivo, index=False)

    return send_file(PIXEL_PATH, mimetype="image/png")


# =========================================
# TRACKING CLICKS
# =========================================

@app.route("/click")
def click():

    email = request.args.get("email", "")

    destino = request.args.get("destino", "")

    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    ip = request.remote_addr

    user_agent = request.headers.get("User-Agent")

    print("\nCLICK DETECTADO")
    print(email)
    print(destino)
    print(fecha)

    archivo = "clicks.xlsx"

    nuevo = pd.DataFrame([{
        "email": email,
        "destino": destino,
        "fecha_click": fecha,
        "ip": ip,
        "user_agent": user_agent
    }])

    if os.path.exists(archivo):

        existente = pd.read_excel(archivo)

        final = pd.concat([existente, nuevo], ignore_index=True)

    else:

        final = nuevo

    final.to_excel(archivo, index=False)

    return redirect(destino)


# =========================================
# INICIAR FLASK
# =========================================

if __name__ == "__main__":

    app.run(port=5000)