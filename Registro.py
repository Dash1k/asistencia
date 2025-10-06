from flask import Flask
import openpyxl
import os
from datetime import datetime

app = Flask(__name__)
archivo = os.path.expanduser("~/Downloads/asistencias.xlsx")

Grupo = "552"

# 🔹 Configuración de horarios → hoja a usar
HORARIOS = {
    "ERGONOMIA": range(15, 16),  # de 14:00 a 14:59
    "CONTROL": range(16, 17),         # de 15:00 a 15:59
    "IO2": range(17, 18)      # de 16:00 a 16:59
}

@app.route("/registrar/<num_control>")
def registrar(num_control):
    # 🔹 Solo abrir el archivo, asumiendo que ya existe
    if not os.path.exists(archivo):
        return "El archivo de asistencias no existe. Créalo primero.", 500

    wb = openpyxl.load_workbook(archivo)

    # Obtener hora actual
    ahora = datetime.now()
    fecha = ahora.strftime("%Y-%m-%d")
    hora = ahora.strftime("%H:%M:%S")
    hora_actual = ahora.hour  # solo la hora entera

    # 🔹 Determinar hoja según la hora
    hoja_destino = None
    for materia, horas in HORARIOS.items():
        if hora_actual in horas:
            hoja_destino = materia
            break

    if not hoja_destino:
        return f"No hay materia configurada para esta hora ({hora_actual})", 400

    # Seleccionar la hoja correspondiente
    if hoja_destino not in wb.sheetnames:
        return f"La hoja {hoja_destino} no existe en el archivo.", 500

    ws = wb[hoja_destino]
    ws.append([num_control, Grupo, fecha, hora])

    wb.save(archivo)

    return f"""
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Registro de Asistencia</title>
        <style>
            body {{
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100vh;
                font-size: 3em;
                text-align: center;
                font-family: Arial, sans-serif;
                background-color: #f9f9f9;
            }}
            .contenedor {{
                display: flex;
                flex-direction: column;
                gap: 20px; /* espacio entre líneas */
            }}
        </style>
    </head>
    <body>
        <div class="contenedor">
            <div>✅ ASISTENCIA REGISTRADA PARA:</div>
            <div><strong>{num_control}</strong></div>
            <div>Materia:</div>
            <div><strong>{hoja_destino}</strong></div>
        </div>
    </body>
    </html>
    """

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
