from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
import os
from extractor import extraer_datos_factura, guardar_en_excel

app = Flask(__name__, static_folder=".")
CORS(app)

UPLOAD_FOLDER = "uploads"
EXCEL_PATH = "resultados.xlsx"

# Sirve el HTML principal
@app.route("/")
def inicio():
    return send_from_directory(".", "index.html")

@app.route("/procesar", methods=["POST"])
def procesar_facturas():
    archivos = request.files.getlist("archivos")

    if not archivos:
        return jsonify({"error": "No se han enviado archivos"}), 400

    extensiones_permitidas = {".pdf", ".jpg", ".jpeg", ".png", ".webp"}
    resultados = []

    for archivo in archivos:
        extension = os.path.splitext(archivo.filename)[1].lower()

        if extension not in extensiones_permitidas:
            resultados.append({
                "archivo": archivo.filename,
                "error": "Formato no permitido"
            })
            continue

        ruta_pdf = os.path.join(UPLOAD_FOLDER, archivo.filename)
        archivo.save(ruta_pdf)

        datos = extraer_datos_factura(ruta_pdf)
        guardar_en_excel(datos, EXCEL_PATH)

        resultados.append({
            "archivo": archivo.filename,
            "datos": datos
        })

    return jsonify({"resultados": resultados})


@app.route("/descargar", methods=["GET"])
def descargar_excel():
    if os.path.exists(EXCEL_PATH):
        return send_file(EXCEL_PATH, as_attachment=True)
    return jsonify({"error": "No hay Excel generado todavía"}), 404


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))

    app.run(host="0.0.0.0", port=port)
