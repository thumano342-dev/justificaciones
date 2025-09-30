from flask import Flask, render_template, request, redirect, url_for, send_file, flash
import mysql.connector
import pandas as pd
import os
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "secreto"

# Carpeta donde se guardarán los excels subidos
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# Conexión a MySQL
def get_db():
    return mysql.connector.connect(
        host="localhost",
        user="root",
        password="mi2024",
        database="justificaciones"
    )

# Página principal del admin
@app.route("/")
def index():
    conn = get_db()
    cur = conn.cursor(dictionary=True)
    cur.execute("SELECT * FROM archivos_subidos ORDER BY fecha_subida DESC")
    archivos = cur.fetchall()
    conn.close()
    return render_template("admin.html", archivos=archivos)

# Subir archivo Excel
@app.route("/subir", methods=["POST"])
def subir():
    if "archivo" not in request.files:
        flash("No se seleccionó archivo", "danger")
        return redirect(url_for("index"))

    file = request.files["archivo"]
    if file.filename == "":
        flash("Nombre de archivo vacío", "danger")
        return redirect(url_for("index"))

    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    file.save(file_path)

    # Guardar info en MySQL desde Excel
    df = pd.read_excel(file_path)
    conn = get_db()
    cur = conn.cursor()

    for _, row in df.iterrows():
        cur.execute("""
            INSERT INTO attendance (canal, fecha, dato1, dato2, completado)
            VALUES (%s, %s, %s, %s, %s)
        """, (
            row["canal"],
            row["fecha"],
            row["dato1"],
            row["dato2"],
            row["completado"]
        ))

    # Registrar el archivo
    cur.execute("""
        INSERT INTO archivos_subidos (nombre_archivo, fecha_subida)
        VALUES (%s, %s)
    """, (filename, datetime.now().date()))

    conn.commit()
    conn.close()

    flash("Archivo subido y datos guardados en MySQL", "success")
    return redirect(url_for("index"))

# Descargar archivo filtrado por fecha
@app.route("/descargar/<fecha>")
def descargar(fecha):
    conn = get_db()
    cur = conn.cursor(dictionary=True)
    cur.execute("SELECT * FROM attendance WHERE fecha = %s AND completado = TRUE", (fecha,))
    datos = cur.fetchall()
    conn.close()

    if not datos:
        flash("No hay canales completados en esa fecha", "warning")
        return redirect(url_for("index"))

    df = pd.DataFrame(datos)
    output_path = f"attendance_{fecha.replace('-', '')}.xlsx"
    df.to_excel(output_path, index=False)

    return send_file(output_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
