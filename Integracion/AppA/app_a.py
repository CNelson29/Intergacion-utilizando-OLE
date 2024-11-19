from flask import Flask, render_template, request, redirect, url_for
import mysql.connector
import subprocess
import shutil

app = Flask(__name__)

# Configuración de conexión MySQL
DB_CONFIG = {
    'host': 'localhost',
    'user': 'root',
    'password': 'Admin123',
    'database': 'facturas_db'
}



def simular_escaneo(ruta_imagen):
    # Simula un archivo de imagen copiando una imagen existente o creando un marcador
    imagen_simulada = "static/images/placeholder.png"  # Imagen preexistente para usar como simulación
    try:
        shutil.copy(imagen_simulada, ruta_imagen)
        print(f"Imagen simulada guardada en: {ruta_imagen}")
    except FileNotFoundError:
        # Si no tienes una imagen, crea un archivo vacío
        with open(ruta_imagen, "w") as f:
            f.write("Esta es una imagen simulada de una factura.")
        print(f"Archivo simulado creado en: {ruta_imagen}")



# Ruta principal para registrar facturas
@app.route("/", methods=["GET", "POST"])
def registrar_factura():
    if request.method == "POST":
        no_factura = request.form["no_factura"]
        condiciones = request.form["condiciones"]
        id_cliente = request.form["id_cliente"]
        fecha_factura = request.form["fecha_factura"]
        monto = request.form["monto"]
        estado = request.form["estado"]

        # Generar la ruta para la imagen de la factura
        ruta_imagen = f"static/images/{no_factura}.png"

        # Ejecutar el comando para escanear y guardar la imagen
        simular_escaneo(ruta_imagen)

        # Guardar datos en la base de datos MySQL
        connection = mysql.connector.connect(**DB_CONFIG)
        cursor = connection.cursor()
        cursor.execute('''
            INSERT INTO facturas (no_factura, condiciones, id_cliente, fecha_factura, monto, estado, ruta_imagen)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        ''', (no_factura, condiciones, id_cliente, fecha_factura, monto, estado, ruta_imagen))
        connection.commit()
        cursor.close()
        connection.close()

        return redirect(url_for("registrar_factura"))

    return render_template("form.html")

if __name__ == "__main__":
    app.run(debug=True)