import pythoncom
from flask import Flask, render_template, request, send_file, send_from_directory
import mysql.connector
import os
import win32com.client
from datetime import datetime, date  # Importar 'date' desde datetime

app = Flask(__name__)

# Configuración de conexión MySQL
DB_CONFIG = {
    'host': 'localhost',
    'user': 'root',
    'password': 'Admin123',
    'database': 'facturas_db'
}

# Ruta para la carpeta de archivos generados
FILES_FOLDER = r"C:\Users\Usuario\Desktop\PROGRAMACION\Python\Flask\Integracion\AppB\FILES_FOLDER"

if not os.path.exists(FILES_FOLDER):
    os.makedirs(FILES_FOLDER)


@app.route("/", methods=["GET", "POST"])
def consultar_facturas():
    facturas = []
    if request.method == "POST":
        id_cliente = request.form["id_cliente"]

        # Consultar las facturas en la base de datos
        connection = mysql.connector.connect(**DB_CONFIG)
        cursor = connection.cursor()
        cursor.execute('''
            SELECT no_factura, condiciones, id_cliente, fecha_factura, monto, estado, ruta_imagen 
            FROM facturas WHERE id_cliente = %s
        ''', (id_cliente,))
        facturas = cursor.fetchall()
        cursor.close()
        connection.close()

        # Generar el archivo Excel utilizando OLE
        if facturas:
            generate_excel_ole(facturas)

    return render_template("consulta.html", facturas=facturas)


def generate_excel_ole(facturas):
    # Inicializar COM
    pythoncom.CoInitialize()

    # Crear una instancia de Excel
    excel = win32com.client.Dispatch("Excel.Application")
    
    # Verificar si la instancia de Excel es válida
    if not excel:
        print("No se pudo iniciar Excel.")
        return

    # Crear un libro de Excel
    try:
        workbook = excel.Workbooks.Add()  # Esto debería funcionar
    except Exception as e:
        print(f"Error al crear el libro de trabajo: {e}")
        return
    
    sheet = workbook.Sheets(1)
    sheet.Name = "Facturas"

    # Títulos de columnas en la primera fila
    sheet.Cells(1, 1).Value = "Factura"
    sheet.Cells(1, 2).Value = "Fecha"
    sheet.Cells(1, 3).Value = "Cliente"
    sheet.Cells(1, 4).Value = "Monto"
    
    # Llenar las celdas con los datos de las facturas
    row = 2  # Comenzamos desde la segunda fila
    for factura in facturas:
        factura_num = factura.get('factura', '')
        fecha = factura.get('fecha', datetime.today()).date()
        cliente = factura.get('cliente', '')
        monto = factura.get('monto', 0.0)
        
        # Asegúrate de manejar correctamente el tipo de dato
        if isinstance(fecha, date):
            sheet.Cells(row, 2).Value = fecha.strftime("%d/%m/%Y")  # Fecha en formato d/m/a
        else:
            sheet.Cells(row, 2).Value = fecha
        
        sheet.Cells(row, 1).Value = factura_num
        sheet.Cells(row, 3).Value = cliente
        sheet.Cells(row, 4).Value = monto
        
        row += 1

    # Guardar el archivo Excel
    file_path = "C:\\path\\to\\your\\facturas.xlsx"  # Cambia la ruta del archivo
    workbook.SaveAs(file_path)
    
    # Cerrar Excel
    workbook.Close(False)  # No guardar los cambios al cerrar (ya lo hicimos)
    excel.Quit()  # Cerrar la aplicación Excel

@app.route("/descargar_imagen/<path:filename>")
def descargar_imagen(filename):
    return send_from_directory(FILES_FOLDER, filename, as_attachment=True)


@app.route("/descargar_excel")
def descargar_excel():
    excel_path = os.path.join(FILES_FOLDER, "facturas.xlsx")
    return send_file(excel_path, as_attachment=True, download_name="facturas.xlsx")


if __name__ == "__main__":
    app.run(debug=True, port=5001)

