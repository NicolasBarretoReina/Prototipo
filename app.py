from flask import Flask, render_template, request, send_file
import pandas as pd
import os
import uuid

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    archivo = request.files['archivo']
    if archivo.filename == '':
        return "No se subió ningún archivo", 400

    ext = archivo.filename.split('.')[-1].lower()
    nombre_archivo = f"{uuid.uuid4()}.{ext}"
    ruta_guardado = os.path.join(UPLOAD_FOLDER, nombre_archivo)
    archivo.save(ruta_guardado)

    if ext == 'csv':
        df = pd.read_csv(ruta_guardado)
    elif ext == 'xlsx':
        df = pd.read_excel(ruta_guardado, engine='openpyxl')
    else:
        return "Formato de archivo no válido", 400

    columnas = list(df.columns)
    df.to_csv(os.path.join(UPLOAD_FOLDER, 'temp.csv'), index=False)
    return render_template('seleccionar.html', columnas=columnas, archivo=nombre_archivo)

@app.route('/generar', methods=['POST'])
def generar():
    columnas_seleccionadas = request.form.getlist('columnas')
    archivo = request.form['archivo']
    ruta_archivo = os.path.join(UPLOAD_FOLDER, archivo)

    if len(columnas_seleccionadas) < 2:
        return "Selecciona al menos dos columnas para graficar", 400

    ext = archivo.split('.')[-1].lower()
    if ext == 'csv':
        df = pd.read_csv(ruta_archivo)
    elif ext == 'xlsx':
        df = pd.read_excel(ruta_archivo, engine='openpyxl')
    else:
        return "Archivo no compatible", 400

    salida = os.path.join(UPLOAD_FOLDER, archivo.rsplit('.', 1)[0] + '_grafico.xlsx')

    writer = pd.ExcelWriter(salida, engine='xlsxwriter')
    df_filtrado = df[columnas_seleccionadas]
    hoja_datos = 'DatosFiltrados'
    df_filtrado.to_excel(writer, sheet_name=hoja_datos, index=False)

    workbook = writer.book
    worksheet = writer.sheets[hoja_datos]

    col_x = columnas_seleccionadas[0]

    for i, col_y in enumerate(columnas_seleccionadas[1:], start=1):
        chart = workbook.add_chart({'type': 'line'})

        col_x_letter = 'A'
        col_y_letter = chr(ord('A') + i)

        chart.add_series({
            'name':       col_y,
            'categories': f"='{hoja_datos}'!${col_x_letter}$2:${col_x_letter}${len(df_filtrado)+1}",
            'values':     f"='{hoja_datos}'!${col_y_letter}$2:${col_y_letter}${len(df_filtrado)+1}",
            'marker':     {'type': 'circle'},
        })

        chart.set_title({'name': f'{col_y} vs {col_x}'})
        chart.set_x_axis({'name': col_x})
        chart.set_y_axis({'name': col_y})

        hoja_grafico = f'Grafico_{col_y}'
        chart_sheet = workbook.add_worksheet(hoja_grafico)
        chart_sheet.insert_chart('B2', chart)

    writer.close()
    return send_file(salida, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)