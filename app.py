from flask import Flask, render_template, redirect, url_for, request, flash, session
import pandas as pd
from flask import jsonify
from conexion.conexion_base import obtener_conexion
from conexion.consulta_sql import ejecutar_consulta
import unicodedata
import os, zipfile, tempfile
from werkzeug.utils import secure_filename
from excel_utils import llenar_excel

app = Flask(__name__)
app.secret_key = 'ocaglobal_secret'  # Necesario para usar flash


@app.route('/')
def menu():
    return render_template('menu.html')


@app.route('/asistencia')
def asistencia():
    return render_template('asistencia.html')


@app.route('/checklist')
def checklist():
    return redirect(url_for('menu'))  # Aquí luego irá tu lógica real


@app.route('/procesar_asistencia', methods=['POST'])
def procesar_asistencia():
    checklist_id = request.form.get('checklist_id')
    if not checklist_id.isdigit():
        flash("ID inválido", "error")
        return redirect(url_for('asistencia'))

    conn = obtener_conexion()
    if not conn:
        flash("No se pudo conectar a la base de datos", "error")
        return redirect(url_for('asistencia'))

    data = ejecutar_consulta(conn, checklist_id)
    conn.close()

    if data is None or data.empty:
        flash("No se encontraron datos para ese ID", "error")
        return redirect(url_for('asistencia'))

    # Solo guardamos el ID para futuras consultas
    session['checklist_id'] = checklist_id

    temas = sorted(data['Tema Tratado'].dropna().unique())
    return render_template('asistencia.html', data=data.to_dict(orient='records'), columnas=data.columns, temas=temas, fecha_actual=None)


@app.route('/filtrar_fecha_tema', methods=['POST'])
def filtrar_fecha_tema():
    fecha = request.form.get('fecha_filtro')
    tema = request.form.get('tema_filtro')

    if not fecha:
        flash("Debe seleccionar una fecha", "error")
        return redirect(url_for('asistencia'))

    conn = obtener_conexion()
    if not conn or 'checklist_id' not in session:
        flash("Error de conexión o falta el ID", "error")
        return redirect(url_for('asistencia'))

    data = ejecutar_consulta(conn, session['checklist_id'])
    conn.close()

    # Convertimos y filtramos por fecha
    data['FechaSistema'] = pd.to_datetime(data['FechaSistema'], errors='coerce')
    fecha_dt = pd.to_datetime(fecha, errors='coerce')
    data_fecha = data[data['FechaSistema'].dt.date == fecha_dt.date()]

    # Obtenemos temas SOLO de esa fecha
    temas_disponibles = sorted(data_fecha['Tema Tratado'].dropna().unique())

    # Si ya seleccionó tema, lo filtramos
    if tema:
        data_filtrada = data_fecha[data_fecha['Tema Tratado'].astype(str).str.lower().str.contains(tema.lower())]
    else:
        data_filtrada = data_fecha

    if data_filtrada.empty:
        flash("No se encontraron registros para ese filtro", "error")

    return render_template('asistencia.html',
                           data=data_filtrada.to_dict(orient='records'),
                           columnas=data.columns,
                           temas=temas_disponibles,
                           fecha_actual=fecha)


@app.route('/ajax_filtrar', methods=['POST'])
def ajax_filtrar():
    from unicodedata import normalize

    fecha = request.json.get('fecha')
    temas_seleccionados = request.json.get('temas', [])

    if 'checklist_id' not in session:
        return jsonify({"data": [], "temas": [], "columnas": []})

    conn = obtener_conexion()
    data = ejecutar_consulta(conn, session['checklist_id'])
    conn.close()

    data['FechaSistema'] = pd.to_datetime(data['FechaSistema'], errors='coerce')
    fecha_dt = pd.to_datetime(fecha, errors='coerce')
    data_fecha = data[data['FechaSistema'].dt.date == fecha_dt.date()]

    temas_disponibles = sorted(data_fecha['Tema Tratado'].dropna().unique())

    def normalizar(texto):
        texto = str(texto).lower().strip()
        return normalize('NFKD', texto).encode('ascii', 'ignore').decode('utf-8')

    if temas_seleccionados:
        temas_normalizados = [normalizar(t) for t in temas_seleccionados if t]
        data_filtrada = data_fecha[
            data_fecha['Tema Tratado'].apply(lambda x: normalizar(x) in temas_normalizados)
        ]
    else:
        data_filtrada = data_fecha

    columnas = list(data.columns)
    registros = data_filtrada.to_dict(orient='records')

    # Guardar selección en sesión para exportación
    session['fecha_actual'] = fecha
    session['temas_filtrados'] = temas_seleccionados

    return jsonify({"columnas": columnas, "data": registros, "temas": temas_disponibles})


@app.route('/exportar_excel', methods=['POST'])
def exportar_excel():
    import unicodedata

    texto_g50 = request.form.get("texto_g50")
    texto_g52 = request.form.get("texto_g52")
    firma_imagen = request.files.get("firma_imagen")
    firmas_zip = request.files.get("firmas_zip")

    if 'checklist_id' not in session:
        flash("Primero debes cargar un ID de checklist", "error")
        return redirect(url_for('asistencia'))

    conn = obtener_conexion()
    data = ejecutar_consulta(conn, session['checklist_id'])
    conn.close()

    # Aplicar mismos filtros que en ajax_filtrar
    fecha = session.get('fecha_actual')
    temas = session.get('temas_filtrados', [])

    data['FechaSistema'] = pd.to_datetime(data['FechaSistema'], errors='coerce')
    fecha_dt = pd.to_datetime(fecha, errors='coerce')
    data_fecha = data[data['FechaSistema'].dt.date == fecha_dt.date()]

    def normalizar(texto):
        texto = str(texto).lower().strip()
        return unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('utf-8')

    if temas:
        temas_normalizados = [normalizar(t) for t in temas]
        data_filtrada = data_fecha[
            data_fecha['Tema Tratado'].apply(lambda x: normalizar(x) in temas_normalizados)
        ]
    else:
        data_filtrada = data_fecha

    # Crear carpeta de exportación
    output_folder = os.path.join("static", "exports")
    os.makedirs(output_folder, exist_ok=True)

    # Guardar firma personal
    firma_path = os.path.join(tempfile.gettempdir(), secure_filename(firma_imagen.filename))
    firma_imagen.save(firma_path)

    # Extraer ZIP de firmas
    temp_dir = tempfile.mkdtemp()
    zip_path = os.path.join(temp_dir, secure_filename(firmas_zip.filename))
    firmas_zip.save(zip_path)
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    # Generar archivo Excel
    fecha_str = fecha_dt.strftime("%d-%m-%Y")
    nombre_archivo = f"Charla_Capacitación_{fecha_str}.xlsx"
    output_path = os.path.join(output_folder, nombre_archivo)

    from conexion.config import plantilla_path
    llenar_excel(data_filtrada, output_path, plantilla_path, texto_g50, texto_g52, firma_path, temp_dir)

    # Redirigir a la descarga directa
    return redirect(url_for('static', filename=f"exports/{nombre_archivo}"))


if __name__ == '__main__':
    app.run(debug=True)

