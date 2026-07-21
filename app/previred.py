import os
import uuid
from flask import Blueprint, render_template, request, send_file, current_app, flash
from flask_login import login_required

from .logic.previred_engine import extraer_datos_pdf, consolidar_datos, exportar_excel

previred_bp = Blueprint('previred', __name__)


@previred_bp.route('/', methods=['GET', 'POST'])
@login_required
def formulario():
    if request.method == 'POST':
        archivos = request.files.getlist('archivos_pdf')
        archivos = [a for a in archivos if a and a.filename]

        if not archivos:
            flash('Debes subir al menos un PDF de Previred.', 'error')
            return render_template('previred/formulario.html')

        upload_dir = current_app.config['UPLOAD_FOLDER']
        token = uuid.uuid4().hex[:8]
        rutas_guardadas = []

        try:
            registros = []
            for archivo in archivos:
                ruta = os.path.join(upload_dir, f'{token}_{archivo.filename}')
                archivo.save(ruta)
                rutas_guardadas.append(ruta)
                registros.extend(extraer_datos_pdf(ruta))

            if not registros:
                flash('No se pudo extraer información de los PDF subidos. Revisa que sean los documentos correctos de Previred.', 'error')
                return render_template('previred/formulario.html')

            df = consolidar_datos(registros)
            if df.empty:
                flash('No se encontraron trabajadores al consolidar los datos.', 'error')
                return render_template('previred/formulario.html')

            nombre_salida = f'{token}_previred_consolidado.xlsx'
            ruta_salida = os.path.join(current_app.config['OUTPUT_FOLDER'], nombre_salida)
            exportar_excel(df, ruta_salida)
        except Exception as e:
            flash(f'Error al procesar los PDF: {e}', 'error')
            return render_template('previred/formulario.html')
        finally:
            for p in rutas_guardadas:
                if os.path.exists(p):
                    os.remove(p)

        return send_file(ruta_salida, as_attachment=True, download_name='previred_consolidado.xlsx')

    return render_template('previred/formulario.html')
