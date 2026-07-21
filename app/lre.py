import os
import uuid
from flask import Blueprint, render_template, request, send_file, current_app, flash
from flask_login import login_required

from .logic.lre_engine import leer_previred, leer_liquidaciones, consolidar, construir_fila_lre, EMPRESA
import pandas as pd

lre_bp = Blueprint('lre', __name__)


@lre_bp.route('/', methods=['GET', 'POST'])
@login_required
def formulario():
    if request.method == 'POST':
        archivo_previred = request.files.get('archivo_previred')
        archivo_liquidaciones = request.files.get('archivo_liquidaciones')
        mes = request.form.get('mes', '').zfill(2)
        anio = request.form.get('anio', '').strip()

        if not archivo_previred or not archivo_liquidaciones:
            flash('Debes subir ambos archivos (Previred y Liquidaciones).', 'error')
            return render_template('lre/formulario.html')

        if not (mes.isdigit() and len(mes) == 2 and anio.isdigit() and len(anio) == 4):
            flash('Periodo inválido. Usa mes (MM) y año (AAAA).', 'error')
            return render_template('lre/formulario.html')

        upload_dir = current_app.config['UPLOAD_FOLDER']
        token = uuid.uuid4().hex[:8]
        ruta_previred = os.path.join(upload_dir, f'{token}_previred_{archivo_previred.filename}')
        ruta_liq = os.path.join(upload_dir, f'{token}_liq_{archivo_liquidaciones.filename}')
        archivo_previred.save(ruta_previred)
        archivo_liquidaciones.save(ruta_liq)

        try:
            previred = leer_previred(ruta_previred)
            liquidaciones = leer_liquidaciones(ruta_liq)
            trabajadores = consolidar(previred, liquidaciones)
            if not trabajadores:
                flash('No se encontraron trabajadores al cruzar ambos archivos. Revisa los RUT.', 'error')
                return render_template('lre/formulario.html')

            filas = [construir_fila_lre(t) for t in trabajadores]
            df = pd.DataFrame(filas)

            nombre = f"{EMPRESA['rut_archivo']}_{anio}{mes}.csv"
            ruta_salida = os.path.join(current_app.config['OUTPUT_FOLDER'], f'{token}_{nombre}')
            df.to_csv(ruta_salida, index=False, encoding='latin-1', sep=';')
        except Exception as e:
            flash(f'Error al generar el LRE: {e}', 'error')
            return render_template('lre/formulario.html')
        finally:
            for p in (ruta_previred, ruta_liq):
                if os.path.exists(p):
                    os.remove(p)

        return send_file(ruta_salida, as_attachment=True, download_name=nombre)

    return render_template('lre/formulario.html')
