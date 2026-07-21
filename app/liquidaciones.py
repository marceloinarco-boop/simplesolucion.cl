import os
from flask import Blueprint, render_template, request, send_file, current_app, flash
from flask_login import login_required, current_user

from .logic.calculadora import calcular_liquidacion
from .logic.generador_pdf import generar_pdf_liquidacion
from .logic.generador_excel import generar_excel_liquidacion

liquidaciones_bp = Blueprint('liquidaciones', __name__)

CAMPOS_NUMERICOS = [
    'sueldo_base', 'gratificacion_manual', 'dias_trabajados', 'horas_extra',
    'bono_imponible', 'colacion', 'movilizacion', 'viaticos',
    'cargas_familiares', 'monto_isapre', 'dias_mes',
]


def _parse_form(form):
    datos = {}
    for campo in CAMPOS_NUMERICOS:
        valor = form.get(campo, '0').strip() or '0'
        try:
            datos[campo] = float(valor)
        except ValueError:
            datos[campo] = 0.0
    datos['tipo_contrato'] = form.get('tipo_contrato', 'Indefinido')
    datos['gratificacion_auto'] = form.get('gratificacion_auto') == 'on'
    datos['afp'] = form.get('afp', 'Habitat')
    datos['tipo_salud'] = form.get('tipo_salud', 'FONASA')
    return datos


@liquidaciones_bp.route('/', methods=['GET', 'POST'])
@login_required
def formulario():
    resultado = None
    datos = None
    trabajador = None
    if request.method == 'POST':
        datos = _parse_form(request.form)
        trabajador = {
            'nombre': request.form.get('nombre_trabajador', '').strip(),
            'rut': request.form.get('rut_trabajador', '').strip(),
            'cargo': request.form.get('cargo', '').strip(),
        }
        try:
            resultado = calcular_liquidacion(datos)
        except Exception as e:
            flash(f'Error al calcular: {e}', 'error')
    return render_template('liquidaciones/formulario.html', resultado=resultado,
                            datos=datos, trabajador=trabajador,
                            afps=['Habitat', 'Capital', 'Cuprum', 'Provida', 'Planvital', 'Modelo', 'Uno'])


@liquidaciones_bp.route('/descargar/<formato>', methods=['POST'])
@login_required
def descargar(formato):
    datos = _parse_form(request.form)
    trabajador = {
        'nombre': request.form.get('nombre_trabajador', '').strip() or 'Trabajador',
        'rut': request.form.get('rut_trabajador', '').strip(),
        'cargo': request.form.get('cargo', '').strip(),
    }
    resultado = calcular_liquidacion(datos)
    empresa = {
        'nombre': request.form.get('empresa_nombre', '') or 'Empresa',
        'rut': request.form.get('empresa_rut', '') or '00.000.000-0',
        'direccion': request.form.get('empresa_direccion', '') or 'No especificada',
        'ciudad': request.form.get('empresa_ciudad', '') or 'Santiago',
    }

    if formato == 'pdf':
        ruta = generar_pdf_liquidacion(trabajador, datos, resultado, empresa=empresa)
    elif formato == 'excel':
        ruta = generar_excel_liquidacion(trabajador, datos, resultado, empresa=empresa)
    else:
        flash('Formato no válido', 'error')
        return render_template('liquidaciones/formulario.html', resultado=resultado, datos=datos, trabajador=trabajador)

    nombre_archivo = os.path.basename(ruta)
    return send_file(ruta, as_attachment=True, download_name=nombre_archivo)
