from flask import Blueprint, render_template
from flask_login import login_required

main_bp = Blueprint('main', __name__)


@main_bp.route('/')
@login_required
def dashboard():
    herramientas = [
        {
            'nombre': 'Liquidaciones',
            'descripcion': 'Calcula liquidaciones de sueldo y genera PDF/Excel.',
            'url': 'liquidaciones.formulario',
        },
        {
            'nombre': 'Generador LRE',
            'descripcion': 'Genera el Libro de Remuneraciones Electrónico (CSV) desde tus archivos Previred.',
            'url': 'lre.formulario',
        },
    ]
    return render_template('dashboard.html', herramientas=herramientas)
