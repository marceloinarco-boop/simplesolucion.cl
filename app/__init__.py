import os
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager
from authlib.integrations.flask_client import OAuth

db = SQLAlchemy()
login_manager = LoginManager()
oauth = OAuth()

def create_app():
    app = Flask(__name__)
    app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'cambia-esta-clave-en-produccion')
    basedir = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
    app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get(
        'DATABASE_URL', f"sqlite:///{os.path.join(basedir, 'app.db')}"
    )
    app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
    app.config['UPLOAD_FOLDER'] = os.path.join(basedir, 'uploads')
    app.config['OUTPUT_FOLDER'] = os.path.join(basedir, 'outputs')
    app.config['GOOGLE_CLIENT_ID'] = os.environ.get('GOOGLE_CLIENT_ID', '')
    app.config['GOOGLE_CLIENT_SECRET'] = os.environ.get('GOOGLE_CLIENT_SECRET', '')

    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

    db.init_app(app)
    login_manager.init_app(app)
    login_manager.login_view = 'auth.login'
    login_manager.login_message = 'Por favor inicia sesión para continuar.'
    oauth.init_app(app)

    if app.config['GOOGLE_CLIENT_ID']:
        oauth.register(
            name='google',
            client_id=app.config['GOOGLE_CLIENT_ID'],
            client_secret=app.config['GOOGLE_CLIENT_SECRET'],
            server_metadata_url='https://accounts.google.com/.well-known/openid-configuration',
            client_kwargs={'scope': 'openid email profile'},
        )

    from .models import User

    @login_manager.user_loader
    def load_user(user_id):
        return User.query.get(int(user_id))

    from .auth import auth_bp
    from .main import main_bp
    from .liquidaciones import liquidaciones_bp
    from .lre import lre_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(main_bp)
    app.register_blueprint(liquidaciones_bp, url_prefix='/liquidaciones')
    app.register_blueprint(lre_bp, url_prefix='/lre')

    with app.app_context():
        db.create_all()

    # Redirigir las carpetas de exportacion de los modulos heredados
    # hacia una carpeta segura del servidor (en vez del Escritorio local).
    from .logic import constantes as _constantes
    from .logic import generador_excel as _generador_excel

    def _export_dir_servidor(anio=None, mes=None, obra=None):
        return app.config['OUTPUT_FOLDER']

    def _export_dir_servidor_masiva(anio=None, mes=None):
        return app.config['OUTPUT_FOLDER']

    _constantes._get_directorio_exportacion = _export_dir_servidor
    _constantes._get_directorio_exportacion_masiva = _export_dir_servidor_masiva
    _generador_excel._get_directorio_exportacion = _export_dir_servidor

    return app
