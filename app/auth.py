from flask import Blueprint, render_template, redirect, url_for, flash, request, current_app
from flask_login import login_user, logout_user, login_required, current_user
from . import db, oauth
from .models import User

auth_bp = Blueprint('auth', __name__)


@auth_bp.route('/registro', methods=['GET', 'POST'])
def registro():
    if current_user.is_authenticated:
        return redirect(url_for('main.dashboard'))
    if request.method == 'POST':
        email = request.form.get('email', '').strip().lower()
        nombre = request.form.get('nombre', '').strip()
        password = request.form.get('password', '')
        if not email or not password:
            flash('Correo y contraseña son obligatorios.', 'error')
            return render_template('auth/registro.html')
        if User.query.filter_by(email=email).first():
            flash('Ya existe una cuenta con ese correo.', 'error')
            return render_template('auth/registro.html')
        user = User(email=email, nombre=nombre or email.split('@')[0])
        user.set_password(password)
        db.session.add(user)
        db.session.commit()
        login_user(user)
        flash('Cuenta creada correctamente.', 'success')
        return redirect(url_for('main.dashboard'))
    return render_template('auth/registro.html')


@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('main.dashboard'))
    if request.method == 'POST':
        email = request.form.get('email', '').strip().lower()
        password = request.form.get('password', '')
        user = User.query.filter_by(email=email).first()
        if user and user.check_password(password):
            login_user(user)
            return redirect(url_for('main.dashboard'))
        flash('Correo o contraseña incorrectos.', 'error')
    google_disponible = bool(current_app.config.get('GOOGLE_CLIENT_ID'))
    return render_template('auth/login.html', google_disponible=google_disponible)


@auth_bp.route('/login/google')
def login_google():
    redirect_uri = url_for('auth.auth_google_callback', _external=True)
    return oauth.google.authorize_redirect(redirect_uri)


@auth_bp.route('/login/google/callback')
def auth_google_callback():
    token = oauth.google.authorize_access_token()
    userinfo = token.get('userinfo') or oauth.google.userinfo()
    email = userinfo['email'].lower()
    google_id = userinfo['sub']
    user = User.query.filter_by(email=email).first()
    if not user:
        user = User(email=email, nombre=userinfo.get('name', email.split('@')[0]), google_id=google_id)
        db.session.add(user)
        db.session.commit()
    elif not user.google_id:
        user.google_id = google_id
        db.session.commit()
    login_user(user)
    return redirect(url_for('main.dashboard'))


@auth_bp.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('auth.login'))
