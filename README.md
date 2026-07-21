# Panel de Herramientas (Liquidaciones + Generador LRE)

App web en Flask con login (Google o correo/contraseña) que reemplaza los
programas de escritorio. Corre en un servidor con Python (NO en el hosting
compartido de WordPress — ver sección "Cómo conectar con WordPress" abajo).

## Contenido
- **Liquidaciones**: calcula liquidaciones de sueldo chilenas y genera PDF/Excel.
- **Generador LRE**: sube tu Previred + Liquidaciones (Excel) y descarga el CSV
  del Libro de Remuneraciones Electrónico.

## ⚠️ Importante: "Previred a Excel" no está incluido
El zip que subiste para esa herramienta solo traía el instalador de Windows,
no el código fuente del programa (el .exe ya compilado no vino en el zip).
Por lo tanto no pude migrarla a la web. Si me pasas el código fuente (carpeta
del proyecto en Python, no el .exe), la agrego con el mismo criterio que las
otras dos.

## 1. Probar en tu computador (opcional, antes de subir a un servidor)

```bash
cd webapp
python3 -m venv venv
source venv/bin/activate        # En Windows: venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env            # y edita los valores
python run.py
```

Abre http://localhost:5000

## 2. Configurar login con Google (opcional, si no lo usas basta con correo/contraseña)

1. Ve a https://console.cloud.google.com/apis/credentials
2. Crea un proyecto (o usa uno existente) → "Crear credenciales" → "ID de cliente de OAuth"
3. Tipo de aplicación: "Aplicación web"
4. En "URI de redirección autorizados" agrega:
   - `http://localhost:5000/login/google/callback` (para probar local)
   - `https://TU-DOMINIO-DE-LA-APP/login/google/callback` (para producción)
5. Copia el "ID de cliente" y "Secreto de cliente" al archivo `.env`:
   ```
   GOOGLE_CLIENT_ID=xxxxxxxx.apps.googleusercontent.com
   GOOGLE_CLIENT_SECRET=xxxxxxxx
   ```

Si no configuras Google, la app funciona igual solo con correo/contraseña.

## 3. Desplegar en un servicio con Python (recomendado: Render.com, gratis para empezar)

1. Sube esta carpeta `webapp/` a un repositorio de GitHub.
2. Entra a https://render.com → "New +" → "Web Service" → conecta tu repo.
3. Configuración:
   - **Build command:** `pip install -r requirements.txt`
   - **Start command:** `gunicorn run:app`
4. En "Environment", agrega las variables:
   - `SECRET_KEY` (cualquier texto largo y aleatorio)
   - `GOOGLE_CLIENT_ID` y `GOOGLE_CLIENT_SECRET` (si usas Google login)
5. Render te da una URL tipo `https://tu-app.onrender.com` — esa es tu app funcionando.
6. Actualiza en Google Cloud Console el URI de redirección con esa URL real.

(Alternativas equivalentes: Railway.app, PythonAnywhere, o un VPS con Gunicorn + Nginx.)

## 4. Cómo conectar con tu WordPress (hosting compartido)

Tu hosting de WordPress **no puede ejecutar esta app** (es PHP, no Python).
Lo que sí puedes hacer:

- **Opción simple:** agrega un botón/menú en WordPress que enlace a
  `https://tu-app.onrender.com` (se abre la app en su propia página).
- **Opción integrada:** si tu dominio principal es `tuempresa.cl` (en WordPress),
  puedes crear un subdominio `app.tuempresa.cl` apuntando (CNAME) a la URL de
  Render/Railway, para que se vea como parte de tu sitio.
- **Opción visual:** insertar un `<iframe src="https://tu-app.onrender.com">`
  en una página de WordPress (funciona, pero es menos recomendable para login).

## 5. Datos de prueba realizados

Antes de entregarte esto, probé de punta a punta: registro de usuario,
cálculo de una liquidación, descarga de PDF y descarga de Excel — los tres
pasos funcionaron correctamente. El módulo de Generador LRE importa y carga
sin errores; pruébalo tú con tus archivos Previred reales, ya que no tenía
un archivo de ejemplo para probarlo end-to-end.
