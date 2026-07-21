"""
====================================================
  CONSTANTES LEGALES - CHILE 2026
====================================================
ParÃ¡metros actualizados a la normativa vigente.
Fuentes: DT, SII, Superintendencia de Pensiones.
Actualizar estos valores cada vez que cambie la ley.
"""

# ââââââââââââââââââââââââââââââââââââââââââââââââââ
#  INGRESOS MÃNIMOS
# ââââââââââââââââââââââââââââââââââââââââââââââââââ

# Ingreso MÃ­nimo Mensual (IMM) vigente 2025
IMM = 510_000  # $ 510.000 CLP (actualizar segÃºn decreto)

# Tope Imponible = 81,6 UF â usar equivalencia en pesos aprox.
# Se usa como tope para cÃ¡lculo de AFP y Salud
TOPE_IMPONIBLE_UF = 81.6

# Valor UF referencial (actualizar diariamente o usar API)
VALOR_UF = 39_000  # Valor UF aproximado CLP (actualizar mensualmente)

TOPE_IMPONIBLE = TOPE_IMPONIBLE_UF * VALOR_UF  # ~$3.182.400

# ââââââââââââââââââââââââââââââââââââââââââââââââââ
#  GRATIFICACIÃN LEGAL (Art. 50 CÃ³digo del Trabajo)
# ââââââââââââââââââââââââââââââââââââââââââââââââââ

PORCENTAJE_GRATIFICACION = 0.25          # 25% del sueldo base mensual
TOPE_GRATIFICACION_ANUAL = 4.75 * IMM   # 4,75 IMM anuales
TOPE_GRATIFICACION_MENSUAL = TOPE_GRATIFICACION_ANUAL / 12  # ~$201.875

# ââââââââââââââââââââââââââââââââââââââââââââââââââ
#  HORAS EXTRAS
# ââââââââââââââââââââââââââââââââââââââââââââââââââ

HORAS_SEMANALES_ORDINARIAS = 45   # Jornada mÃ¡xima legal
RECARGO_HORA_EXTRA = 0.50         # 50% sobre valor hora ordinaria

# ââââââââââââââââââââââââââââââââââââââââââââââââââ
#  AFP - PORCENTAJES COTIZACIÃN OBLIGATORIA 2025
#  (solo cargo trabajador, excluye SIS que paga empleador)
# ââââââââââââââââââââââââââââââââââââââââââââââââââ

AFPS = {
    "Capital":   {"tasa": 0.1144, "nombre": "AFP Capital"},
    "Cuprum":    {"tasa": 0.1144, "nombre": "AFP Cuprum"},
    "Habitat":   {"tasa": 0.1127, "nombre": "AFP HÃ¡bitat"},
    "Modelo":    {"tasa": 0.1058, "nombre": "AFP Modelo"},
    "PlanVital": {"tasa": 0.1116, "nombre": "AFP PlanVital"},
    "Provida":   {"tasa": 0.1145, "nombre": "AFP Provida"},  # actualizado 2026
    "Uno":       {"tasa": 0.1046, "nombre": "AFP Uno"},      # actualizado 2026
}

# ââââââââââââââââââââââââââââââââââââââââââââââââââ
#  SALUD
# ââââââââââââââââââââââââââââââââââââââââââââââââââ

TASA_FONASA = 0.07  # 7% cotizaciÃ³n obligatoria de salud (sin cambio en 2026)

# Desde febrero 2026 (Ley 21.796 + Circular SUSESO NÂ°3912):
# Si el empleador estÃ¡ afiliado a una CCAF, el 7% se distribuye asÃ­:
#   - 4,2% â CCAF
#   - 2,8% â FONASA
# Si NO hay CCAF, el 7% va Ã­ntegro a FONASA.
# El descuento al trabajador sigue siendo 7% en ambos casos.
TASA_FONASA_CON_CCAF    = 0.028  # 2,8% a FONASA cuando hay CCAF
TASA_CCAF               = 0.042  # 4,2% a CCAF (solo empleadores con CCAF)

# ââââââââââââââââââââââââââââââââââââââââââââââââââ
#  SEGURO DE CESANTÃA (Ley 19.728)
# ââââââââââââââââââââââââââââââââââââââââââââââââââ

SEGURO_CESANTIA_INDEFINIDO = 0.006   # 0.6% cargo trabajador contrato indefinido
SEGURO_CESANTIA_PLAZO_FIJO = 0.0    # 0% cargo trabajador contrato plazo fijo

# ââââââââââââââââââââââââââââââââââââââââââââââââââ
#  IMPUESTO ÃNICO DE SEGUNDA CATEGORÃA - TABLA 2025
#  Base: UTM mensual vigente
#  Fuente: SII Chile - Tabla mensual progresiva
# ââââââââââââââââââââââââââââââââââââââââââââââââââ

# Valor UTM mensual vigente (actualizar mensualmente desde SII)
VALOR_UTM = 68_306  # $ 68.306 CLP (actualizar cada mes)

# Tabla de tramos en UTM con factor y cantidad a rebajar (en UTM)
# Formato: (lÃ­mite_superior_utm, factor, cantidad_a_rebajar_utm)
# El Ãºltimo tramo no tiene lÃ­mite superior â float('inf')
TABLA_IMPUESTO_SEGUNDA_CATEGORIA = [
    # (LÃ­mite superior en UTM, Factor %, Cantidad a Rebajar en UTM)
    (13.5,    0.000, 0.000),   # Exento
    (30.0,    0.040, 0.540),   # 4%
    (50.0,    0.080, 1.740),   # 8%
    (70.0,    0.135, 4.490),   # 13.5%
    (90.0,    0.230, 11.140),  # 23%
    (120.0,   0.304, 17.800),  # 30.4%
    (150.0,   0.350, 23.320),  # 35%
    (float('inf'), 0.400, 30.820),  # 40%
]

# ââââââââââââââââââââââââââââââââââââââââââââââââââ
#  ASIGNACIÃN FAMILIAR (SUF) - Tramos DFL NÂ°150
#  Vigente desde: actualizar segÃºn decreto
# ââââââââââââââââââââââââââââââââââââââââââââââââââ

# Montos por carga familiar segÃºn renta del trabajador
# Formato: (renta_tope, monto_por_carga)
ASIGNACION_FAMILIAR_TRAMOS = [
    (459_357,  15_407),   # Tramo A: renta â¤ $459.357 â $15.407 por carga
    (675_995,   9_009),   # Tramo B: renta â¤ $675.995 â $9.009 por carga
    (1_045_163, 2_850),   # Tramo C: renta â¤ $1.045.163 â $2.850 por carga
    (float('inf'), 0),    # Tramo D: renta > $1.045.163 â $0 (no tiene derecho)
]

# ââââââââââââââââââââââââââââââââââââââââââââââââââ
#  TIPOS DE CONTRATO
# ââââââââââââââââââââââââââââââââââââââââââââââââââ

TIPOS_CONTRATO = ["Indefinido", "Plazo Fijo"]

# ââââââââââââââââââââââââââââââââââââââââââââââââââ
#  CONFIGURACIÃN EXPORTACIÃN PDF
# ââââââââââââââââââââââââââââââââââââââââââââââââââ

DIRECTORIO_EXPORTACION = "exports"  # Mantenido por compatibilidad; usar _get_directorio_exportacion()

import os as _os
from datetime import datetime as _datetime

def _get_escritorio() -> str:
    """
    Retorna la ruta real del Escritorio del usuario actual.
    Funciona en Windows (incluso con OneDrive/carpetas redirigidas),
    macOS y Linux.
    """
    import sys as _sys

    # ââ Windows: leer ruta real desde el registro ââââââââââââââââââ
    if _sys.platform == 'win32':
        try:
            import winreg as _winreg
            _key = _winreg.OpenKey(
                _winreg.HKEY_CURRENT_USER,
                r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
            )
            _desktop, _ = _winreg.QueryValueEx(_key, 'Desktop')
            _winreg.CloseKey(_key)
            if _desktop and _os.path.isdir(_desktop):
                return _desktop
        except Exception:
            pass

        # Fallback Windows: Desktop / Escritorio bajo el perfil del usuario
        for _nombre in ('Desktop', 'Escritorio', 'OneDrive\\Escritorio',
                        'OneDrive\\Desktop'):
            _ruta = _os.path.join(_os.path.expanduser('~'), _nombre)
            if _os.path.isdir(_ruta):
                return _ruta

    # ââ macOS ââââââââââââââââââââââââââââââââââââââââââââââââââââââ
    _mac = _os.path.join(_os.path.expanduser('~'), 'Desktop')
    if _os.path.isdir(_mac):
        return _mac

    # ââ Linux ââââââââââââââââââââââââââââââââââââââââââââââââââââââ
    for _nombre in ('Escritorio', 'Desktop'):
        _ruta = _os.path.join(_os.path.expanduser('~'), _nombre)
        if _os.path.isdir(_ruta):
            return _ruta

    # Ãltimo recurso: carpeta home del usuario
    return _os.path.expanduser('~')


def _get_directorio_exportacion(anio=None, mes=None, obra=None):
    """
    Retorna la ruta de exportaciÃ³n individual:
    Escritorio/Liquidaciones/{obra}/{aÃ±o}/{mes:02d} - {Nombre Mes}/

    Si no hay obra, usa "Sin Obra" como carpeta.
    """
    _now = _datetime.now()
    _anio = anio or _now.year
    _mes  = mes  or _now.month
    _nombres_mes = ['', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
                    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    _nombre_mes = _nombres_mes[_mes] if 1 <= _mes <= 12 else str(_mes)
    _carpeta_mes = f"{_mes:02d} - {_nombre_mes}"

    # Limpiar nombre de obra para usarlo como carpeta
    if obra:
        # Tomar solo la primera obra si hay varias (separadas por " / ")
        _obra_limpia = obra.split(' / ')[0].strip()
        # Eliminar caracteres invÃ¡lidos en nombres de carpeta
        import re as _re
        _obra_limpia = _re.sub(r'[<>:"/\\|?*]', '', _obra_limpia).strip()
        if not _obra_limpia:
            _obra_limpia = "Sin Obra"
    else:
        _obra_limpia = "Sin Obra"

    return _os.path.join(_get_escritorio(), "Liquidaciones", _obra_limpia, str(_anio), _carpeta_mes)


def _get_directorio_exportacion_masiva(anio=None, mes=None):
    """
    Retorna la ruta para descargas masivas:
    Escritorio/Liquidaciones/Masivos/{aÃ±o}/{mes:02d} - {Nombre Mes}/
    """
    _now = _datetime.now()
    _anio = anio or _now.year
    _mes  = mes  or _now.month
    _nombres_mes = ['', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
                    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    _nombre_mes = _nombres_mes[_mes] if 1 <= _mes <= 12 else str(_mes)
    _carpeta_mes = f"{_mes:02d} - {_nombre_mes}"
    return _os.path.join(_get_escritorio(), "Liquidaciones", "Masivos", str(_anio), _carpeta_mes)
