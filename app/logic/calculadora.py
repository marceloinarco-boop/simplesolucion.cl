"""
====================================================
  MOTOR DE CÃLCULO - LIQUIDACIÃN DE SUELDO CHILE
====================================================
Contiene toda la lÃ³gica matemÃ¡tica y legal para
calcular correctamente una liquidaciÃ³n de sueldo
segÃºn la normativa chilena vigente 2025.
"""

from .constantes import (
    IMM, TOPE_IMPONIBLE, VALOR_UTM, VALOR_UF,
    PORCENTAJE_GRATIFICACION, TOPE_GRATIFICACION_MENSUAL,
    RECARGO_HORA_EXTRA, HORAS_SEMANALES_ORDINARIAS,
    AFPS, TASA_FONASA,
    SEGURO_CESANTIA_INDEFINIDO, SEGURO_CESANTIA_PLAZO_FIJO,
    TABLA_IMPUESTO_SEGUNDA_CATEGORIA,
    ASIGNACION_FAMILIAR_TRAMOS,
)


class CalculadoraLiquidacion:
    """
    Realiza todos los cÃ¡lculos de una liquidaciÃ³n de sueldo chilena.
    Recibe un diccionario de datos y retorna un diccionario con
    todos los montos calculados.
    """

    def __init__(self, datos: dict):
        """
        Inicializa con los datos del formulario.
        datos = {
            'sueldo_base': float,
            'tipo_contrato': str,       # 'Indefinido' | 'Plazo Fijo'
            'gratificacion_auto': bool,
            'gratificacion_manual': float,
            'dias_trabajados': int,     # 30 = mes completo
            'horas_extra': float,
            'bono_imponible': float,
            'colacion': float,
            'movilizacion': float,
            'viaticos': float,
            'cargas_familiares': int,
            'afp': str,                 # clave en AFPS
            'tipo_salud': str,          # 'FONASA' | 'ISAPRE'
            'monto_isapre': float,      # en pesos si tipo_salud == 'ISAPRE'
            'dias_mes': int,            # Normalmente 30
        }
        """
        self.d = datos
        self.resultado = {}

    # ââââââââââââââââââââââââââââââââââââââââââââââ
    #  MÃTODOS AUXILIARES
    # ââââââââââââââââââââââââââââââââââââââââââââââ

    def _valor_hora_ordinaria(self) -> float:
        """Calcula el valor de la hora ordinaria diurna."""
        sueldo_base = float(self.d.get('sueldo_base', 0))
        # FÃ³rmula legal: sueldo_base / (dÃ­as_mes * horas_diarias)
        # Horas diarias = horas_semanales / 5 dÃ­as (o 6 segÃºn jornada)
        # Simplificado: sueldo / 30 / 8 hrs o usando fÃ³rmula oficial
        # FÃ³rmula DT: valor hora = sueldo_base * 12 / (52 * jornada_semanal)
        jornada = HORAS_SEMANALES_ORDINARIAS
        valor_hora = (sueldo_base * 12) / (52 * jornada)
        return valor_hora

    def _calcular_gratificacion(self, sueldo_base: float) -> float:
        """
        Calcula la gratificaciÃ³n legal mensual.
        Art. 50 CT: 25% del sueldo base mensual con tope de 4,75 IMM anual.
        Se paga mensualmente = tope_anual / 12
        """
        if self.d.get('gratificacion_auto', True):
            gratificacion = sueldo_base * PORCENTAJE_GRATIFICACION
            return min(gratificacion, TOPE_GRATIFICACION_MENSUAL)
        else:
            return float(self.d.get('gratificacion_manual', 0))

    def _calcular_horas_extra(self, sueldo_base: float, horas_extra: float) -> float:
        """
        Calcula el monto por horas extraordinarias.
        Recargo legal: 50% sobre valor hora ordinaria.
        """
        if horas_extra <= 0:
            return 0.0
        valor_hora = self._valor_hora_ordinaria()
        valor_hora_extra = valor_hora * (1 + RECARGO_HORA_EXTRA)
        return valor_hora_extra * horas_extra

    def _calcular_asignacion_familiar(self, renta_bruta: float, cargas: int) -> float:
        """
        Calcula la asignaciÃ³n familiar segÃºn tramos de renta.
        SUF pagado por empleador, no imponible, no tributable.
        """
        if cargas <= 0:
            return 0.0
        for tope, monto in ASIGNACION_FAMILIAR_TRAMOS:
            if renta_bruta <= tope:
                return monto * cargas
        return 0.0

    def _aplicar_proporcionalidad(self, monto: float, dias_trabajados: int) -> float:
        """
        Aplica proporcionalidad si el trabajador no trabajÃ³ el mes completo.
        Base: 30 dÃ­as.
        """
        dias_mes = int(self.d.get('dias_mes', 30))
        if dias_trabajados >= dias_mes:
            return monto
        return (monto / dias_mes) * dias_trabajados

    # ââââââââââââââââââââââââââââââââââââââââââââââ
    #  CÃLCULO PRINCIPAL
    # ââââââââââââââââââââââââââââââââââââââââââââââ

    def calcular(self) -> dict:
        """
        Ejecuta el cÃ¡lculo completo de la liquidaciÃ³n.
        Retorna diccionario con todos los valores desglosados.
        """
        d = self.d

        sueldo_base   = float(d.get('sueldo_base', 0))
        dias_trab     = int(d.get('dias_trabajados', 30))
        horas_extra   = float(d.get('horas_extra', 0))
        bono_imponible = float(d.get('bono_imponible', 0))
        colacion      = float(d.get('colacion', 0))
        movilizacion  = float(d.get('movilizacion', 0))
        viaticos      = float(d.get('viaticos', 0))
        cargas        = int(d.get('cargas_familiares', 0))
        tipo_contrato = d.get('tipo_contrato', 'Indefinido')
        afp_key       = d.get('afp', 'Capital')
        tipo_salud    = d.get('tipo_salud', 'FONASA')
        monto_isapre  = float(d.get('monto_isapre', 0))

        # ââ 1. HABERES IMPONIBLES ââââââââââââââââââ
        sueldo_base_prop = self._aplicar_proporcionalidad(sueldo_base, dias_trab)
        gratificacion    = self._calcular_gratificacion(sueldo_base_prop)
        monto_horas_extra = self._calcular_horas_extra(sueldo_base, horas_extra)

        total_imponible = (
            sueldo_base_prop
            + gratificacion
            + monto_horas_extra
            + bono_imponible
        )

        # Aplicar tope imponible para cÃ¡lculo previsional
        base_previsional = min(total_imponible, TOPE_IMPONIBLE)

        # ââ 2. HABERES NO IMPONIBLES ââââââââââââââ
        asig_familiar = self._calcular_asignacion_familiar(total_imponible, cargas)

        total_no_imponible = (
            colacion
            + movilizacion
            + viaticos
            + asig_familiar
        )

        # ââ 3. TOTAL HABERES (BRUTO) âââââââââââââââ
        total_haberes = total_imponible + total_no_imponible

        # ââ 4. DESCUENTOS PREVISIONALES âââââââââââ
        # AFP
        tasa_afp = AFPS.get(afp_key, AFPS['Capital'])['tasa']
        descuento_afp = base_previsional * tasa_afp

        # Salud
        if tipo_salud == 'FONASA':
            descuento_salud = base_previsional * TASA_FONASA
            nombre_salud    = 'FONASA (7%)'
        else:
            # ISAPRE: el monto ingresado en pesos directamente
            descuento_salud = monto_isapre
            nombre_salud    = 'ISAPRE'

        # Seguro de CesantÃ­a (solo cargo trabajador en contrato indefinido)
        if tipo_contrato == 'Indefinido':
            descuento_cesantia = base_previsional * SEGURO_CESANTIA_INDEFINIDO
        else:
            descuento_cesantia = 0.0

        total_descuentos_prev = descuento_afp + descuento_salud + descuento_cesantia

        # ââ 5. BASE IMPONIBLE IMPUESTO ââââââââââââ
        # Renta imponible para IUSC = Total Imponible - Descuentos Previsionales
        base_impuesto = total_imponible - total_descuentos_prev

        # ââ 6. IMPUESTO ÃNICO SEGUNDA CATEGORÃA ââ
        impuesto_usc = self._calcular_impuesto(base_impuesto)

        # ââ 7. TOTAL DESCUENTOS âââââââââââââââââââ
        total_descuentos = total_descuentos_prev + impuesto_usc

        # ââ 8. SUELDO LÃQUIDO âââââââââââââââââââââ
        sueldo_liquido = total_haberes - total_descuentos

        # ââ ARMAR RESULTADO âââââââââââââââââââââââ
        self.resultado = {
            # Haberes Imponibles
            'sueldo_base':         round(sueldo_base_prop, 0),
            'gratificacion':       round(gratificacion, 0),
            'horas_extra_monto':   round(monto_horas_extra, 0),
            'bono_imponible':      round(bono_imponible, 0),
            'total_imponible':     round(total_imponible, 0),
            'base_previsional':    round(base_previsional, 0),

            # Haberes No Imponibles
            'colacion':            round(colacion, 0),
            'movilizacion':        round(movilizacion, 0),
            'viaticos':            round(viaticos, 0),
            'asig_familiar':       round(asig_familiar, 0),
            'total_no_imponible':  round(total_no_imponible, 0),

            # Total Haberes
            'total_haberes':       round(total_haberes, 0),

            # Descuentos Previsionales â nombres originales
            'afp_nombre':          AFPS.get(afp_key, AFPS['Capital'])['nombre'],
            'afp_tasa':            tasa_afp,
            'descuento_afp':       round(descuento_afp, 0),
            'nombre_salud':        nombre_salud,
            'descuento_salud':     round(descuento_salud, 0),
            'descuento_cesantia':  round(descuento_cesantia, 0),
            'total_desc_prev':     round(total_descuentos_prev, 0),

            # Aliases que usa interfaz.py / preview / generadores
            'afp_monto':           round(descuento_afp, 0),
            'tasa_afp':            tasa_afp,
            'salud_monto':         round(descuento_salud, 0),
            'cesantia_trabajador': round(descuento_cesantia, 0),
            'base_tributable':     round(base_impuesto, 0),
            'impuesto_segunda_cat': round(impuesto_usc, 0),

            # Impuesto
            'base_impuesto':       round(base_impuesto, 0),
            'impuesto_usc':        round(impuesto_usc, 0),

            # Totales
            'total_descuentos':    round(total_descuentos, 0),
            'sueldo_liquido':      round(sueldo_liquido, 0),

            # Metadata
            'valor_utm':           VALOR_UTM,
            'valor_uf':            VALOR_UF,
            'imm':                 IMM,
        }

        return self.resultado

    def _calcular_impuesto(self, renta_imponible: float) -> float:
        """
        Calcula el Impuesto Ãnico de Segunda CategorÃ­a (IUSC).
        Tabla progresiva mensual del SII expresada en UTM.
        
        Procedimiento:
        1. Convertir renta imponible a UTM
        2. Localizar el tramo correspondiente
        3. Aplicar: (renta_utm * factor) - cantidad_a_rebajar_utm
        4. Convertir resultado de vuelta a pesos
        """
        if renta_imponible <= 0:
            return 0.0

        # Convertir a UTM
        renta_utm = renta_imponible / VALOR_UTM

        impuesto_utm = 0.0

        for limite_utm, factor, rebaja_utm in TABLA_IMPUESTO_SEGUNDA_CATEGORIA:
            if renta_utm <= limite_utm:
                impuesto_utm = (renta_utm * factor) - rebaja_utm
                break

        # El impuesto no puede ser negativo
        impuesto_utm = max(0.0, impuesto_utm)

        # Convertir de UTM a pesos
        return impuesto_utm * VALOR_UTM


# ââââââââââââââââââââââââââââââââââââââââââââââââââ
#  FUNCIÃN DE CONVENIENCIA
# ââââââââââââââââââââââââââââââââââââââââââââââââââ

def calcular_liquidacion(datos: dict) -> dict:
    """
    FunciÃ³n de conveniencia para calcular una liquidaciÃ³n.
    Retorna el diccionario con todos los resultados.
    """
    calc = CalculadoraLiquidacion(datos)
    return calc.calcular()
