"""
====================================================
  GENERADOR DE PDF - LIQUIDACIÃN DE SUELDO CHILE
====================================================
Genera un PDF de liquidaciÃ³n de sueldo con formato
estÃ¡ndar chileno usando ReportLab.
"""

import os
from datetime import datetime
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.units import cm, mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether
)
from reportlab.lib.colors import HexColor

from .constantes import DIRECTORIO_EXPORTACION


# ââ Colores corporativos ââââââââââââââââââââââââââ
COLOR_PRIMARIO   = HexColor('#1a3a5c')   # Azul oscuro
COLOR_SECUNDARIO = HexColor('#2e7d9e')   # Azul medio
COLOR_ACENTO     = HexColor('#e8f4f8')   # Azul muy claro (fondos)
COLOR_EXITO      = HexColor('#2e7d32')   # Verde oscuro
COLOR_GRIS       = HexColor('#f5f5f5')   # Gris claro
COLOR_BORDE      = HexColor('#cccccc')   # Gris borde
BLANCO           = colors.white
NEGRO            = colors.black


def _fmt_pesos(valor: float) -> str:
    """Formatea un valor como pesos chilenos. Ej: 1500000 â $ 1.500.000"""
    try:
        v = int(round(float(valor)))
        return f"$ {v:,.0f}".replace(',', '.')
    except (ValueError, TypeError):
        return "$ 0"


def _fmt_porcentaje(valor: float) -> str:
    """Formatea como porcentaje. Ej: 0.1144 â 11,44%"""
    return f"{valor * 100:.2f}%".replace('.', ',')


def _nombre_mes(num_mes: int) -> str:
    """Retorna el nombre del mes en espaÃ±ol."""
    meses = [
        '', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
        'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
    ]
    return meses[num_mes] if 1 <= num_mes <= 12 else str(num_mes)


class GeneradorPDF:
    """Genera el PDF de liquidaciÃ³n de sueldo."""

    def __init__(self, datos_trabajador: dict, datos_entrada: dict, resultado: dict,
                 empresa: dict = None, periodo: str = None):
        """
        datos_trabajador: {nombre, rut, tipo_contrato, ...}
        datos_entrada: inputs del formulario
        resultado: output de CalculadoraLiquidacion.calcular()
        empresa: {nombre, rut, direccion, ciudad} (opcional)
        periodo: 'YYYY-MM'
        """
        self.trabajador = datos_trabajador
        self.entrada    = datos_entrada
        self.res        = resultado
        self.empresa    = empresa or {
            'nombre': 'Empresa S.A.',
            'rut': '76.000.000-0',
            'direccion': 'DirecciÃ³n no especificada',
            'ciudad': 'Santiago',
        }
        self.periodo = periodo or datetime.now().strftime('%Y-%m')

        # Parsear perÃ­odo
        try:
            anio, mes = self.periodo.split('-')
            self.anio = int(anio)
            self.mes  = int(mes)
        except Exception:
            self.anio = datetime.now().year
            self.mes  = datetime.now().month

    def generar(self) -> str:
        """
        Genera el PDF y lo guarda en el directorio de exportaciÃ³n.
        Retorna la ruta del archivo generado.
        """
        from .constantes import _get_directorio_exportacion
        _obra_dir = self.entrada.get('nombre_obra', '')
        directorio = _get_directorio_exportacion(self.anio, self.mes, obra=_obra_dir)
        os.makedirs(directorio, exist_ok=True)

        nombre_archivo = (
            f"Liquidacion_{self.trabajador['nombre'].replace(' ', '_')}"
            f"_{self.periodo}.pdf"
        )
        ruta_pdf = os.path.join(directorio, nombre_archivo)

        # Configurar documento A4
        doc = SimpleDocTemplate(
            ruta_pdf,
            pagesize=A4,
            leftMargin=1.5 * cm,
            rightMargin=1.5 * cm,
            topMargin=1.5 * cm,
            bottomMargin=2 * cm,
            title=f"LiquidaciÃ³n de Sueldo - {self.trabajador['nombre']} - {self.periodo}",
            author="Liquidador de Sueldos Chile",
        )

        # Construir contenido
        historia = self._construir_historia()

        # Generar PDF
        doc.build(historia)
        return ruta_pdf

    def _estilos(self) -> dict:
        """Define los estilos tipogrÃ¡ficos del documento."""
        base = getSampleStyleSheet()
        estilos = {
            'titulo_empresa': ParagraphStyle(
                'titulo_empresa',
                fontSize=14, leading=16,
                fontName='Helvetica-Bold',
                textColor=COLOR_PRIMARIO,
                alignment=TA_CENTER,
            ),
            'subtitulo': ParagraphStyle(
                'subtitulo',
                fontSize=10, leading=12,
                fontName='Helvetica',
                textColor=COLOR_SECUNDARIO,
                alignment=TA_CENTER,
            ),
            'titulo_seccion': ParagraphStyle(
                'titulo_seccion',
                fontSize=9, leading=11,
                fontName='Helvetica-Bold',
                textColor=BLANCO,
            ),
            'normal': ParagraphStyle(
                'normal',
                fontSize=8, leading=10,
                fontName='Helvetica',
                textColor=NEGRO,
            ),
            'normal_bold': ParagraphStyle(
                'normal_bold',
                fontSize=8, leading=10,
                fontName='Helvetica-Bold',
                textColor=NEGRO,
            ),
            'label': ParagraphStyle(
                'label',
                fontSize=7, leading=9,
                fontName='Helvetica',
                textColor=HexColor('#555555'),
            ),
            'monto': ParagraphStyle(
                'monto',
                fontSize=8, leading=10,
                fontName='Helvetica',
                textColor=NEGRO,
                alignment=TA_RIGHT,
            ),
            'monto_bold': ParagraphStyle(
                'monto_bold',
                fontSize=9, leading=11,
                fontName='Helvetica-Bold',
                textColor=NEGRO,
                alignment=TA_RIGHT,
            ),
            'liquido': ParagraphStyle(
                'liquido',
                fontSize=13, leading=15,
                fontName='Helvetica-Bold',
                textColor=COLOR_EXITO,
                alignment=TA_RIGHT,
            ),
            'pie': ParagraphStyle(
                'pie',
                fontSize=7, leading=9,
                fontName='Helvetica',
                textColor=HexColor('#777777'),
                alignment=TA_CENTER,
            ),
        }
        return estilos

    def _construir_historia(self) -> list:
        """Construye la lista de flowables del documento."""
        s = self._estilos()
        historia = []

        # ââ ENCABEZADO ââââââââââââââââââââââââââââ
        historia += self._seccion_encabezado(s)
        historia.append(Spacer(1, 0.3 * cm))

        # ââ DATOS TRABAJADOR Y EMPRESA ââââââââââââ
        historia += self._seccion_datos_partes(s)
        historia.append(Spacer(1, 0.4 * cm))

        # ââ HABERES âââââââââââââââââââââââââââââââ
        historia += self._seccion_haberes(s)
        historia.append(Spacer(1, 0.3 * cm))

        # ââ DESCUENTOS ââââââââââââââââââââââââââââ
        historia += self._seccion_descuentos(s)
        historia.append(Spacer(1, 0.3 * cm))

        # ââ RESUMEN FINAL âââââââââââââââââââââââââ
        historia += self._seccion_resumen(s)
        historia.append(Spacer(1, 0.5 * cm))

        # ââ FIRMAS ââââââââââââââââââââââââââââââââ
        historia += self._seccion_firmas(s)
        historia.append(Spacer(1, 0.3 * cm))

        # ââ PIE DE PÃGINA âââââââââââââââââââââââââ
        historia.append(HRFlowable(width="100%", thickness=0.5, color=COLOR_BORDE))
        historia.append(Spacer(1, 0.1 * cm))
        historia.append(Paragraph(
            f"Documento generado el {datetime.now().strftime('%d/%m/%Y %H:%M')} Â· "
            f"Liquidador de Sueldos Chile Â· UTM: ${self.res.get('valor_utm', 0):,.0f} Â· "
            f"UF: ${self.res.get('valor_uf', 0):,.0f}",
            s['pie']
        ))

        return historia

    def _seccion_encabezado(self, s: dict) -> list:
        """Genera el encabezado del documento."""
        titulo_liq = Paragraph(
            f"LIQUIDACIÃN DE REMUNERACIONES",
            s['titulo_empresa']
        )
        periodo_txt = Paragraph(
            f"PerÃ­odo: {_nombre_mes(self.mes)} {self.anio}",
            s['subtitulo']
        )

        # LÃ­nea decorativa
        linea = HRFlowable(width="100%", thickness=2, color=COLOR_PRIMARIO)

        return [titulo_liq, periodo_txt, Spacer(1, 0.2 * cm), linea]

    def _seccion_datos_partes(self, s: dict) -> list:
        """Genera la secciÃ³n con datos de empresa y trabajador."""
        t = self.trabajador
        e = self.empresa

        # Calcular dÃ­as trabajados
        dias = self.entrada.get('dias_trabajados', 30)
        fecha_inicio = (t.get('fecha_inicio_contrato', '') or
                        self.entrada.get('fecha_inicio_contrato', '')).strip()
        nombre_obra = self.entrada.get('nombre_obra', '').strip()

        datos = [
            # Cabecera de la tabla
            [
                Paragraph("<b>EMPLEADOR</b>", s['normal_bold']),
                Paragraph("<b>TRABAJADOR</b>", s['normal_bold']),
            ],
            [
                Paragraph(f"RazÃ³n Social: {e['nombre']}", s['normal']),
                Paragraph(f"Nombre: {t.get('nombre', '')}", s['normal']),
            ],
            [
                Paragraph(f"RUT: {e['rut']}", s['normal']),
                Paragraph(f"RUT: {t.get('rut', '')}", s['normal']),
            ],
            [
                Paragraph(f"DirecciÃ³n: {e['direccion']}", s['normal']),
                Paragraph(f"Tipo de Contrato: {t.get('tipo_contrato', '')}", s['normal']),
            ],
            [
                Paragraph(f"Ciudad: {e['ciudad']}", s['normal']),
                Paragraph(f"DÃ­as Trabajados: {dias} dÃ­as", s['normal']),
            ],
        ]

        if fecha_inicio:
            datos.append([
                Paragraph('', s['normal']),
                Paragraph(f"Inicio de Contrato: {fecha_inicio}", s['normal']),
            ])

        if nombre_obra:
            for obra in nombre_obra.split(' / '):
                obra = obra.strip()
                if obra:
                    datos.append([
                        Paragraph('', s['normal']),
                        Paragraph(f"Obra: {obra}", s['normal_bold']),
                    ])

        tabla = Table(datos, colWidths=[9 * cm, 9 * cm])
        tabla.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), COLOR_PRIMARIO),
            ('TEXTCOLOR',  (0, 0), (-1, 0), BLANCO),
            ('FONTNAME',   (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE',   (0, 0), (-1, 0), 8),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [BLANCO, COLOR_GRIS]),
            ('GRID',   (0, 0), (-1, -1), 0.5, COLOR_BORDE),
            ('VALIGN',  (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING',    (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING',   (0, 0), (-1, -1), 6),
        ]))

        return [tabla]

    def _seccion_haberes(self, s: dict) -> list:
        """Genera la secciÃ³n de haberes (imponibles y no imponibles)."""
        r = self.res

        # ââ Cabecera de secciÃ³n ââ
        cab = Table(
            [[Paragraph("  HABERES", s['titulo_seccion'])]],
            colWidths=[18 * cm]
        )
        cab.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), COLOR_SECUNDARIO),
            ('TOPPADDING',    (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))

        # ââ Subcabecera columnas ââ
        filas = [
            [
                Paragraph("<b>Concepto</b>", s['normal_bold']),
                Paragraph("<b>Tipo</b>", s['normal_bold']),
                Paragraph("<b>Monto</b>", s['normal_bold']),
            ],
        ]

        # Haberes imponibles
        imponibles = [
            ('Sueldo Base', 'Imponible', r['sueldo_base']),
            ('GratificaciÃ³n Legal (25%)', 'Imponible', r['gratificacion']),
        ]
        if r['horas_extra_monto'] > 0:
            imponibles.append(('Horas Extraordinarias (50%)', 'Imponible', r['horas_extra_monto']))
        if r['bono_imponible'] > 0:
            imponibles.append(('Bonos Imponibles', 'Imponible', r['bono_imponible']))

        # Haberes no imponibles
        no_imponibles = []
        if r['colacion'] > 0:
            no_imponibles.append(('AsignaciÃ³n de ColaciÃ³n', 'No Imponible', r['colacion']))
        if r['movilizacion'] > 0:
            no_imponibles.append(('AsignaciÃ³n de MovilizaciÃ³n', 'No Imponible', r['movilizacion']))
        if r['viaticos'] > 0:
            no_imponibles.append(('ViÃ¡ticos', 'No Imponible', r['viaticos']))
        if r['asig_familiar'] > 0:
            no_imponibles.append(('AsignaciÃ³n Familiar (SUF)', 'No Imponible', r['asig_familiar']))

        for concepto, tipo, monto in imponibles + no_imponibles:
            filas.append([
                Paragraph(concepto, s['normal']),
                Paragraph(tipo, s['label']),
                Paragraph(_fmt_pesos(monto), s['monto']),
            ])

        # Subtotales
        filas.append([
            Paragraph("<b>Total Haberes Imponibles</b>", s['normal_bold']),
            Paragraph('', s['normal']),
            Paragraph(f"<b>{_fmt_pesos(r['total_imponible'])}</b>", s['monto_bold']),
        ])
        filas.append([
            Paragraph("<b>Total Haberes No Imponibles</b>", s['normal_bold']),
            Paragraph('', s['normal']),
            Paragraph(f"<b>{_fmt_pesos(r['total_no_imponible'])}</b>", s['monto_bold']),
        ])
        filas.append([
            Paragraph("<b>TOTAL HABERES (BRUTO)</b>", s['normal_bold']),
            Paragraph('', s['normal']),
            Paragraph(f"<b>{_fmt_pesos(r['total_haberes'])}</b>", s['monto_bold']),
        ])

        n = len(filas)
        tabla = Table(filas, colWidths=[11 * cm, 3 * cm, 4 * cm])
        tabla.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), COLOR_ACENTO),
            ('ROWBACKGROUNDS', (0, 1), (-1, n - 4), [BLANCO, COLOR_GRIS]),
            ('BACKGROUND', (0, n - 3), (-1, n - 3), HexColor('#dcedf5')),
            ('BACKGROUND', (0, n - 2), (-1, n - 2), HexColor('#dcedf5')),
            ('BACKGROUND', (0, n - 1), (-1, n - 1), COLOR_PRIMARIO),
            ('TEXTCOLOR',  (0, n - 1), (-1, n - 1), BLANCO),
            ('FONTNAME',   (0, n - 1), (-1, n - 1), 'Helvetica-Bold'),
            ('GRID',    (0, 0), (-1, -1), 0.5, COLOR_BORDE),
            ('VALIGN',  (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING',    (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING',   (0, 0), (-1, -1), 6),
        ]))

        return [cab, tabla]

    def _seccion_descuentos(self, s: dict) -> list:
        """Genera la secciÃ³n de descuentos."""
        r = self.res

        cab = Table(
            [[Paragraph("  DESCUENTOS", s['titulo_seccion'])]],
            colWidths=[18 * cm]
        )
        cab.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), HexColor('#8B0000')),
            ('TOPPADDING',    (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))

        filas = [
            [
                Paragraph("<b>Concepto</b>", s['normal_bold']),
                Paragraph("<b>Base / Tasa</b>", s['normal_bold']),
                Paragraph("<b>Monto</b>", s['normal_bold']),
            ],
            [
                Paragraph(r['afp_nombre'], s['normal']),
                Paragraph(f"{_fmt_pesos(r['base_previsional'])} Ã {_fmt_porcentaje(r['afp_tasa'])}", s['label']),
                Paragraph(_fmt_pesos(r['descuento_afp']), s['monto']),
            ],
            [
                Paragraph(r['nombre_salud'], s['normal']),
                Paragraph(f"{_fmt_pesos(r['base_previsional'])} Ã 7,00%" if r['nombre_salud'] == 'FONASA (7%)' else "Monto pactado", s['label']),
                Paragraph(_fmt_pesos(r['descuento_salud']), s['monto']),
            ],
            [
                Paragraph("Seguro de CesantÃ­a (trabajador)", s['normal']),
                Paragraph(f"{_fmt_pesos(r['base_previsional'])} Ã 0,60%" if r['descuento_cesantia'] > 0 else "Contrato plazo fijo (exento)", s['label']),
                Paragraph(_fmt_pesos(r['descuento_cesantia']), s['monto']),
            ],
            [
                Paragraph("<b>Total Descuentos Previsionales</b>", s['normal_bold']),
                Paragraph('', s['normal']),
                Paragraph(f"<b>{_fmt_pesos(r['total_desc_prev'])}</b>", s['monto_bold']),
            ],
            [
                Paragraph("Impuesto Ãnico de Segunda CategorÃ­a", s['normal']),
                Paragraph(f"Base imponible: {_fmt_pesos(r['base_impuesto'])}", s['label']),
                Paragraph(_fmt_pesos(r['impuesto_usc']), s['monto']),
            ],
            [
                Paragraph("<b>TOTAL DESCUENTOS</b>", s['normal_bold']),
                Paragraph('', s['normal']),
                Paragraph(f"<b>{_fmt_pesos(r['total_descuentos'])}</b>", s['monto_bold']),
            ],
        ]

        n = len(filas)
        tabla = Table(filas, colWidths=[9 * cm, 5 * cm, 4 * cm])
        tabla.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), COLOR_ACENTO),
            ('ROWBACKGROUNDS', (0, 1), (-1, n - 3), [BLANCO, COLOR_GRIS]),
            ('BACKGROUND', (0, n - 3), (-1, n - 3), HexColor('#fce8e8')),
            ('BACKGROUND', (0, n - 2), (-1, n - 2), HexColor('#fff9e6')),
            ('BACKGROUND', (0, n - 1), (-1, n - 1), HexColor('#8B0000')),
            ('TEXTCOLOR',  (0, n - 1), (-1, n - 1), BLANCO),
            ('FONTNAME',   (0, n - 1), (-1, n - 1), 'Helvetica-Bold'),
            ('GRID',    (0, 0), (-1, -1), 0.5, COLOR_BORDE),
            ('VALIGN',  (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING',    (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING',   (0, 0), (-1, -1), 6),
        ]))

        return [cab, tabla]

    def _seccion_resumen(self, s: dict) -> list:
        """Genera el cuadro resumen con el sueldo lÃ­quido."""
        r = self.res

        filas = [
            [
                Paragraph("Total Haberes:", s['normal_bold']),
                Paragraph(_fmt_pesos(r['total_haberes']), s['monto']),
            ],
            [
                Paragraph("Total Descuentos:", s['normal_bold']),
                Paragraph(f"- {_fmt_pesos(r['total_descuentos'])}", s['monto']),
            ],
            [
                Paragraph("SUELDO LÃQUIDO A PAGAR:", ParagraphStyle(
                    'liq_label', fontSize=12, fontName='Helvetica-Bold',
                    textColor=BLANCO
                )),
                Paragraph(_fmt_pesos(r['sueldo_liquido']), ParagraphStyle(
                    'liq_monto', fontSize=13, fontName='Helvetica-Bold',
                    textColor=BLANCO, alignment=TA_RIGHT
                )),
            ],
        ]

        tabla = Table(filas, colWidths=[10 * cm, 8 * cm])
        tabla.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), HexColor('#dcedf5')),
            ('BACKGROUND', (0, 1), (-1, 1), HexColor('#fce8e8')),
            ('BACKGROUND', (0, 2), (-1, 2), COLOR_EXITO),
            ('TEXTCOLOR',  (0, 2), (-1, 2), BLANCO),
            ('GRID',    (0, 0), (-1, -1), 1, COLOR_BORDE),
            ('VALIGN',  (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING',    (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('LEFTPADDING',   (0, 0), (-1, -1), 10),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 10),
        ]))

        return [tabla]

    def _seccion_firmas(self, s: dict) -> list:
        """Genera la secciÃ³n de firmas."""
        filas = [[
            Paragraph("___________________________", s['normal']),
            Paragraph("___________________________", s['normal']),
        ], [
            Paragraph("Firma Empleador", s['label']),
            Paragraph("Firma Trabajador / RecibÃ­ Conforme", s['label']),
        ], [
            Paragraph(self.empresa['nombre'], s['label']),
            Paragraph(self.trabajador.get('nombre', ''), s['label']),
        ]]

        tabla = Table(filas, colWidths=[9 * cm, 9 * cm])
        tabla.setStyle(TableStyle([
            ('ALIGN',  (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING',    (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ]))

        return [Spacer(1, 0.5 * cm), tabla]


# ââââââââââââââââââââââââââââââââââââââââââââââââââ
#  FUNCIÃN DE CONVENIENCIA
# ââââââââââââââââââââââââââââââââââââââââââââââââââ

def generar_pdf_liquidacion(datos_trabajador: dict, datos_entrada: dict,
                             resultado: dict, empresa: dict = None,
                             periodo: str = None) -> str:
    """
    Genera un PDF de liquidaciÃ³n de sueldo.
    Retorna la ruta del archivo generado.
    """
    gen = GeneradorPDF(datos_trabajador, datos_entrada, resultado, empresa, periodo)
    return gen.generar()
