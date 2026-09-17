"""
Generador del Informe Anual de Calidad del Aire - Jalisco.

Uso:
    python generar_informe.py --anio 2024
    python generar_informe.py --anio 2024 --salida "Informe 2024.pdf"

Este script arma el PDF sección por sección. Por ahora incluye:
    1. Introducción
    2. Sistema de Monitoreo Atmosférico de Jalisco (SIMAJ)

Las secciones siguientes (cumplimiento de NOM, Índice Aire y Salud,
comportamiento de contaminantes, etc.) se agregan como nuevas funciones
"_seccion_xxx(anio)" que devuelven una lista de flowables, y se van
sumando a la lista `story` en generar_informe().

Los datos que cambian de un año a otro (población, imagen de la red de
monitoreo, cifras, tablas...) se definen en DATOS_POR_ANIO más abajo, o
se pueden mover a un archivo/base de datos aparte cuando esté lista.
"""
import argparse
import unicodedata
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Image, KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

BASE_DIR = Path(__file__).resolve().parent
FONTS_DIR = BASE_DIR / "assets" / "fonts"
IMG_DIR = BASE_DIR / "assets" / "img"

# ---------------------------------------------------------------------------
# Paleta e identidad visual (misma paleta usada en el dashboard)
# ---------------------------------------------------------------------------
AQUA = colors.HexColor("#00C4B4")
NAVY = colors.HexColor("#173D4C")
LIGHT = colors.HexColor("#E3E9EC")
TEXT = colors.HexColor("#324D59")
MUTED = colors.HexColor("#6C8894")

PAGE_SIZE = letter
MARGIN = 2.2 * cm
CONTENT_WIDTH = PAGE_SIZE[0] - 2 * MARGIN

# ---------------------------------------------------------------------------
# Datos que varían por año de informe
# ---------------------------------------------------------------------------
DATOS_POR_ANIO = {
    2024: {
        "poblacion_amg": "5,268,642",
        "porcentaje_poblacion_amg": "63",
        "fuente_poblacion": "IIEG, 2022; INEGI, 2021",
        "imagen_red_monitoreo": IMG_DIR / "red_monitoreo_2024.png",
        "caption_imagen_red": (
            "Imagen 1. Ubicación de las estaciones de monitoreo que conforman "
            "el Sistema de Monitoreo Atmosférico de Jalisco (SIMAJ) en el "
            "Área Metropolitana de Guadalajara."
        ),
        # Tabla 1 y Tabla 2: estaciones, su área de influencia, año de inicio
        # de operación y contaminantes que mide cada una. "nueva" marca las
        # tres estaciones incorporadas en septiembre de 2024. El histórico de
        # contaminantes de las 10 estaciones originales viene de la Tabla 9.1
        # del Informe Nacional de Calidad del Aire (INECC); Oblatos y Atemajac
        # no cuentan con equipo para PM2.5.
        "estaciones": [
            {"municipio": "San Pedro Tlaquepaque", "area_influencia": "San Pedro Tlaquepaque y Guadalajara", "estacion": "Tlaquepaque", "simbolo": "TLA", "nueva": False, "anio_inicio": 1993, "contaminantes": {"PM10": True, "PM25": True, "O3": True, "SO2": True, "NO2": True, "CO": True}},
            {"municipio": "El Salto", "area_influencia": "El Salto, San Pedro Tlaquepaque y Tlajomulco de Zúñiga", "estacion": "Pintas", "simbolo": "PIN", "nueva": False, "anio_inicio": 2011, "contaminantes": {"PM10": True, "PM25": True, "O3": True, "SO2": True, "NO2": True, "CO": True}},
            {"municipio": "Guadalajara", "area_influencia": "Guadalajara", "estacion": "Centro", "simbolo": "CEN", "nueva": False, "anio_inicio": 1993, "contaminantes": {"PM10": True, "PM25": True, "O3": True, "SO2": True, "NO2": True, "CO": True}},
            {"municipio": "Guadalajara", "area_influencia": "Zapopan y Guadalajara", "estacion": "Country", "simbolo": "COU", "nueva": True, "anio_inicio": 2024, "contaminantes": {"PM10": True, "PM25": True, "O3": True, "SO2": True, "NO2": True, "CO": True}},
            {"municipio": "Guadalajara", "area_influencia": "Guadalajara y San Pedro Tlaquepaque", "estacion": "Miravalle", "simbolo": "MIR", "nueva": False, "anio_inicio": 1993, "contaminantes": {"PM10": True, "PM25": True, "O3": True, "SO2": True, "NO2": True, "CO": True}},
            {"municipio": "Guadalajara", "area_influencia": "Guadalajara", "estacion": "Oblatos", "simbolo": "OBL", "nueva": False, "anio_inicio": 1993, "contaminantes": {"PM10": True, "PM25": False, "O3": True, "SO2": True, "NO2": True, "CO": True}},
            {"municipio": "Tlajomulco de Zúñiga", "area_influencia": "Tlajomulco de Zúñiga y San Pedro Tlaquepaque", "estacion": "Santa Anita", "simbolo": "SAN", "nueva": True, "anio_inicio": 2024, "contaminantes": {"PM10": True, "PM25": True, "O3": True, "SO2": True, "NO2": True, "CO": True}},
            {"municipio": "Tlajomulco de Zúñiga", "area_influencia": "Tlajomulco de Zúñiga, San Pedro Tlaquepaque y El Salto", "estacion": "Santa Fe", "simbolo": "SFE", "nueva": False, "anio_inicio": 2013, "contaminantes": {"PM10": True, "PM25": True, "O3": True, "SO2": True, "NO2": True, "CO": True}},
            {"municipio": "Tonalá", "area_influencia": "Tonalá y Guadalajara", "estacion": "Loma Dorada", "simbolo": "LDO", "nueva": False, "anio_inicio": 1993, "contaminantes": {"PM10": True, "PM25": True, "O3": True, "SO2": True, "NO2": True, "CO": True}},
            {"municipio": "Zapopan", "area_influencia": "Zapopan, Guadalajara y San Pedro Tlaquepaque", "estacion": "Águilas", "simbolo": "AGU", "nueva": False, "anio_inicio": 1993, "contaminantes": {"PM10": True, "PM25": True, "O3": True, "SO2": True, "NO2": True, "CO": True}},
            {"municipio": "Zapopan", "area_influencia": "Zapopan y Guadalajara", "estacion": "Atemajac", "simbolo": "ATM", "nueva": False, "anio_inicio": 1993, "contaminantes": {"PM10": True, "PM25": False, "O3": True, "SO2": True, "NO2": True, "CO": True}},
            {"municipio": "Zapopan", "area_influencia": "Zapopan", "estacion": "Santa Margarita", "simbolo": "SMT", "nueva": True, "anio_inicio": 2024, "contaminantes": {"PM10": True, "PM25": True, "O3": True, "SO2": True, "NO2": True, "CO": True}},
            {"municipio": "Zapopan", "area_influencia": "Zapopan y Guadalajara", "estacion": "Vallarta", "simbolo": "VAL", "nueva": False, "anio_inicio": 1993, "contaminantes": {"PM10": True, "PM25": True, "O3": True, "SO2": True, "NO2": True, "CO": True}},
        ],
    },
}


def _datos_del_anio(anio):
    if anio not in DATOS_POR_ANIO:
        anios_disp = ", ".join(str(a) for a in sorted(DATOS_POR_ANIO))
        raise ValueError(
            f"No hay datos configurados para el año {anio}. "
            f"Años disponibles: {anios_disp}."
        )
    return DATOS_POR_ANIO[anio]


# ---------------------------------------------------------------------------
# Tipografía
# ---------------------------------------------------------------------------
def _registrar_fuentes():
    pdfmetrics.registerFont(TTFont("Montserrat", FONTS_DIR / "Montserrat-Regular.ttf"))
    pdfmetrics.registerFont(TTFont("Montserrat-SemiBold", FONTS_DIR / "Montserrat-SemiBold.ttf"))
    pdfmetrics.registerFont(TTFont("Montserrat-Bold", FONTS_DIR / "Montserrat-Bold.ttf"))


def _estilos():
    return {
        "h2": ParagraphStyle(
            "h2", fontName="Montserrat-Bold", fontSize=14, leading=18,
            textColor=NAVY, spaceBefore=22, spaceAfter=10,
        ),
        "cuerpo": ParagraphStyle(
            "cuerpo", fontName="Montserrat", fontSize=10.3, leading=16.5,
            textColor=TEXT, alignment=TA_JUSTIFY, spaceAfter=10,
        ),
        "caption": ParagraphStyle(
            "caption", fontName="Montserrat-SemiBold", fontSize=8.6, leading=11.5,
            textColor=MUTED, alignment=TA_CENTER, spaceBefore=8, spaceAfter=4,
        ),
        "tabla_caption": ParagraphStyle(
            "tabla_caption", fontName="Montserrat-SemiBold", fontSize=10, leading=13,
            textColor=NAVY, spaceBefore=14, spaceAfter=8,
        ),
        "nota": ParagraphStyle(
            "nota", fontName="Montserrat", fontSize=8.3, leading=11,
            textColor=MUTED, spaceBefore=5,
        ),
        "tabla_header": ParagraphStyle(
            "tabla_header", fontName="Montserrat-Bold", fontSize=9.5, leading=12,
            textColor=colors.white, alignment=TA_CENTER,
        ),
        "tabla_celda": ParagraphStyle(
            "tabla_celda", fontName="Montserrat", fontSize=9, leading=12,
            textColor=TEXT,
        ),
        "tabla_celda_centrada": ParagraphStyle(
            "tabla_celda_centrada", fontName="Montserrat", fontSize=9, leading=12,
            textColor=TEXT, alignment=TA_CENTER,
        ),
        "tabla_celda_bold": ParagraphStyle(
            "tabla_celda_bold", fontName="Montserrat-Bold", fontSize=9, leading=12,
            textColor=TEXT,
        ),
        "tabla_celda_centrada_bold": ParagraphStyle(
            "tabla_celda_centrada_bold", fontName="Montserrat-Bold", fontSize=9, leading=12,
            textColor=TEXT, alignment=TA_CENTER,
        ),
        "estacion_nombre": ParagraphStyle(
            "estacion_nombre", fontName="Montserrat-Bold", fontSize=9.3, leading=12,
            textColor=NAVY,
        ),
        "estacion_clave": ParagraphStyle(
            "estacion_clave", fontName="Montserrat-SemiBold", fontSize=7.3, leading=10,
            textColor=MUTED,
        ),
        "badge_medido": ParagraphStyle(
            "badge_medido", fontName="Montserrat-Bold", fontSize=6.6, leading=8,
            textColor=colors.white, alignment=TA_CENTER,
        ),
        "badge_no_medido": ParagraphStyle(
            "badge_no_medido", fontName="Montserrat-SemiBold", fontSize=6.6, leading=8,
            textColor=colors.HexColor("#9AAAB1"), alignment=TA_CENTER,
        ),
    }


# ---------------------------------------------------------------------------
# Encabezado / pie de página (se dibujan en cada página)
# ---------------------------------------------------------------------------
def _fondo_pagina(anio):
    def _dibujar(canvas, doc):
        canvas.saveState()
        ancho, alto = PAGE_SIZE

        # Franja superior de marca
        banda_alto = 3.1 * cm
        canvas.setFillColor(NAVY)
        canvas.rect(0, alto - banda_alto, ancho, banda_alto, fill=1, stroke=0)
        canvas.setFillColor(AQUA)
        canvas.rect(0, alto - banda_alto - 0.12 * cm, ancho, 0.12 * cm, fill=1, stroke=0)

        canvas.setFillColor(AQUA)
        canvas.setFont("Montserrat-Bold", 9)
        canvas.drawString(MARGIN, alto - 1.15 * cm, "SIMAJ · JALISCO")
        canvas.setFillColor(colors.white)
        canvas.setFont("Montserrat-Bold", 16.5)
        canvas.drawString(MARGIN, alto - 1.95 * cm, f"Informe Anual de Calidad del Aire {anio}")
        canvas.setFillColor(colors.HexColor("#CFE8E4"))
        canvas.setFont("Montserrat", 9.5)
        canvas.drawString(MARGIN, alto - 2.55 * cm, "Área Metropolitana de Guadalajara")

        # Pie de página
        canvas.setFillColor(MUTED)
        canvas.setFont("Montserrat", 8)
        canvas.drawString(MARGIN, 1.1 * cm, "Sistema de Monitoreo Atmosférico de Jalisco (SIMAJ)")
        canvas.drawRightString(ancho - MARGIN, 1.1 * cm, f"Página {doc.page}")
        canvas.setStrokeColor(colors.HexColor("#D7DEE1"))
        canvas.setLineWidth(0.6)
        canvas.line(MARGIN, 1.4 * cm, ancho - MARGIN, 1.4 * cm)

        canvas.restoreState()

    return _dibujar


# ---------------------------------------------------------------------------
# Secciones del informe
# ---------------------------------------------------------------------------
def _seccion_introduccion(anio, estilos):
    story = [Paragraph("Introducción", estilos["h2"])]
    parrafos = [
        f"El Informe Anual de Calidad del Aire – Jalisco {anio} presenta un panorama de la calidad del "
        f"aire registrada durante {anio} en el Área Metropolitana de Guadalajara (AMG), con base en los "
        "datos generados por el Sistema de Monitoreo Atmosférico de Jalisco (SIMAJ) y validados por el "
        "Instituto Nacional de Ecología y Cambio Climático (INECC).",

        "En este informe se analizan los contaminantes criterio ozono (O₃), partículas PM₂.₅, "
        "partículas PM₁₀, dióxido de azufre (SO₂) y monóxido de carbono (CO). Además, se presenta una "
        "evaluación de la calidad del aire basada en el cumplimiento de las Normas Oficiales Mexicanas y "
        "en el Índice Aire y Salud, así como un análisis del comportamiento de los principales "
        f"contaminantes registrados en {anio}.",

        "El propósito de este informe es presentar de manera clara y accesible la información sobre la "
        "calidad del aire, identificar los principales contaminantes que afectaron al AMG durante el año "
        "y proporcionar elementos que apoyen la toma de decisiones para proteger la salud de la "
        "población.",
    ]
    for p in parrafos:
        story.append(Paragraph(p, estilos["cuerpo"]))
    return story


def _seccion_simaj(anio, estilos, datos):
    story = [Paragraph("Sistema de Monitoreo Atmosférico de Jalisco", estilos["h2"])]

    texto = (
        "El Sistema de Monitoreo Atmosférico de Jalisco (SIMAJ) es la red estatal encargada de medir de "
        "manera continua la calidad del aire. Su operación se concentra en el AMG, donde habita "
        f"aproximadamente el {datos['porcentaje_poblacion_amg']} % de la población del estado "
        f"({datos['poblacion_amg']} habitantes) ({datos['fuente_poblacion']}), por lo que representa la "
        "zona con mayor exposición potencial a la contaminación atmosférica. La Imagen 1 muestra la "
        "ubicación de las estaciones de monitoreo que conforman la red."
    )
    story.append(Paragraph(texto, estilos["cuerpo"]))

    ruta_imagen = Path(datos["imagen_red_monitoreo"])
    if ruta_imagen.exists():
        img_reader = Image(str(ruta_imagen))
        ancho_original, alto_original = img_reader.imageWidth, img_reader.imageHeight
        ancho_final = CONTENT_WIDTH * 0.92
        alto_final = ancho_final * (alto_original / ancho_original)
        img_reader.drawWidth = ancho_final
        img_reader.drawHeight = alto_final
        img_reader.hAlign = "CENTER"

        story.append(Spacer(1, 6))
        story.append(KeepTogether([
            img_reader,
            Paragraph(datos["caption_imagen_red"], estilos["caption"]),
        ]))
    else:
        story.append(Paragraph(
            f"[Falta la imagen de la red de monitoreo en: {ruta_imagen}]",
            estilos["caption"],
        ))

    story.append(Paragraph(
        "Las estaciones del SIMAJ registran de manera continua los contaminantes criterio (PM₂.₅, "
        "PM₁₀, O₃, NO₂, SO₂ y CO), además de variables meteorológicas que ayudan a interpretar el "
        "comportamiento de la contaminación atmosférica. La información se actualiza cada hora y está "
        "disponible para consulta pública en el portal oficial www.aire.jalisco.gob.mx y en el Sistema "
        "Nacional de Información de la Calidad del Aire (SINAICA).",
        estilos["cuerpo"],
    ))

    story.append(Paragraph(
        f"Durante el primer semestre de {anio}, la red operó con diez estaciones de monitoreo. Como "
        "parte del programa de modernización del SIMAJ, en septiembre se incorporaron tres nuevas "
        "estaciones: Country (Guadalajara), Santa Margarita (Zapopan) y Santa Anita (Tlajomulco de "
        "Zúñiga), alcanzando un total de trece estaciones en operación al cierre del año. La Tabla 1 "
        "presenta la ubicación y el área de influencia de cada estación, mientras que la Tabla 2 "
        "muestra los contaminantes monitoreados en cada una de ellas.",
        estilos["cuerpo"],
    ))

    story += _tabla_estaciones(estilos, datos["estaciones"])
    story.append(Spacer(1, 10))
    story.append(KeepTogether(_tabla_contaminantes(estilos, datos["estaciones"])))

    return story


def _spans_de_grupos(valores):
    """A partir de una lista de valores, agrupa los índices consecutivos
    iguales. Devuelve una lista de tuplas (indice_inicio, indice_fin)."""
    grupos = []
    inicio = 0
    for i in range(1, len(valores) + 1):
        if i == len(valores) or valores[i] != valores[inicio]:
            grupos.append((inicio, i - 1))
            inicio = i
    return grupos


def _tabla_estaciones(estilos, filas):
    encabezados = ["Municipio", "Área de influencia", "Estación de monitoreo", "Símbolo"]
    data = [[Paragraph(h, estilos["tabla_header"]) for h in encabezados]]
    for f in filas:
        nombre = f["estacion"] + ("*" if f["nueva"] else "")
        estilo_nombre = estilos["tabla_celda_bold"] if f["nueva"] else estilos["tabla_celda"]
        estilo_simbolo = estilos["tabla_celda_centrada_bold"] if f["nueva"] else estilos["tabla_celda_centrada"]
        data.append([
            Paragraph(f["municipio"], estilos["tabla_celda"]),
            Paragraph(f["area_influencia"], estilos["tabla_celda"]),
            Paragraph(nombre, estilo_nombre),
            Paragraph(f["simbolo"], estilo_simbolo),
        ])

    anchos = [CONTENT_WIDTH * w for w in (0.19, 0.40, 0.27, 0.14)]
    tabla = Table(data, colWidths=anchos, repeatRows=1)

    estilo_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D7DEE1")),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]

    municipios = [f["municipio"] for f in filas]
    grupos = _spans_de_grupos(municipios)

    # Combina la celda de Municipio para cada grupo de filas consecutivas.
    for (i0, i1) in grupos:
        if i1 > i0:
            estilo_cmds.append(("SPAN", (0, i0 + 1), (0, i1 + 1)))

    # Bandas alternas por grupo de municipio.
    for idx, (i0, i1) in enumerate(grupos):
        if idx % 2 == 1:
            estilo_cmds.append(("BACKGROUND", (0, i0 + 1), (-1, i1 + 1), LIGHT))

    tabla.setStyle(TableStyle(estilo_cmds))

    return [
        Paragraph(
            "Tabla 1. Estaciones de monitoreo atmosférico del SIMAJ y su área de influencia.",
            estilos["tabla_caption"],
        ),
        tabla,
        Paragraph("* Estación nueva.", estilos["nota"]),
    ]


CONTAMINANTES_ORDEN = [
    ("PM10", "PM₁₀"),
    ("PM25", "PM₂.₅"),
    ("O3", "O₃"),
    ("SO2", "SO₂"),
    ("NO2", "NO₂"),
    ("CO", "CO"),
]


def _clave_alfabetica(texto):
    """Quita acentos para poder ordenar 'Águilas' junto con el resto del
    alfabeto, sin alterar el texto que se muestra en la tabla."""
    normalizado = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in normalizado if not unicodedata.combining(c)).upper()


def _celda_badges(fila, estilos):
    celdas = []
    for clave, etiqueta in CONTAMINANTES_ORDEN:
        medido = fila["contaminantes"][clave]
        estilo = estilos["badge_medido"] if medido else estilos["badge_no_medido"]
        celdas.append(Paragraph(etiqueta, estilo))

    ancho_badge = (CONTENT_WIDTH * 0.46) / len(CONTAMINANTES_ORDEN)
    badges = Table([celdas], colWidths=[ancho_badge] * len(CONTAMINANTES_ORDEN))
    badges.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("BOX", (0, 0), (-1, -1), 2, colors.white),
        ("INNERGRID", (0, 0), (-1, -1), 2, colors.white),
    ] + [
        ("BACKGROUND", (i, 0), (i, 0), AQUA if fila["contaminantes"][clave] else colors.HexColor("#E7ECEE"))
        for i, (clave, _) in enumerate(CONTAMINANTES_ORDEN)
    ]))
    return badges


def _tabla_contaminantes(estilos, filas):
    filas_ordenadas = sorted(filas, key=lambda f: _clave_alfabetica(f["estacion"]))

    encabezados = ["Estación", "Municipio", "Inicio de operación", "Contaminantes medidos"]
    data = [[Paragraph(h, estilos["tabla_header"]) for h in encabezados]]
    for f in filas_ordenadas:
        estacion_cell = Paragraph(
            f"{f['estacion']}{'*' if f['nueva'] else ''}<br/>"
            f"<font color='#6C8894' size=7.3>{f['simbolo']}</font>",
            estilos["estacion_nombre"],
        )
        data.append([
            estacion_cell,
            Paragraph(f["municipio"], estilos["tabla_celda"]),
            Paragraph(str(f["anio_inicio"]), estilos["tabla_celda_centrada"]),
            _celda_badges(f, estilos),
        ])

    anchos = [CONTENT_WIDTH * w for w in (0.19, 0.20, 0.15, 0.46)]
    tabla = Table(data, colWidths=anchos, repeatRows=1)
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (2, 0), (2, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D7DEE1")),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        # La columna de contaminantes lleva su propia tabla de recuadros ya
        # ajustada al ancho exacto de la columna: sin relleno extra aquí para
        # que no se salga del margen derecho.
        ("LEFTPADDING", (3, 0), (3, -1), 0),
        ("RIGHTPADDING", (3, 0), (3, -1), 0),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
    ]))

    return [
        Paragraph(
            "Tabla 2. Contaminantes monitoreados en cada estación del SIMAJ.",
            estilos["tabla_caption"],
        ),
        tabla,
        Paragraph(
            "* Estación nueva. Los recuadros en color indican los contaminantes que mide cada "
            "estación; en gris se señalan los contaminantes que la estación no mide.",
            estilos["nota"],
        ),
    ]


# ---------------------------------------------------------------------------
# Ensamblado del documento
# ---------------------------------------------------------------------------
def generar_informe(anio, salida=None):
    datos = _datos_del_anio(anio)
    _registrar_fuentes()
    estilos = _estilos()

    if salida is None:
        salida = BASE_DIR / f"Informe Anual de Calidad del Aire - Jalisco {anio}.pdf"
    salida = Path(salida)

    doc = SimpleDocTemplate(
        str(salida),
        pagesize=PAGE_SIZE,
        topMargin=3.6 * cm,
        bottomMargin=1.9 * cm,
        leftMargin=MARGIN,
        rightMargin=MARGIN,
        title=f"Informe Anual de Calidad del Aire - Jalisco {anio}",
        author="SIMAJ",
    )

    story = []
    story += _seccion_introduccion(anio, estilos)
    story += _seccion_simaj(anio, estilos, datos)

    pintar_fondo = _fondo_pagina(anio)
    doc.build(story, onFirstPage=pintar_fondo, onLaterPages=pintar_fondo)
    return salida


def _main():
    parser = argparse.ArgumentParser(description="Genera el Informe Anual de Calidad del Aire de Jalisco.")
    parser.add_argument("--anio", type=int, required=True, help="Año del informe, p. ej. 2024")
    parser.add_argument("--salida", type=str, default=None, help="Ruta del PDF de salida")
    args = parser.parse_args()

    ruta = generar_informe(args.anio, args.salida)
    print(f"PDF generado en: {ruta}")


if __name__ == "__main__":
    _main()
