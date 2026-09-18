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
import io
import unicodedata
from pathlib import Path

from PIL import Image as PILImage

from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)
from reportlab.lib.utils import ImageReader

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
        "imagen_portada": IMG_DIR / "portada_2024.jpg",
        "entidad_editora": "Secretaría de Medio Ambiente y Desarrollo Territorial – SEMADET",
        "fecha_publicacion": "Publicación: Octubre 2025",
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
        # Tabla 3: contaminante criterio y la NOM de salud aplicable a cada uno.
        "normas_nom": [
            {"contaminante": "Ozono", "simbolo": "O₃", "nom": "NOM-020-SSA1-2021"},
            {"contaminante": "Monóxido de Carbono", "simbolo": "CO", "nom": "NOM-021-SSA1-2021"},
            {"contaminante": "Dióxido de azufre", "simbolo": "SO₂", "nom": "NOM-022-SSA1-2019"},
            {"contaminante": "Partículas (diámetro ≤ 10 μm)", "simbolo": "PM₁₀", "nom": "NOM-025-SSA1-2021"},
            {"contaminante": "Partículas (diámetro ≤ 2.5 μm)", "simbolo": "PM₂.₅", "nom": "NOM-025-SSA1-2021"},
        ],
        # Nota que explica por qué NO2 (medido por la red) no entra a la
        # evaluación anual de cumplimiento de NOM ni a la Tabla 3.
        "nota_no2": (
            "La evaluación en este informe considera <b>O₃, PM₁₀, PM₂.₅, SO₂ y CO</b>. NO₂ fue excluido "
            "de la evaluación anual debido a la modernización de la red en 2024, ya que se reemplazaron "
            "los analizadores y no cumple con la suficiencia anual requerida por la NOM para su análisis."
        ),
        # Tabla 4: categorías del Índice Aire y Salud (NOM-172-SEMARNAT-2023).
        # "reco_unica" son filas donde la recomendación es la misma para toda
        # la población; "reco_general"/"reco_sensibles" cuando difiere.
        "categorias_ias": [
            {"categoria": "Buena", "color": "#2E9E3B", "riesgo": "Bajo",
             "reco_unica": "Disfruta de las actividades al aire libre."},
            {"categoria": "Aceptable", "color": "#D9B32C", "riesgo": "Moderado",
             "reco_general": "Disfruta de las actividades al aire libre.",
             "reco_sensibles": "Reduce las actividades físicas vigorosas al aire libre."},
            {"categoria": "Mala", "color": "#E8813A", "riesgo": "Alto",
             "reco_general": "Reduce las actividades físicas vigorosas al aire libre.",
             "reco_sensibles": "Evita las actividades físicas al aire libre."},
            {"categoria": "Muy Mala", "color": "#D0453F", "riesgo": "Muy alto",
             "reco_general": "Evita las actividades físicas al aire libre.",
             "reco_sensibles": "Limita estar al aire libre."},
            {"categoria": "Extremadamente mala", "color": "#7A3FA0", "riesgo": "Extremadamente alto",
             "reco_unica": "Permanece en interiores y evita cualquier esfuerzo físico al aire libre."},
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
    # Permite usar <b>texto</b> dentro de cualquier Paragraph en fuente Montserrat.
    pdfmetrics.registerFontFamily(
        "Montserrat", normal="Montserrat", bold="Montserrat-Bold",
        italic="Montserrat", boldItalic="Montserrat-Bold",
    )


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
        "nota_destacada": ParagraphStyle(
            "nota_destacada", fontName="Montserrat", fontSize=9.7, leading=14.5,
            textColor=NAVY,
        ),
        "tabla_celda_chica": ParagraphStyle(
            "tabla_celda_chica", fontName="Montserrat", fontSize=8.6, leading=12,
            textColor=TEXT,
        ),
        "tabla_celda_chica_bold": ParagraphStyle(
            "tabla_celda_chica_bold", fontName="Montserrat-Bold", fontSize=8.8, leading=12,
            textColor=NAVY,
        ),
        "tabla_subheader": ParagraphStyle(
            "tabla_subheader", fontName="Montserrat-Bold", fontSize=8.3, leading=11,
            textColor=colors.white, alignment=TA_CENTER,
        ),
    }


# ---------------------------------------------------------------------------
# Portada (página 1, a sangre completa)
# ---------------------------------------------------------------------------
def _imagen_gradiente_navy(alto_px=600, alpha_max=0.78):
    """Genera en memoria un degradado vertical (transparente -> navy) para
    usarlo como velo detrás del texto de la portada. Se hace como imagen,
    en vez de franjas rectangulares apiladas, para que la transición sea
    perfectamente suave (las franjas producían líneas por redondeo)."""
    r, g, b = (0x17, 0x3D, 0x4C)
    gradiente = PILImage.new("RGBA", (2, alto_px))
    for y in range(alto_px):
        frac = y / (alto_px - 1)
        alpha = int(255 * alpha_max * frac)
        gradiente.putpixel((0, y), (r, g, b, alpha))
        gradiente.putpixel((1, y), (r, g, b, alpha))
    buffer = io.BytesIO()
    gradiente.save(buffer, format="PNG")
    buffer.seek(0)
    return buffer


def _dibujar_portada(anio, datos):
    ruta_imagen = Path(datos["imagen_portada"])

    def _dibujar(canvas, doc):
        canvas.saveState()
        ancho, alto = PAGE_SIZE

        if ruta_imagen.exists():
            img = ImageReader(str(ruta_imagen))
            img_w, img_h = img.getSize()
            escala = max(ancho / img_w, alto / img_h)
            dib_w, dib_h = img_w * escala, img_h * escala
            x = (ancho - dib_w) / 2
            y = (alto - dib_h) / 2
            canvas.drawImage(img, x, y, width=dib_w, height=dib_h)
        else:
            canvas.setFillColor(NAVY)
            canvas.rect(0, 0, ancho, alto, fill=1, stroke=0)

        # Velo degradado en la parte inferior para que el texto resalte
        # sobre la foto. Se usa una imagen (no franjas dibujadas a mano) para
        # que el degradado sea perfectamente suave, sin bandas de redondeo.
        alto_velo = alto * 0.46
        gradiente = ImageReader(_imagen_gradiente_navy())
        canvas.drawImage(gradiente, 0, 0, width=ancho, height=alto_velo, mask="auto")

        # Bloque de título
        canvas.setFillColor(AQUA)
        canvas.setFont("Montserrat-Bold", 13)
        canvas.drawString(MARGIN, 8.6 * cm, "I N F O R M E   A N U A L")

        canvas.setFillColor(colors.white)
        canvas.setFont("Montserrat-Bold", 40)
        canvas.drawString(MARGIN, 7.35 * cm, "Calidad del Aire")

        canvas.setFillColor(AQUA)
        canvas.setFont("Montserrat-SemiBold", 19)
        canvas.drawString(MARGIN, 6.35 * cm, f"— Jalisco {anio} —")

        # Línea y pie con la entidad y fecha de publicación
        canvas.setStrokeColor(AQUA)
        canvas.setLineWidth(1.1)
        canvas.line(MARGIN, 2.35 * cm, ancho - MARGIN, 2.35 * cm)

        canvas.setFillColor(colors.white)
        canvas.setFont("Montserrat-SemiBold", 9.5)
        canvas.drawString(MARGIN, 1.75 * cm, datos["entidad_editora"])
        canvas.setFillColor(colors.HexColor("#CFE8E4"))
        canvas.setFont("Montserrat", 9.5)
        canvas.drawRightString(ancho - MARGIN, 1.75 * cm, datos["fecha_publicacion"])

        canvas.restoreState()

    return _dibujar


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

    story.append(Paragraph(
        "Para garantizar información confiable y comparable a nivel nacional, el SIMAJ utiliza equipos, "
        "procedimientos y criterios de calidad conforme a las Normas Oficiales Mexicanas para cada "
        "contaminante: O₃ (NOM-036-SEMARNAT-1993), NOx (NOM-037-SEMARNAT-1993), CO (NOM-034-SEMARNAT-1993) "
        "y SO₂ (NOM-038-SEMARNAT-1993). La instalación, operación y validez de las redes también cumplen "
        "con la NOM-156-SEMARNAT-2012 y las especificaciones de los Manuales SINAICA.",
        estilos["cuerpo"],
    ))

    return story


def _seccion_evaluacion_nom(estilos, datos):
    story = [Paragraph("Evaluación de Normas Oficiales Mexicanas de calidad del aire", estilos["h2"])]

    story.append(Paragraph(
        "La calidad del aire se evalúa mediante dos herramientas complementarias. La primera es el "
        "cumplimiento de las Normas Oficiales Mexicanas (NOM), que permite determinar si las "
        "concentraciones de los contaminantes se mantienen dentro de los límites establecidos para "
        "proteger la salud. La segunda es el Índice Aire y Salud (IAS), establecido en la "
        "NOM-172-SEMARNAT-2023, que traduce esas concentraciones en categorías de riesgo fáciles de "
        "comprender para la población.",
        estilos["cuerpo"],
    ))

    story.append(Paragraph(
        "El cumplimiento normativo consiste en comparar las concentraciones registradas de cada "
        "contaminante con los valores establecidos en las NOM. Cuando estos límites son superados, se "
        "considera “fuera de norma”, por lo que aumenta el riesgo de efectos adversos en la "
        "salud, especialmente para niñas y niños, personas adultas mayores, mujeres embarazadas y "
        "personas con enfermedades respiratorias o cardiovasculares. La Tabla 3 presenta los "
        "contaminantes evaluados y la NOM aplicable a cada uno.",
        estilos["cuerpo"],
    ))

    story.append(KeepTogether(_tabla_normas(estilos, datos["normas_nom"])))
    story.append(Spacer(1, 10))
    story.append(_callout(datos["nota_no2"], estilos))
    story.append(Spacer(1, 4))

    story.append(Paragraph(
        "Por otra parte, el IAS clasifica la calidad del aire en cinco categorías: Buena, Aceptable, "
        "Mala, Muy mala y Extremadamente mala. Cada categoría está representada por un color y se "
        "acompaña de recomendaciones para reducir la exposición cuando los niveles de contaminación "
        "representan un riesgo para la salud. La Tabla 4 muestra las categorías del Índice Aire y Salud "
        "y los niveles de riesgo asociados.",
        estilos["cuerpo"],
    ))

    story.append(KeepTogether(_tabla_ias(estilos, datos["categorias_ias"])))

    return story


def _callout(texto_html, estilos):
    celda = Table(
        [[Paragraph(texto_html, estilos["nota_destacada"])]],
        colWidths=[CONTENT_WIDTH],
    )
    celda.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#EAF8F6")),
        ("BOX", (0, 0), (-1, -1), 1.1, AQUA),
        ("ROUNDEDCORNERS", [10, 10, 10, 10]),
        ("TOPPADDING", (0, 0), (-1, -1), 12),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 12),
        ("LEFTPADDING", (0, 0), (-1, -1), 14),
        ("RIGHTPADDING", (0, 0), (-1, -1), 14),
    ]))
    return celda


def _tabla_normas(estilos, filas):
    encabezados = ["Contaminante Criterio", "Símbolo", "Norma Oficial Mexicana"]
    data = [[Paragraph(h, estilos["tabla_header"]) for h in encabezados]]
    for f in filas:
        data.append([
            Paragraph(f["contaminante"], estilos["tabla_celda"]),
            Paragraph(f["simbolo"], estilos["tabla_celda_centrada"]),
            Paragraph(f["nom"], estilos["tabla_celda"]),
        ])

    anchos = [CONTENT_WIDTH * w for w in (0.46, 0.16, 0.38)]
    tabla = Table(data, colWidths=anchos, repeatRows=1)
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D7DEE1")),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
    ]))

    return [
        Paragraph(
            "Tabla 3. Relación de Contaminante Criterio con la Norma Oficial Mexicana aplicable.",
            estilos["tabla_caption"],
        ),
        tabla,
    ]


def _punto_color(hex_color):
    punto = Table([[""]], colWidths=[11], rowHeights=[11])
    punto.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(hex_color)),
        ("ROUNDEDCORNERS", [5.5, 5.5, 5.5, 5.5]),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    punto.hAlign = "CENTER"
    return punto


def _tabla_ias(estilos, filas):
    fila_encabezado_1 = [
        Paragraph("Calidad del aire", estilos["tabla_header"]),
        Paragraph("Color", estilos["tabla_header"]),
        Paragraph("Riesgo asociado", estilos["tabla_header"]),
        Paragraph("Recomendaciones", estilos["tabla_header"]),
        "",
    ]
    fila_encabezado_2 = [
        "", "", "",
        Paragraph("Población general", estilos["tabla_subheader"]),
        Paragraph("Grupos sensibles", estilos["tabla_subheader"]),
    ]

    data = [fila_encabezado_1, fila_encabezado_2]
    for f in filas:
        if "reco_unica" in f:
            reco_general = Paragraph(f["reco_unica"], estilos["tabla_celda_chica"])
            reco_sensibles = ""
        else:
            reco_general = Paragraph(f["reco_general"], estilos["tabla_celda_chica"])
            reco_sensibles = Paragraph(f["reco_sensibles"], estilos["tabla_celda_chica"])
        data.append([
            Paragraph(f["categoria"], estilos["tabla_celda_chica_bold"]),
            _punto_color(f["color"]),
            Paragraph(f["riesgo"], estilos["tabla_celda_chica"]),
            reco_general,
            reco_sensibles,
        ])

    anchos = [CONTENT_WIDTH * w for w in (0.19, 0.09, 0.20, 0.26, 0.26)]
    tabla = Table(data, colWidths=anchos, repeatRows=2)

    estilo_cmds = [
        ("BACKGROUND", (0, 0), (-1, 1), NAVY),
        ("SPAN", (0, 0), (0, 1)),
        ("SPAN", (1, 0), (1, 1)),
        ("SPAN", (2, 0), (2, 1)),
        ("SPAN", (3, 0), (4, 0)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 2), (1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D7DEE1")),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (0, -1), 5),
        ("RIGHTPADDING", (0, 0), (0, -1), 5),
        ("LEFTPADDING", (2, 0), (2, -1), 5),
        ("RIGHTPADDING", (2, 0), (2, -1), 5),
        ("ROWBACKGROUNDS", (0, 2), (-1, -1), [colors.white, LIGHT]),
    ]

    # Filas con una sola recomendación para toda la población: se combinan
    # las dos columnas de recomendaciones.
    for i, f in enumerate(filas):
        if "reco_unica" in f:
            fila = i + 2
            estilo_cmds.append(("SPAN", (3, fila), (4, fila)))

    tabla.setStyle(TableStyle(estilo_cmds))

    return [
        Paragraph(
            "Tabla 4. Categorías, riesgo asociado para la salud y recomendaciones del Índice Aire y Salud.",
            estilos["tabla_caption"],
        ),
        tabla,
    ]


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

    story = [PageBreak()]
    story += _seccion_introduccion(anio, estilos)
    story += _seccion_simaj(anio, estilos, datos)
    story += _seccion_evaluacion_nom(estilos, datos)

    dibujar_portada = _dibujar_portada(anio, datos)
    pintar_fondo = _fondo_pagina(anio)
    doc.build(story, onFirstPage=dibujar_portada, onLaterPages=pintar_fondo)
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
