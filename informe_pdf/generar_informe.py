"""
Generador del Informe Anual de Calidad del Aire - Jalisco.

Uso (un solo comando, con el año que se quiera):
    python generar_informe.py --anio 2024

La primera vez que se pide un año, este script descarga automáticamente la
base de datos horaria de ese año (BD_{anio} en Google Sheets) y calcula sus
cifras usando calculo_datos.py -- un port fiel de las funciones de cálculo
del notebook de diagnóstico (IAS, NOM-172, NowCast, etc.). El resultado se
guarda en datos/resumen_historico.json, así que las siguientes corridas para
ese mismo año son instantáneas (no vuelve a descargar/calcular a menos que
se use --recalcular).

    python generar_informe.py --anio 2024 --recalcular      # fuerza recálculo
    python generar_informe.py --anio 2024 --sin-datos-nuevos  # solo con lo ya guardado
    python generar_informe.py --anio 2024 --salida "Informe 2024.pdf"

Este script arma el PDF sección por sección. Por ahora incluye:
    1. Introducción
    2. Sistema de Monitoreo Atmosférico de Jalisco (SIMAJ)
    3. Evaluación de Normas Oficiales Mexicanas de calidad del aire
    4. Panorama general de la calidad del aire en el AMG

Las secciones siguientes se agregan como nuevas funciones
"_seccion_xxx(anio)" que devuelven una lista de flowables, y se van
sumando a la lista `story` en generar_informe().

Los datos editoriales/fijos (textos, tablas normativas, imágenes) se
definen en DATOS_POR_ANIO más abajo. Los datos calculados a partir de la
base de datos de cada año (por ahora, los días Buena/Aceptable de la
Figura 2) viven en datos/resumen_historico.json y sobreescriben a
DATOS_POR_ANIO cuando están disponibles -- ver _datos_del_anio().
"""
import argparse
import io
import json
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
from reportlab.graphics.shapes import Drawing, Circle, Line, Rect, String

BASE_DIR = Path(__file__).resolve().parent
FONTS_DIR = BASE_DIR / "assets" / "fonts"
IMG_DIR = BASE_DIR / "assets" / "img"
DATOS_DIR = BASE_DIR / "datos"
RUTA_RESUMEN_HISTORICO = DATOS_DIR / "resumen_historico.json"

# Cuántos años hacia atrás se muestran en la serie histórica (Figura 2, etc.)
ANIOS_HISTORICO = 6

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

# Mismo orden de estaciones que EST_ORDER_BASE en calculo_datos.py / tu notebook.
ORDEN_ESTACIONES_HORAS = ["AGU", "ATM", "CEN", "COU", "LDO", "MIR", "OBL", "PIN", "SAN", "SFE", "SMT", "TLA", "VAL"]

# Colores de las categorías IAS (mismos que en categorias_ias, para Tabla 4)
# más gris para "D.I.". Las claves deben coincidir exactamente con las que
# calculo_datos.py escribe en horas_categoria_calidad (CAT_ORDER usa "Muy
# mala" con minúscula, a diferencia de "Muy Mala" en categorias_ias).
CAT_ORDEN_HORAS = ["Buena", "Aceptable", "Mala", "Muy mala", "Extremadamente mala", "D.I."]
COLOR_CATEGORIA_IAS = {
    "Buena": colors.HexColor("#2E9E3B"),
    "Aceptable": colors.HexColor("#D9B32C"),
    "Mala": colors.HexColor("#E8813A"),
    "Muy mala": colors.HexColor("#D0453F"),
    "Extremadamente mala": colors.HexColor("#7A3FA0"),
    "D.I.": colors.HexColor("#B7C2C7"),
}

# Colores y mapeo de estatus para la Tabla 5 (cumplimiento de NOM). "cumple"
# y "no_cumple" reutilizan los verdes/rojos del IAS; DI/FO/sin_equipo usan
# grises neutros para no competir visualmente con los veredictos.
COLOR_NOM_CUMPLE = COLOR_CATEGORIA_IAS["Buena"]
COLOR_NOM_NO_CUMPLE = COLOR_CATEGORIA_IAS["Muy mala"]
COLOR_NOM_DI = COLOR_CATEGORIA_IAS["D.I."]
COLOR_NOM_FO = colors.HexColor("#E7ECEE")
COLOR_NOM_SIN_EQUIPO = colors.HexColor("#F5F7F8")

# calculo_datos.py devuelve "sin_datos" cuando una estación no tiene ningún
# registro válido en el año; aquí se resuelve si es porque el equipo estuvo
# fuera de operación (FO) o porque la estación no mide ese contaminante (¤),
# usando la capacidad ya declarada en datos["estaciones"][i]["contaminantes"].
CAPACIDAD_POR_CONTAMINANTE = {
    "PM10": "PM10", "PM2.5": "PM25", "O3": "O3", "CO": "CO", "NO2": "NO2", "SO2": "SO2",
}

# Etiquetas cortas para el encabezado de municipio en la Tabla 5: cuando el
# grupo cubre una sola estación (~30pt de ancho), "San Pedro Tlaquepaque"
# no cabe ni partiéndolo por sílabas sueltas; se usa un guion manual en el
# punto de corte para evitar que reportlab lo parta a media palabra.
MUNICIPIO_CORTO = {
    "San Pedro Tlaquepaque": "Tlaque-<br/>paque",
}

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
        # Días con IAS global "Buena" o "Aceptable" en el AMG, año del informe y los
        # 5 años anteriores, en orden reciente -> antiguo (así se muestran en la Figura 2).
        "serie_dias_buena_aceptable": [
            (2024, 80), (2023, 57), (2022, 88), (2021, 120), (2020, 101), (2019, 46),
        ],
        # % de horas del año en cada categoría IAS GLOBAL (dominante entre
        # los 6 contaminantes criterio), por estación (Figura 3). Calculado
        # con calculo_datos.py sobre BD_2024; ver datos/resumen_historico.json
        # para el valor con el que generar_informe.py arranca de verdad.
        "horas_categoria_calidad": {
            "AGU": {"Buena": 68.49, "Aceptable": 18.06, "Mala": 12.34, "Muy mala": 0.08, "D.I.": 1.04},
            "ATM": {"Buena": 27.55, "Aceptable": 5.56, "Mala": 3.19, "Muy mala": 0.01, "D.I.": 63.70},
            "CEN": {"Buena": 68.52, "Aceptable": 18.27, "Mala": 12.24, "Muy mala": 0.31, "Extremadamente mala": 0.11, "D.I.": 0.55},
            "COU": {"Buena": 20.88, "Aceptable": 5.26, "Mala": 6.08, "Muy mala": 0.10, "D.I.": 67.68},
            "LDO": {"Buena": 45.33, "Aceptable": 26.14, "Mala": 22.71, "Muy mala": 0.13, "Extremadamente mala": 0.01, "D.I.": 5.68},
            "MIR": {"Buena": 54.06, "Aceptable": 16.75, "Mala": 23.37, "Muy mala": 2.23, "Extremadamente mala": 0.24, "D.I.": 3.35},
            "OBL": {"Buena": 36.36, "Aceptable": 4.38, "Mala": 6.07, "Muy mala": 0.07, "Extremadamente mala": 0.03, "D.I.": 53.09},
            "PIN": {"Buena": 39.69, "Aceptable": 11.96, "Mala": 39.66, "Muy mala": 6.32, "Extremadamente mala": 1.84, "D.I.": 0.52},
            "SAN": {"Buena": 17.24, "Aceptable": 3.86, "Mala": 6.85, "Muy mala": 0.08, "D.I.": 71.97},
            "SFE": {"Buena": 25.69, "Aceptable": 3.87, "Mala": 14.77, "Muy mala": 3.70, "Extremadamente mala": 1.24, "D.I.": 50.73},
            "SMT": {"Buena": 25.72, "Aceptable": 3.57, "Mala": 1.34, "D.I.": 69.36},
            "TLA": {"Buena": 56.34, "Aceptable": 21.43, "Mala": 18.43, "Muy mala": 0.19, "D.I.": 3.61},
            "VAL": {"Buena": 51.50, "Aceptable": 25.67, "Mala": 15.48, "Muy mala": 0.08, "D.I.": 7.26},
        },
    },
}


def _cargar_resumen_historico():
    """Lee datos/resumen_historico.json (si existe). Ese archivo lo genera el
    notebook de Colab -- una entrada por año, con las cifras calculadas a
    partir de la base de datos de ese año (ver exportar_resumen_informe() en
    el notebook). Si no existe, el informe usa los valores de ejemplo fijados
    en DATOS_POR_ANIO más abajo."""
    if not RUTA_RESUMEN_HISTORICO.exists():
        return {}
    return json.loads(RUTA_RESUMEN_HISTORICO.read_text(encoding="utf-8"))


def _datos_del_anio(anio):
    if anio not in DATOS_POR_ANIO:
        anios_disp = ", ".join(str(a) for a in sorted(DATOS_POR_ANIO))
        raise ValueError(
            f"No hay datos configurados para el año {anio}. "
            f"Años disponibles: {anios_disp}."
        )
    datos = dict(DATOS_POR_ANIO[anio])

    historico = _cargar_resumen_historico()
    if str(anio) not in historico:
        return datos

    # El histórico manda: cualquier cifra que el notebook haya calculado para
    # este año reemplaza al valor de ejemplo. Metadatos editoriales (textos,
    # imágenes, tablas normativas fijas) no viven en el histórico y se quedan
    # como están en DATOS_POR_ANIO.
    entrada_anio = historico[str(anio)]

    if "dias_buena_aceptable" in entrada_anio:
        serie = []
        for i in range(ANIOS_HISTORICO):
            anio_i = anio - i
            valor = historico.get(str(anio_i), {}).get("dias_buena_aceptable")
            if valor is not None:
                serie.append((anio_i, valor))
        if serie:
            datos["serie_dias_buena_aceptable"] = serie

    # Cualquier otra cifra del histórico que ya tenga el mismo nombre que un
    # campo de DATOS_POR_ANIO simplemente lo sobreescribe (p. ej. población).
    for clave, valor in entrada_anio.items():
        if clave != "dias_buena_aceptable":
            datos[clave] = valor

    return datos


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
        "nom_etiqueta": ParagraphStyle(
            "nom_etiqueta", fontName="Montserrat-Bold", fontSize=7.6, leading=9.8,
            textColor=NAVY,
        ),
        "nom_celda_clara": ParagraphStyle(
            "nom_celda_clara", fontName="Montserrat-Bold", fontSize=7.6, leading=9.5,
            textColor=colors.white, alignment=TA_CENTER,
        ),
        "nom_celda_oscura": ParagraphStyle(
            "nom_celda_oscura", fontName="Montserrat-SemiBold", fontSize=7.3, leading=9.5,
            textColor=MUTED, alignment=TA_CENTER,
        ),
        "tabla_subheader_chico": ParagraphStyle(
            "tabla_subheader_chico", fontName="Montserrat-Bold", fontSize=6.8, leading=8.4,
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


def _seccion_panorama_general(anio, estilos, datos):
    story = [Paragraph(f"Panorama general de la calidad del aire en el AMG durante {anio}", estilos["h2"])]

    story.append(Paragraph(
        "Para presentar un panorama general de la calidad del aire en el AMG, se contabilizan los días "
        "en los que todos los contaminantes criterio se mantuvieron dentro de las categorías Buena o "
        "Aceptable del Índice Aire y Salud.",
        estilos["cuerpo"],
    ))

    serie = datos["serie_dias_buena_aceptable"]
    story.append(Paragraph(_texto_panorama_general(anio, serie), estilos["cuerpo"]))

    story.append(KeepTogether([
        Paragraph(
            "Figura 2. Días con calidad del aire Buena o Aceptable en el AMG.",
            estilos["tabla_caption"],
        ),
        _grafica_dias_buena_aceptable(serie),
    ]))

    return story


def _texto_panorama_general(anio, serie):
    """Arma el párrafo de resultados a partir de la serie histórica, sin
    valores fijos por año: el texto se recalcula con los datos que traiga
    'serie' para el año del informe y los años que lo acompañan."""
    valores_por_anio = dict(serie)
    anio_anterior = anio - 1
    valor_actual = valores_por_anio.get(anio)
    valor_anterior = valores_por_anio.get(anio_anterior)

    anio_max, valor_max = max(serie, key=lambda t: t[1])
    anio_min, valor_min = min(serie, key=lambda t: t[1])
    anio_ini, anio_fin = min(a for a, _ in serie), max(a for a, _ in serie)

    frase_comparacion = ""
    if valor_actual is not None and valor_anterior:
        diferencia = valor_actual - valor_anterior
        variacion = "más" if diferencia >= 0 else "menos"
        porcentaje = round(abs(diferencia) / valor_anterior * 100)
        frase_comparacion = (
            f", lo que representa {abs(diferencia)} días {variacion} que en {anio_anterior}, "
            f"equivalentes a un{'a disminución' if diferencia < 0 else ' incremento'} aproximado de {porcentaje} %"
        )

    return (
        f"Durante {anio} se registraron {valor_actual} días con calidad del aire Buena o Aceptable en "
        f"el AMG{frase_comparacion}. La Figura 2 muestra la evolución anual de este indicador entre "
        f"{anio_ini} y {anio_fin}. En este periodo, el mayor número de días con calidad del aire Buena "
        f"o Aceptable se registró en {anio_max}, con {valor_max} días, mientras que {anio_min} presentó "
        f"el valor más bajo, con {valor_min} días."
    )


def _grafica_dias_buena_aceptable(serie):
    """Dibuja la Figura 2: círculos y línea con los días Buena/Aceptable por
    año, en el mismo orden (reciente -> antiguo) en que se lista 'serie'."""
    ancho = CONTENT_WIDTH
    alto = 6.4 * cm
    pad_izq, pad_der, pad_arriba, pad_abajo = 30, 16, 18, 30
    plot_w = ancho - pad_izq - pad_der
    plot_h = alto - pad_arriba - pad_abajo

    valores = [v for _, v in serie]
    y_max = (max(valores) // 20 + 2) * 20
    n = len(serie)
    inset_puntos = 20  # deja espacio para que el primer/último círculo no tape las etiquetas del eje Y

    def x_de(i):
        ancho_util = plot_w - 2 * inset_puntos
        return pad_izq + inset_puntos + (ancho_util * i / (n - 1) if n > 1 else ancho_util / 2)

    def y_de(v):
        return pad_abajo + plot_h * (v / y_max)

    d = Drawing(ancho, alto)

    ticks = int(y_max // 20) + 1
    for t in range(ticks):
        valor_tick = t * 20
        y = y_de(valor_tick)
        d.add(Line(pad_izq, y, ancho - pad_der, y, strokeColor=colors.HexColor("#E1E7EA"), strokeWidth=0.6))
        d.add(String(pad_izq - 8, y - 3, str(valor_tick), fontName="Montserrat", fontSize=8,
                     fillColor=MUTED, textAnchor="end"))

    puntos = [(x_de(i), y_de(v)) for i, (_, v) in enumerate(serie)]
    for (x1, y1), (x2, y2) in zip(puntos, puntos[1:]):
        d.add(Line(x1, y1, x2, y2, strokeColor=AQUA, strokeWidth=2))

    for (anio_i, valor), (x, y) in zip(serie, puntos):
        d.add(Circle(x, y, 13, fillColor=colors.HexColor("#CDEFEA"), strokeColor=AQUA, strokeWidth=1.2))
        d.add(String(x, y - 3.5, str(valor), fontName="Montserrat-Bold", fontSize=9.5,
                     fillColor=NAVY, textAnchor="middle"))
        d.add(String(x, pad_abajo - 16, str(anio_i), fontName="Montserrat-SemiBold", fontSize=9,
                     fillColor=TEXT, textAnchor="middle"))

    return d


def _seccion_horas_categoria(anio, estilos, datos):
    story = [Paragraph("Horas del año según categoría de calidad del aire", estilos["h2"])]

    story.append(Paragraph(
        "Clasificación horaria por estación, según el Índice Aire y Salud dominante entre los seis "
        "contaminantes criterio (el contaminante que en cada hora presenta la categoría más "
        "desfavorable). Cada barra representa el porcentaje de horas del año que la estación "
        "registró en cada categoría; en gris se señalan las horas sin dato suficiente (D.I.), ya "
        "sea por falla de algún equipo o, en el caso de las estaciones incorporadas en 2024, por no "
        "contar todavía con un ciclo anual completo de operación. Más adelante el informe desglosa "
        "el comportamiento de cada contaminante por separado.",
        estilos["cuerpo"],
    ))

    horas = datos.get("horas_categoria_calidad", {})
    if not horas:
        story.append(Paragraph("[Faltan los datos de horas por categoría]", estilos["caption"]))
        return story

    story.append(Paragraph(_texto_horas_categoria(anio, horas), estilos["cuerpo"]))

    story.append(KeepTogether([
        Paragraph(
            "Figura 3. Horas del año según categoría de calidad del aire, por estación.",
            estilos["tabla_caption"],
        ),
        _grafica_horas_categoria(horas),
    ]))
    story.append(Spacer(1, 10))
    story.append(_leyenda_categorias_ias())

    return story


UMBRAL_DI_SUFICIENTE = 50  # % máximo de horas D.I. para considerar una estación comparable en este resumen


def _texto_horas_categoria(anio, datos_estacion, orden_estaciones=ORDEN_ESTACIONES_HORAS):
    """Arma el párrafo que resume la Figura 3, calculando directamente de
    'datos_estacion' qué estación tuvo más horas favorables (Buena o
    Aceptable) y cuál más horas desfavorables (Mala o peor), y señalando las
    estaciones cuyos datos son insuficientes para esa comparación. No hay
    valores fijos: si cambian los datos del año, cambia el texto."""
    filas = []
    for est in orden_estaciones:
        cats = datos_estacion.get(est, {})
        di = cats.get("D.I.", 0)
        favorable = cats.get("Buena", 0) + cats.get("Aceptable", 0)
        desfavorable = cats.get("Mala", 0) + cats.get("Muy mala", 0) + cats.get("Extremadamente mala", 0)
        filas.append((est, favorable, desfavorable, di))

    frase_intro = (
        f"La Figura 3 muestra, para cada una de las {len(orden_estaciones)} estaciones de la red, el "
        "porcentaje de horas del año que se clasificó en cada categoría del Índice Aire y Salud, "
        "según el contaminante con la categoría más desfavorable en cada hora."
    )

    incompletas = [f[0] for f in filas if f[3] >= UMBRAL_DI_SUFICIENTE]
    frase_incompletas = ""
    if incompletas:
        if len(incompletas) == 1:
            lista = incompletas[0]
        else:
            lista = ", ".join(incompletas[:-1]) + " y " + incompletas[-1]
        verbo = "concentra" if len(incompletas) == 1 else "concentran"
        frase_incompletas = (
            f" Las estaciones {lista} {verbo} una proporción alta de horas sin dato suficiente (D.I.), "
            f"varias de ellas por haberse incorporado a la red a mediados de {anio} y no contar "
            "todavía con un ciclo anual completo de mediciones."
        )

    suficientes = [f for f in filas if f[3] < UMBRAL_DI_SUFICIENTE]
    if not suficientes:
        return (
            f"{frase_intro} La mayoría de las estaciones no cuentan todavía con suficientes horas "
            f"válidas en el año para comparar su desempeño.{frase_incompletas}"
        )

    mejor = max(suficientes, key=lambda f: f[1])
    peor = max(suficientes, key=lambda f: f[2])

    return (
        f"{frase_intro} Entre las estaciones con datos suficientes durante {anio}, {mejor[0]} registró "
        f"la mayor proporción de horas en categoría Buena o Aceptable ({mejor[1]:.0f} % del año), "
        f"mientras que {peor[0]} presentó la mayor proporción de horas en categoría Mala o peor "
        f"({peor[2]:.0f} % del año).{frase_incompletas}"
    )


def _grafica_horas_categoria(datos_estacion, orden_estaciones=ORDEN_ESTACIONES_HORAS):
    """Barras horizontales 100% apiladas: para cada estación, qué porcentaje
    de horas del año cayó en cada categoría IAS (o D.I.)."""
    ancho = CONTENT_WIDTH
    etiqueta_w = 34
    fila_h = 17
    espacio = 6
    pad_arriba, pad_abajo = 4, 4
    n = len(orden_estaciones)
    alto = pad_arriba + n * fila_h + (n - 1) * espacio + pad_abajo
    barra_w = ancho - etiqueta_w

    d = Drawing(ancho, alto)

    for i, est in enumerate(orden_estaciones):
        y = alto - pad_arriba - (i + 1) * fila_h - i * espacio
        d.add(String(etiqueta_w - 8, y + fila_h / 2 - 3, est, fontName="Montserrat-Bold",
                     fontSize=8.4, fillColor=NAVY, textAnchor="end"))

        valores = datos_estacion.get(est, {})
        x = etiqueta_w
        for cat in CAT_ORDEN_HORAS:
            pct = valores.get(cat, 0)
            if pct <= 0:
                continue
            ancho_seg = barra_w * pct / 100
            d.add(Rect(x, y, ancho_seg, fila_h, fillColor=COLOR_CATEGORIA_IAS[cat],
                      strokeColor=colors.white, strokeWidth=0.6))
            x += ancho_seg

    return d


def _leyenda_categorias_ias():
    fuente, tam = "Montserrat", 8.5
    swatch = 10
    gap_swatch_texto = 4
    gap_items = 16
    alto = 14

    items = []
    x = 0.0
    for cat in CAT_ORDEN_HORAS:
        ancho_texto = pdfmetrics.stringWidth(cat, fuente, tam)
        items.append((cat, x, ancho_texto))
        x += swatch + gap_swatch_texto + ancho_texto + gap_items
    ancho_total = x - gap_items

    d = Drawing(ancho_total, alto)
    for cat, x0, ancho_texto in items:
        d.add(Rect(x0, alto / 2 - swatch / 2, swatch, swatch, fillColor=COLOR_CATEGORIA_IAS[cat], strokeColor=None))
        d.add(String(x0 + swatch + gap_swatch_texto, alto / 2 - 3, cat, fontName=fuente, fontSize=tam, fillColor=TEXT))
    return d


def _seccion_cumplimiento_nom(anio, estilos, datos):
    story = [Paragraph(f"Cumplimiento de las NOM de calidad del aire, {anio}", estilos["h2"])]

    story.append(Paragraph(
        "Para cada estación de la red se compara el valor estadístico correspondiente (percentil 99, "
        "promedio o máximo, según lo que establece cada Norma Oficial Mexicana) contra el límite "
        f"vigente para {anio}. Cuando una estación no alcanza el 75&nbsp;% de días válidos en el año, "
        "el resultado se reporta como dato insuficiente (D.I.).",
        estilos["cuerpo"],
    ))

    filas_nom = datos.get("cumplimiento_nom")
    if not filas_nom:
        story.append(Paragraph("[Faltan los datos de cumplimiento de NOM]", estilos["caption"]))
        return story

    story.append(KeepTogether([
        Paragraph(
            "Tabla 5. Cumplimiento de las NOM de calidad del aire, por estación.",
            estilos["tabla_caption"],
        ),
        _tabla_cumplimiento_nom(datos["estaciones"], filas_nom, estilos),
        Spacer(1, 8),
        _leyenda_cumplimiento_nom(),
        Paragraph(
            "D.I.: la estación no alcanzó el 75&nbsp;% de días válidos requerido en el año. "
            "FO: la estación cuenta con el equipo pero permaneció fuera de operación todo el año. "
            "¤: la estación no mide ese contaminante.",
            estilos["nota"],
        ),
    ]))

    return story


def _formatear_valor_nom(valor, fila_nom):
    if valor is None:
        return ""
    if fila_nom["unidad"] == "µg/m³":
        return f"{valor:.0f}"
    if fila_nom["contaminante"] == "CO":
        return f"{valor:.1f}"
    return f"{valor:.3f}"


def _celda_nom_contenido(resultado, fila_nom, contaminantes_capacidad, estilos):
    """Devuelve (texto, color_fondo, estilo) para una celda de la Tabla 5,
    resolviendo aquí (no en calculo_datos.py) si un "sin_datos" es en
    realidad una estación fuera de operación (FO) o sin ese equipo (¤)."""
    status = resultado.get("status")
    valor = resultado.get("valor")

    if status == "sin_datos":
        clave_capacidad = CAPACIDAD_POR_CONTAMINANTE[fila_nom["contaminante"]]
        tiene_equipo = (contaminantes_capacidad or {}).get(clave_capacidad, True)
        status = "FO" if tiene_equipo else "sin_equipo"

    if status == "cumple":
        return _formatear_valor_nom(valor, fila_nom), COLOR_NOM_CUMPLE, estilos["nom_celda_clara"]
    if status == "no_cumple":
        return _formatear_valor_nom(valor, fila_nom), COLOR_NOM_NO_CUMPLE, estilos["nom_celda_clara"]
    if status == "DI":
        return "D.I.", COLOR_NOM_DI, estilos["nom_celda_oscura"]
    if status == "FO":
        return "FO", COLOR_NOM_FO, estilos["nom_celda_oscura"]
    return "¤", COLOR_NOM_SIN_EQUIPO, estilos["nom_celda_oscura"]


def _tabla_cumplimiento_nom(estaciones, filas_nom, estilos):
    """Tabla 5: filas = parámetro NOM (contaminante + periodo), columnas =
    estaciones agrupadas por municipio (mismo orden que la Tabla 1)."""
    claves = [e["simbolo"] for e in estaciones]
    municipios = [e["municipio"] for e in estaciones]
    grupos_municipio = _spans_de_grupos(municipios)
    equipo_por_clave = {e["simbolo"]: e["contaminantes"] for e in estaciones}

    col_etiqueta = CONTENT_WIDTH * 0.185
    col_estacion = (CONTENT_WIDTH - col_etiqueta) / len(claves)

    fila_header_1 = [Paragraph("Parámetro · Periodo", estilos["tabla_header"])]
    for _ in claves:
        fila_header_1.append("")
    for (i0, i1) in grupos_municipio:
        etiqueta_municipio = MUNICIPIO_CORTO.get(municipios[i0], municipios[i0])
        fila_header_1[1 + i0] = Paragraph(etiqueta_municipio, estilos["tabla_subheader_chico"])
    fila_header_2 = [""] + [Paragraph(c, estilos["tabla_subheader"]) for c in claves]

    data = [fila_header_1, fila_header_2]
    fondos = []  # (fila, col, color_fondo) para las celdas de resultado
    for r, fila_nom in enumerate(filas_nom, start=2):
        etiqueta = Paragraph(
            f"{fila_nom['simbolo']} · {fila_nom['periodo']}<br/>"
            f"<font color='#6C8894' size=6.4>{fila_nom['limite_txt']} {fila_nom['unidad']}</font>",
            estilos["nom_etiqueta"],
        )
        renglon = [etiqueta]
        for c, clave in enumerate(claves, start=1):
            resultado = fila_nom["estaciones"].get(clave, {})
            texto, color_fondo, estilo_txt = _celda_nom_contenido(
                resultado, fila_nom, equipo_por_clave.get(clave), estilos,
            )
            renglon.append(Paragraph(texto, estilo_txt))
            fondos.append((r, c, color_fondo))
        data.append(renglon)

    anchos = [col_etiqueta] + [col_estacion] * len(claves)
    tabla = Table(data, colWidths=anchos, repeatRows=2)

    estilo_cmds = [
        ("BACKGROUND", (0, 0), (-1, 1), NAVY),
        ("SPAN", (0, 0), (0, 1)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D7DEE1")),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (1, 0), (-1, -1), 2),
        ("RIGHTPADDING", (1, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (0, -1), 6),
        ("RIGHTPADDING", (0, 0), (0, -1), 6),
        ("LEFTPADDING", (1, 0), (-1, 0), 1),
        ("RIGHTPADDING", (1, 0), (-1, 0), 1),
    ]

    for (i0, i1) in grupos_municipio:
        if i1 > i0:
            estilo_cmds.append(("SPAN", (1 + i0, 0), (1 + i1, 0)))

    for r, c, color_fondo in fondos:
        estilo_cmds.append(("BACKGROUND", (c, r), (c, r), color_fondo))

    tabla.setStyle(TableStyle(estilo_cmds))
    return tabla


def _leyenda_cumplimiento_nom():
    fuente, tam = "Montserrat", 8.5
    swatch = 10
    gap_swatch_texto = 4
    gap_items = 16
    alto = 14

    items_leyenda = [
        ("Cumple", COLOR_NOM_CUMPLE),
        ("No cumple", COLOR_NOM_NO_CUMPLE),
        ("D.I.", COLOR_NOM_DI),
        ("FO", COLOR_NOM_FO),
        ("¤ Sin equipo", COLOR_NOM_SIN_EQUIPO),
    ]

    items = []
    x = 0.0
    for texto, color_fondo in items_leyenda:
        ancho_texto = pdfmetrics.stringWidth(texto, fuente, tam)
        items.append((texto, color_fondo, x, ancho_texto))
        x += swatch + gap_swatch_texto + ancho_texto + gap_items
    ancho_total = x - gap_items

    d = Drawing(ancho_total, alto)
    for texto, color_fondo, x0, ancho_texto in items:
        d.add(Rect(x0, alto / 2 - swatch / 2, swatch, swatch, fillColor=color_fondo,
                  strokeColor=colors.HexColor("#B7C2C7"), strokeWidth=0.5))
        d.add(String(x0 + swatch + gap_swatch_texto, alto / 2 - 3, texto, fontName=fuente, fontSize=tam, fillColor=TEXT))
    return d


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
    story += _seccion_panorama_general(anio, estilos, datos)
    story += _seccion_horas_categoria(anio, estilos, datos)
    story += _seccion_cumplimiento_nom(anio, estilos, datos)

    dibujar_portada = _dibujar_portada(anio, datos)
    pintar_fondo = _fondo_pagina(anio)
    doc.build(story, onFirstPage=dibujar_portada, onLaterPages=pintar_fondo)
    return salida


def _asegurar_datos_anio(anio, forzar=False):
    """Si el año pedido no está todavía en datos/resumen_historico.json (o se
    pide --recalcular), descarga BD_{anio} de Google Sheets y calcula sus
    cifras con calculo_datos.py, antes de armar el PDF. Así "python
    generar_informe.py --anio 2024" funciona con un solo comando, aunque sea
    la primera vez que se pide ese año."""
    historico = _cargar_resumen_historico()
    if not forzar and str(anio) in historico:
        return

    try:
        import calculo_datos
    except ImportError as exc:
        print(f"Aviso: no se pudo importar calculo_datos.py ({exc}). Se usarán los valores de ejemplo si existen.")
        return

    try:
        calculo_datos.actualizar_historico(anio)
    except Exception as exc:
        print(
            f"Aviso: no se pudieron descargar/calcular los datos reales de {anio} ({exc}). "
            "Se usarán los valores de ejemplo si existen para ese año."
        )


def _main():
    parser = argparse.ArgumentParser(description="Genera el Informe Anual de Calidad del Aire de Jalisco.")
    parser.add_argument("--anio", type=int, required=True, help="Año del informe, p. ej. 2024")
    parser.add_argument("--salida", type=str, default=None, help="Ruta del PDF de salida")
    parser.add_argument(
        "--recalcular", action="store_true",
        help="Vuelve a descargar BD_{anio} y a calcular sus cifras aunque el año ya esté en el histórico",
    )
    parser.add_argument(
        "--sin-datos-nuevos", action="store_true",
        help="No descarga ni calcula nada nuevo; usa solo lo que ya haya en datos/resumen_historico.json",
    )
    args = parser.parse_args()

    if not args.sin_datos_nuevos:
        _asegurar_datos_anio(args.anio, forzar=args.recalcular)

    ruta = generar_informe(args.anio, args.salida)
    print(f"PDF generado en: {ruta}")


if __name__ == "__main__":
    _main()
