"""
Cálculo automático de los datos anuales del informe, a partir de la base de
datos horaria de SIMAJ (BD_{anio} en Google Sheets).

Es un port directo de las funciones de tu notebook de diagnóstico (IAS
NOM-172-SEMARNAT-2023): mismas fórmulas, mismos umbrales de suficiencia,
mismo criterio de clasificación IAS (RANGOS de PM10/PM2.5 vigente en tu
script, el marcado como "criterio de PM para 2026"). No se reinterpreta ni
se ajusta nada: es el mismo cálculo, corriendo aquí en vez de en Colab.

Uso:
    python calculo_datos.py --anio 2024
    (descarga BD_2024, calcula, y actualiza datos/resumen_historico.json)
"""
import argparse
import calendar
import json
import re
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import numpy as np
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
RUTA_RESUMEN_HISTORICO = BASE_DIR / "datos" / "resumen_historico.json"

# ---------------------------------------------------------------------------
# Bases de datos por año (Google Sheets de SIMAJ, compartidas como "cualquiera
# con el link puede ver"). Se descargan vía el endpoint de exportación a xlsx.
# ---------------------------------------------------------------------------
SHEETS_POR_ANIO = {
    2024: "1MuEqhq4XUXP9ViIy_QSHkL1kpJwXe40T",
    2023: "15A5ttGasHk3tbAwgv3-ru82X3ANDimIm",
    2022: "1-AYGOl2Ya_YAa3s555qr9kzHVEGfolOr",
    2021: "1Cf3fvYjsFOi7WIHDhaE8Ly1Z_A4PDERA",
    2020: "1oZ3f6fBlDeXMcFuuGJjE0gW3N3aVoP_8",
    2019: "1CEHHu-H_5zSfoA7BXXOcStxbrCJE0K74",
}


def _extraer_id_hoja(url_o_id):
    """Acepta tanto un ID de Google Sheets como su URL completa de edición."""
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url_o_id)
    return m.group(1) if m else url_o_id


def descargar_bd(anio, destino=None, url_o_id=None):
    """Descarga BD_{anio} (hoja 'Data') como .xlsx. No requiere credenciales
    porque las hojas están compartidas por enlace."""
    id_hoja = _extraer_id_hoja(url_o_id) if url_o_id else SHEETS_POR_ANIO.get(anio)
    if not id_hoja:
        raise ValueError(f"No tengo el ID de la hoja de Google para {anio}. Pásalo con --sheet.")

    if destino is None:
        destino = BASE_DIR / "datos" / "_cache" / f"BD_{anio}.xlsx"
    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)

    import urllib.request
    url = f"https://docs.google.com/spreadsheets/d/{id_hoja}/export?format=xlsx"
    print(f"Descargando BD_{anio} de Google Sheets…")
    urllib.request.urlretrieve(url, destino)
    print(f"  guardado en {destino}")
    return destino


# ---------------------------------------------------------------------------
# Funciones IAS (idénticas a las del notebook: round_half_up, NowCast,
# clasifica, RANGOS...)
# ---------------------------------------------------------------------------
def round_half_up(value, ndigits):
    if value is None or pd.isna(value):
        return np.nan
    q = Decimal(str(value)).quantize(Decimal("1." + "0" * ndigits), rounding=ROUND_HALF_UP)
    return float(q)


def rolling_8h(series):
    return series.rolling(8, min_periods=6).mean()


def rolling_24h(series):
    return series.rolling(24, min_periods=18).mean()


def NowCast(valores, PM):
    ultimas_3 = valores[-3:] if len(valores) >= 3 else valores
    if sum(1 for x in ultimas_3 if x is not None) < 2:
        return None

    valores_inv = valores[::-1]
    vals_indexados = []
    hora = 0
    for v in valores_inv:
        if v is not None:
            vals_indexados.append((float(v), hora))
        hora += 1

    if len(vals_indexados) < 2:
        return None

    solo_valores = [v for v, _ in vals_indexados]
    if all(v == 0 for v in solo_valores):
        return 0

    v_max = max(solo_valores)
    v_min = min(solo_valores)
    rango = v_max - v_min

    if v_max == 0:
        return 0
    tasa = round_half_up(1 - (rango / v_max), 2)
    factor = tasa if tasa >= 0.5 else 0.5

    num = 0.0
    den = 0.0
    for v, i in vals_indexados:
        peso = factor ** i
        num += v * peso
        den += peso

    if den == 0:
        return None

    promedio_ponderado = round_half_up(num / den, 0)

    if PM == 0:
        promedio_ponderado = round_half_up(promedio_ponderado * 0.714, 0)
    else:
        promedio_ponderado = round_half_up(promedio_ponderado * 0.694, 0)

    return int(promedio_ponderado)


def serie_nowcast_por_estacion(df, pol_col, pm_flag):
    s = pd.to_numeric(df[pol_col], errors="coerce")
    out = []
    for idx in range(len(s)):
        start = max(0, idx - 11)
        ventana = s.iloc[start:idx + 1]
        vals = [(None if pd.isna(v) else float(v)) for v in ventana.tolist()]
        out.append(NowCast(vals, pm_flag))
    return pd.Series(out, index=df.index, name=f"{pol_col}_NOWCAST")


CAT_ORDER = ["Buena", "Aceptable", "Mala", "Muy mala", "Extremadamente mala"]
CAT_PUNTAJE = {c: i + 1 for i, c in enumerate(CAT_ORDER)}

# Mismo RANGOS activo en tu notebook (criterio de PM vigente para 2026,
# aplicado de forma consistente a todos los años para poder comparar la
# tendencia histórica bajo un mismo criterio).
RANGOS = {
    "PM10": [(None, 45, "Buena"), (45, 50, "Aceptable"), (50, 132, "Mala"),
             (132, 213, "Muy mala"), (213, None, "Extremadamente mala")],
    "PM2.5": [(None, 15, "Buena"), (15, 25, "Aceptable"), (25, 79, "Mala"),
              (79, 130, "Muy mala"), (130, None, "Extremadamente mala")],
    "O3": [(None, 0.058, "Buena"), (0.058, 0.090, "Aceptable"), (0.090, 0.135, "Mala"),
           (0.135, 0.175, "Muy mala"), (0.175, None, "Extremadamente mala")],
    "NO2": [(None, 0.053, "Buena"), (0.053, 0.106, "Aceptable"), (0.106, 0.160, "Mala"),
            (0.160, 0.213, "Muy mala"), (0.213, None, "Extremadamente mala")],
    "SO2": [(None, 0.035, "Buena"), (0.035, 0.075, "Aceptable"), (0.075, 0.185, "Mala"),
            (0.185, 0.304, "Muy mala"), (0.304, None, "Extremadamente mala")],
    "CO": [(None, 5.00, "Buena"), (5.00, 9.00, "Aceptable"), (9.00, 12.00, "Mala"),
           (12.00, 16.00, "Muy mala"), (16.00, None, "Extremadamente mala")],
}


def clasifica(valor, pol):
    if valor is None or pd.isna(valor):
        return None, None
    for lo, hi, cat in RANGOS[pol]:
        lo_ok = (lo is None) or (valor > lo)
        hi_ok = (hi is None) or (valor <= hi)
        if lo_ok and hi_ok:
            return cat, CAT_PUNTAJE[cat]
    return None, None


# ---------------------------------------------------------------------------
# Sección 1: carga y limpieza (igual que load_and_prepare_db)
# ---------------------------------------------------------------------------
CONTAMINANTES = ['O3', 'NO', 'NO2', 'NOX', 'SO2', 'CO', 'PM10', 'PM2.5']
METEOROLOGIA = ['IT', 'ET', 'RH', 'WS', 'WD', 'PP', 'ATM', 'RS', 'UVI']
INVALID_FLAGS = ["IF", "IO", "IR", "ND", "VE", "SE", "NE", "IC", "VZ", ""]

RANGO_MAX = {
    "O3": 0.475, "NO": 0.475, "NO2": 0.475, "NOX": 0.475, "SO2": 0.475,
    "CO": 47, "PM10": 950, "PM2.5": 950,
    "IT": 60, "ET": 50, "RH": 100, "WS": 50, "WD": 360, "PP": 10,
    "ATM": 760, "RS": 2000, "UVI": 300,
}


def _fix_small_neg(s, min_floor=-0.006):
    s = s.copy()
    s.loc[(s < 0) & (s >= min_floor)] = 0
    s.loc[s < min_floor] = np.nan
    return s


def validate_ranges(df, contaminantes, meteorologia):
    for c in contaminantes + meteorologia:
        if c not in df.columns:
            df[c] = np.nan
        df.loc[df[c] > RANGO_MAX[c], c] = np.nan

    for c in contaminantes:
        df[c] = _fix_small_neg(pd.to_numeric(df[c], errors="coerce"))

    df['ET'] = pd.to_numeric(df['ET'], errors='coerce')
    df.loc[(df['ET'] < 0) & (df['ET'] >= -0.006), 'ET'] = 0
    df.loc[df['ET'] < -5, 'ET'] = np.nan
    df.loc[df['ET'] > RANGO_MAX['ET'], 'ET'] = np.nan

    df['IT'] = _fix_small_neg(pd.to_numeric(df['IT'], errors='coerce'))
    df.loc[df['IT'] > RANGO_MAX['IT'], 'IT'] = np.nan

    df['RH'] = pd.to_numeric(df['RH'], errors='coerce').clip(lower=0, upper=100)

    df['WD'] = pd.to_numeric(df['WD'], errors='coerce')
    df.loc[(df['WD'] < 0) | (df['WD'] > 360), 'WD'] = np.nan

    for c in ['WS', 'PP', 'RS', 'UVI', 'ATM']:
        df[c] = _fix_small_neg(pd.to_numeric(df[c], errors='coerce'))
        df.loc[df[c] > RANGO_MAX[c], c] = np.nan

    return df


def load_and_prepare_db(ruta_excel, hoja_datos="Data"):
    df = pd.read_excel(ruta_excel, sheet_name=hoja_datos, engine="openpyxl")
    df.columns = df.columns.str.strip().str.replace(' +', '_', regex=True)

    df['STATION'] = df['STATION'].str.strip()
    df['STATION'] = df['STATION'].apply(lambda x: re.sub(r'[^A-Za-z0-9 ]+', '', str(x)))
    df['STATION'] = df['STATION'].str.strip()

    df.replace(INVALID_FLAGS, np.nan, inplace=True)

    for c in CONTAMINANTES + METEOROLOGIA:
        if c not in df.columns:
            df[c] = np.nan
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df = validate_ranges(df, CONTAMINANTES, METEOROLOGIA)

    df["DATE"] = pd.to_datetime(df["DATE"], errors="coerce")
    df = df.dropna(subset=["DATE"])
    df = df.sort_values(["STATION", "DATE"]).reset_index(drop=True)

    years, counts = np.unique(df["DATE"].dt.year.values, return_counts=True)
    anio = int(years[np.argmax(counts)])
    return df, anio


# ---------------------------------------------------------------------------
# Rolling / NowCast por estación, previos a la agregación diaria (igual que
# la parte necesaria de tu Sección 2)
# ---------------------------------------------------------------------------
def calcular_columnas_rolling(dfh):
    dfh = dfh.copy()
    if "PM10" in dfh.columns:
        dfh["PM10_NOWCAST"] = (dfh.groupby("STATION", group_keys=False)
                                .apply(lambda g: serie_nowcast_por_estacion(g, "PM10", 0)))
    if "PM2.5" in dfh.columns:
        dfh["PM2.5_NOWCAST"] = (dfh.groupby("STATION", group_keys=False)
                                 .apply(lambda g: serie_nowcast_por_estacion(g, "PM2.5", 1)))
    if "CO" in dfh.columns:
        dfh["CO_8H"] = (dfh.groupby("STATION")["CO"].apply(rolling_8h)
                        .reset_index(level=0, drop=True))
    if "O3" in dfh.columns:
        dfh["O3_8H"] = (dfh.groupby("STATION")["O3"].apply(rolling_8h)
                        .reset_index(level=0, drop=True))
    return dfh


# Orden de estaciones usado en tus gráficas (EST_ORDER_BASE del notebook).
EST_ORDER_BASE = ['AGU', 'ATM', 'CEN', 'COU', 'LDO', 'MIR', 'OBL', 'PIN', 'SAN', 'SFE', 'SMT', 'TLA', 'VAL']

# Columna horaria que se clasifica para IAS, por contaminante (igual que
# col_valor_por_pol en la Sección 2 de tu notebook). Nota que CO usa su
# promedio móvil de 8h y O3 su valor de 1h -- igual que en tu notebook, no
# es un error.
COL_VALOR_HORARIO_GLOBAL = {
    "PM10": "PM10_NOWCAST", "PM2.5": "PM2.5_NOWCAST",
    "CO": "CO_8H", "O3": "O3_1H", "NO2": "NO2_1H", "SO2": "SO2_1H",
}


def calcular_columnas_horarias_gases(dfh):
    """Redondeos NOM-172 de gases a nivel horario (O3_1H, NO2_1H, SO2_1H,
    CO_1H), necesarios -- junto con las columnas de calcular_columnas_rolling
    -- para poder calcular la categoría IAS GLOBAL por hora."""
    dfh = dfh.copy()
    if "O3" in dfh.columns:
        dfh["O3_1H"] = dfh["O3"].apply(lambda v: round_half_up(v, 3))
    if "NO2" in dfh.columns:
        dfh["NO2_1H"] = dfh["NO2"].apply(lambda v: round_half_up(v, 3))
    if "SO2" in dfh.columns:
        dfh["SO2_1H"] = dfh["SO2"].apply(lambda v: round_half_up(v, 3))
    if "CO" in dfh.columns:
        dfh["CO_1H"] = dfh["CO"].apply(lambda v: round_half_up(v, 2))
    return dfh


def calcular_ias_global_horario(dfh):
    """
    Port de "IAS Global (contaminante dominante por fila)" de la Sección 2 de
    tu notebook (decide_dominante), aplicado hora por hora: para cada hora,
    el contaminante con peor categoría IAS "gana"; los empates se resuelven
    por qué tan adentro de su categoría está (fraccion_en_rango) y luego por
    el orden conservador ORDEN_DOM. Agrega la columna IAS_GLOBAL_CAT.
    """
    dfh = dfh.copy()
    info = {}
    for pol, colv in COL_VALOR_HORARIO_GLOBAL.items():
        if colv not in dfh.columns:
            continue
        clasif = dfh[colv].apply(lambda v: clasifica(v, pol))
        info[pol] = {
            "valor": dfh[colv].to_numpy(),
            "cat": clasif.apply(lambda t: t[0]).to_numpy(),
            "score": clasif.apply(lambda t: t[1]).to_numpy(),
        }

    n = len(dfh)
    categoria_global = [None] * n
    for i in range(n):
        candidatos = []
        for pol in ORDEN_DOM:
            if pol not in info:
                continue
            sc = info[pol]["score"][i]
            cat = info[pol]["cat"][i]
            if sc is None or pd.isna(sc) or cat is None:
                continue
            frac = _frac_rango(info[pol]["valor"][i], pol, cat)
            candidatos.append((pol, float(sc), float(frac), cat))
        if not candidatos:
            continue
        candidatos.sort(key=lambda t: (t[1], t[2], -ORDEN_DOM.index(t[0])), reverse=True)
        categoria_global[i] = candidatos[0][3]

    dfh["IAS_GLOBAL_CAT"] = categoria_global
    return dfh


def calcular_frecuencia_horaria_global(dfh, estaciones=EST_ORDER_BASE):
    """
    Cuenta, por estación, cuántas horas del año cayeron en cada categoría
    IAS GLOBAL (o "D.I." si ningún contaminante tuvo dato válido esa hora).
    No incluye la estación virtual AMG. Requiere que dfh ya tenga la columna
    IAS_GLOBAL_CAT (ver calcular_ias_global_horario).

    Regresa un DataFrame largo: STATION, CATEGORIA, FRECUENCIA (horas),
    PORCENTAJE (sobre el total de horas del año para esa estación).
    """
    dfh = dfh[dfh["STATION"].isin(estaciones)].copy()
    cat_order_di = CAT_ORDER + ["D.I."]
    base = dfh.groupby("STATION").size().reset_index(name="BASE")

    tmp = pd.DataFrame({"STATION": dfh["STATION"], "CATEGORIA": dfh["IAS_GLOBAL_CAT"]}).dropna(subset=["CATEGORIA"])
    g = tmp.groupby(["STATION", "CATEGORIA"]).size().reset_index(name="FRECUENCIA")

    clasificadas = g.groupby("STATION")["FRECUENCIA"].sum().reset_index(name="CLAS")
    di = base.merge(clasificadas, on="STATION", how="left")
    di["CLAS"] = di["CLAS"].fillna(0).astype(int)
    di["FRECUENCIA"] = (di["BASE"] - di["CLAS"]).clip(lower=0)
    di["CATEGORIA"] = "D.I."

    g_all = pd.concat([g, di[["STATION", "CATEGORIA", "FRECUENCIA"]]], ignore_index=True)
    g_all = g_all.merge(base, on="STATION", how="left")
    g_all["PORCENTAJE"] = np.where(g_all["BASE"] > 0, 100 * g_all["FRECUENCIA"] / g_all["BASE"], 0.0)
    g_all["CATEGORIA"] = pd.Categorical(g_all["CATEGORIA"], categories=cat_order_di, ordered=True)
    return g_all.sort_values(["STATION", "CATEGORIA"])


def agregar_amg_horario(dfh, estaciones=EST_ORDER_BASE):
    """Agrega la estación virtual AMG a nivel horario: para cada hora y cada
    columna de COL_VALOR_HORARIO_GLOBAL toma el máximo entre estaciones (igual
    que rebuild_amg_from_daily lo hace por día) y clasifica con el mismo IAS
    global. Devuelve dfh con las filas extra STATION == "AMG"."""
    cols = [c for c in COL_VALOR_HORARIO_GLOBAL.values() if c in dfh.columns]
    amg = dfh[dfh["STATION"].isin(estaciones)].groupby("DATE")[cols].max().reset_index()
    amg["STATION"] = "AMG"
    amg = calcular_ias_global_horario(amg)
    return pd.concat([dfh, amg[["STATION", "DATE", "IAS_GLOBAL_CAT"]]], ignore_index=True)


def resumen_horas_categoria(dfh, estaciones=EST_ORDER_BASE):
    """Convierte calcular_frecuencia_horaria_global() en el formato que
    guarda resumen_historico.json: {estacion: {categoria: %}}."""
    freq = calcular_frecuencia_horaria_global(dfh, estaciones=estaciones)
    salida = {}
    for est in estaciones:
        fila_est = freq[freq["STATION"] == est]
        salida[est] = {
            str(cat): round(float(pct), 2)
            for cat, pct in zip(fila_est["CATEGORIA"], fila_est["PORCENTAJE"])
        }
    return salida


def calcular_violines_mensuales(dfh, columna, suavizado=None, estaciones=EST_ORDER_BASE, minimo_horas=0.75):
    """Estadísticos por mes para la gráfica de violines del AMG (sección
    "Gráfica de violín por mes" del notebook). Un mes solo cuenta si tiene
    >= 75 % de sus horas con dato; si no, es D.I. Por mes: caja (Q1, mediana,
    Q3, bigotes a 1.5 IQR, atípicos), media y densidad kernel gaussiana
    (bandwidth de Scott, como matplotlib) normalizada a máximo 1.
    Regresa {mes: {...}} en la unidad original."""
    # AMG = máximo horario entre estaciones de la columna original; el promedio
    # móvil (p. ej. 8 h en CO) se calcula sobre esa serie del AMG, como en el
    # notebook (la columna CO_8H de la estación AMG).
    serie = dfh[dfh["STATION"].isin(estaciones)].groupby("DATE")[columna].max()
    if suavizado is not None:
        serie = suavizado(serie)
    salida = {}
    for mes in range(1, 13):
        del_mes = serie[serie.index.month == mes]
        dias = calendar.monthrange(int(serie.index.year[0]), mes)[1]
        vals = del_mes.dropna().to_numpy(dtype=float)
        if vals.size < minimo_horas * dias * 24 or vals.size < 3:
            salida[str(mes)] = {"valido": False}
            continue
        q1, med, q3 = (float(x) for x in np.percentile(vals, [25, 50, 75]))
        iqr = q3 - q1
        dentro = vals[(vals >= q1 - 1.5 * iqr) & (vals <= q3 + 1.5 * iqr)]
        atipicos = vals[(vals < q1 - 1.5 * iqr) | (vals > q3 + 1.5 * iqr)]
        bw = float(np.std(vals, ddof=1)) * vals.size ** (-1 / 5) or 1e-6
        grid = np.linspace(vals.min(), vals.max(), 60)
        dens = np.exp(-0.5 * ((grid[:, None] - vals[None, :]) / bw) ** 2).mean(axis=1)
        dens = dens / dens.max()
        salida[str(mes)] = {
            "valido": True, "n": int(vals.size), "media": round(float(vals.mean()), 4),
            "q1": round(q1, 4), "mediana": round(med, 4), "q3": round(q3, 4),
            "bigote_inf": round(float(dentro.min()), 4), "bigote_sup": round(float(dentro.max()), 4),
            "atipicos": [round(float(v), 3) for v in np.sort(atipicos)],
            "kde_x": [round(float(v), 4) for v in grid], "kde_y": [round(float(v), 3) for v in dens],
        }
    return salida


def calcular_perfil_horario(dfh, contaminante, estaciones=EST_ORDER_BASE, min_obs_por_hora=6, cobertura_mensual=0.75):
    """Perfil diurno del AMG para un contaminante (sección 4.4 del notebook):
    AMG = máximo horario entre estaciones; luego promedio por (mes, hora del
    día), exigiendo al menos `min_obs_por_hora` días con dato.
    Regresa {mes: {hora: promedio}} en la unidad original de la base."""
    dfh = dfh[dfh["STATION"].isin(estaciones)]
    amg = dfh.groupby("DATE")[contaminante].max().reset_index()
    amg["MES"] = amg["DATE"].dt.month
    amg["HORA"] = amg["DATE"].dt.hour
    perfil = (amg.groupby(["MES", "HORA"])[contaminante]
              .agg(MEAN="mean", N="count").reset_index())
    perfil.loc[perfil["N"] < min_obs_por_hora, "MEAN"] = np.nan

    # Un mes solo se grafica si el AMG tiene dato en >= 75 % de sus horas
    # (igual que en los violines); si no, el mes queda como D.I.
    anio = int(amg["DATE"].dt.year.iloc[0])
    for mes in range(1, 13):
        horas_mes = calendar.monthrange(anio, mes)[1] * 24
        con_dato = int(amg.loc[amg["MES"] == mes, contaminante].notna().sum())
        if con_dato < cobertura_mensual * horas_mes:
            perfil.loc[perfil["MES"] == mes, "MEAN"] = np.nan

    salida = {}
    for mes in range(1, 13):
        fila_mes = perfil[perfil["MES"] == mes].sort_values("HORA")
        salida[str(mes)] = {
            str(int(h)): (round(float(v), 4) if pd.notna(v) else None)
            for h, v in zip(fila_mes["HORA"], fila_mes["MEAN"])
        }
    return salida


# ---------------------------------------------------------------------------
# Sección 3: agregación diaria (igual que build_daily_table / rebuild_amg /
# compute_ias_daily / annual_compiled_by_station)
# ---------------------------------------------------------------------------
SUF_MIN_HORAS = 18  # 75% de 24h


def _round_by_nom(value, pol):
    if pd.isna(value):
        return np.nan
    if pol in ["O3", "NO2", "SO2"]:
        return round_half_up(value, 3)
    if pol == "CO":
        return round_half_up(value, 2)
    if pol in ["PM10", "PM2.5"]:
        return round_half_up(value, 0)
    return value


def _daily_bounds(df_day, pol):
    s = pd.to_numeric(df_day.get(pol, pd.Series(dtype=float)), errors="coerce")
    horas_validas = int(s.notna().sum())
    suf = horas_validas >= SUF_MIN_HORAS
    if not suf:
        return (np.nan, np.nan, False, horas_validas)
    avg24 = _round_by_nom(float(s.mean()), pol)
    mx1h = _round_by_nom(float(s.max()), pol)
    return (avg24, mx1h, True, horas_validas)


def _daily_max8h(df_day, pol):
    col = f"{pol}_8H"
    if col not in df_day:
        return np.nan
    s = pd.to_numeric(df_day[col], errors="coerce").dropna()
    if s.empty:
        return np.nan
    return _round_by_nom(float(s.max()), pol)


def _daily_nowcast_max(df_day, pol):
    col = f"{pol}_NOWCAST"
    if col not in df_day:
        return np.nan
    s = pd.to_numeric(df_day[col], errors="coerce").dropna()
    if s.empty:
        return np.nan
    return _round_by_nom(float(s.max()), pol)


def build_daily_table(dfh):
    dfh = dfh[dfh["STATION"] != "AMG"].copy()
    dfh["FECHA"] = pd.to_datetime(dfh["DATE"]).dt.date

    out = []
    for (est, fec), g in dfh.groupby(["STATION", "FECHA"]):
        row = {"STATION": est, "FECHA": pd.to_datetime(fec)}
        for pol in ['O3', 'NO2', 'SO2', 'CO', 'PM10', 'PM2.5']:
            if pol not in g.columns:
                continue
            avg24, mx1h, suf, hv = _daily_bounds(g, pol)
            row[f"{pol}_HORAS_VALIDAS"] = hv
            row[f"{pol}_AVG_24H"] = avg24
            row[f"{pol}_MAX_1H"] = mx1h
            row[f"{pol}_SUF_DIARIA"] = bool(suf)

        row["O3_MAX_8H"] = _daily_max8h(g, "O3")
        row["CO_MAX_8H"] = _daily_max8h(g, "CO")
        row["PM10_NOWCAST_MAX"] = _daily_nowcast_max(g, "PM10")
        row["PM2.5_NOWCAST_MAX"] = _daily_nowcast_max(g, "PM2.5")
        out.append(row)

    return pd.DataFrame(out).sort_values(["STATION", "FECHA"]).reset_index(drop=True)


def rebuild_amg_from_daily(dfd):
    cols_max = [c for c in dfd.columns if c.endswith(("_AVG_24H", "_MAX_1H", "_MAX_8H", "_NOWCAST_MAX"))]
    base = dfd.groupby("FECHA")[cols_max].max(min_count=1).reset_index()

    suf_amg = {}
    for pol in ["O3", "NO2", "SO2", "CO", "PM10", "PM2.5"]:
        ref = (f"{pol}_MAX_8H" if pol == "CO" else
               f"{pol}_AVG_24H" if pol in ("PM10", "PM2.5") else
               f"{pol}_MAX_1H")
        suf_amg[f"{pol}_SUF_DIARIA"] = base[ref].notna()

    amg = base.copy()
    amg["STATION"] = "AMG"
    for k, v in suf_amg.items():
        amg[k] = v.values.astype(bool)
    for col in dfd.columns:
        if col not in amg.columns:
            amg[col] = np.nan
    amg = amg[dfd.columns]

    return (pd.concat([dfd, amg], ignore_index=True)
            .sort_values(["STATION", "FECHA"]).reset_index(drop=True))


IAS_SOURCE = {
    "PM10": "PM10_AVG_24H", "PM2.5": "PM2.5_AVG_24H",
    "CO": "CO_MAX_8H", "O3": "O3_MAX_1H", "NO2": "NO2_MAX_1H", "SO2": "SO2_MAX_1H",
    "PM10_NOWCAST": "PM10_NOWCAST_MAX", "PM2.5_NOWCAST": "PM2.5_NOWCAST_MAX",
}
ORDEN_DOM = ["PM2.5", "O3", "PM10", "NO2", "SO2", "CO"]


def _frac_rango(valor, pol, cat):
    if valor is None or pd.isna(valor) or cat is None:
        return 0.0
    for lo, hi, c in RANGOS.get(pol, []):
        if c == cat:
            lo_v = float(lo) if lo is not None else -np.inf
            hi_v = float(hi) if hi is not None else np.inf
            if np.isfinite(lo_v) and np.isfinite(hi_v) and (hi_v - lo_v) > 0:
                return (valor - lo_v) / (hi_v - lo_v)
            return 1.0 if np.isinf(hi_v) else 0.0
    return 0.0


def compute_ias_daily(dfd):
    dfd = dfd.copy()
    for pol, vcol in IAS_SOURCE.items():
        if vcol not in dfd.columns:
            continue
        if pol in ("PM10", "PM2.5", "PM10_NOWCAST", "PM2.5_NOWCAST"):
            vals = dfd[vcol].apply(lambda v: np.nan if pd.isna(v) else float(round_half_up(v, 0)))
        elif pol == "CO":
            vals = dfd[vcol].apply(lambda v: np.nan if pd.isna(v) else float(round_half_up(v, 2)))
        else:
            vals = dfd[vcol]
        pol_for_class = pol if pol in RANGOS else pol.replace("_NOWCAST", "")
        resultado = vals.apply(lambda v: pd.Series(clasifica(v, pol_for_class)) if not pd.isna(v) else pd.Series([None, None]))
        dfd[f"IAS_{pol}_CAT_DIA"] = resultado[0]
        dfd[f"IAS_{pol}_SCORE_DIA"] = resultado[1]

    for pol, vcol in {"PM10": "PM10_AVG_24H", "PM2.5": "PM2.5_AVG_24H", "CO": "CO_MAX_8H",
                     "O3": "O3_MAX_1H", "NO2": "NO2_MAX_1H", "SO2": "SO2_MAX_1H"}.items():
        if vcol in dfd.columns and f"IAS_{pol}_VALOR_DIA" not in dfd.columns:
            if pol in ("PM10", "PM2.5"):
                dfd[f"IAS_{pol}_VALOR_DIA"] = dfd[vcol].apply(lambda v: np.nan if pd.isna(v) else float(round_half_up(v, 0)))
            elif pol == "CO":
                dfd[f"IAS_{pol}_VALOR_DIA"] = dfd[vcol].apply(lambda v: np.nan if pd.isna(v) else float(round_half_up(v, 2)))
            else:
                dfd[f"IAS_{pol}_VALOR_DIA"] = dfd[vcol]

    def _dom(row):
        cand = []
        for pol in ORDEN_DOM:
            sc = row.get(f"IAS_{pol}_SCORE_DIA", np.nan)
            if pd.isna(sc):
                continue
            val = row.get(f"IAS_{pol}_VALOR_DIA", np.nan)
            cat = row.get(f"IAS_{pol}_CAT_DIA", None)
            frac = _frac_rango(val, pol, cat) if not pd.isna(val) else 0.0
            cand.append((pol, float(sc), float(frac)))
        if not cand:
            return pd.Series([None, None, None], index=["IAS_GLOBAL_POL_DIA", "IAS_GLOBAL_CAT_DIA", "IAS_GLOBAL_SCORE_DIA"])
        cand.sort(key=lambda t: (t[1], t[2], -ORDEN_DOM.index(t[0])), reverse=True)
        pol, sc, _ = cand[0]
        return pd.Series([pol, row.get(f"IAS_{pol}_CAT_DIA"), sc],
                        index=["IAS_GLOBAL_POL_DIA", "IAS_GLOBAL_CAT_DIA", "IAS_GLOBAL_SCORE_DIA"])

    dfd[["IAS_GLOBAL_POL_DIA", "IAS_GLOBAL_CAT_DIA", "IAS_GLOBAL_SCORE_DIA"]] = dfd.apply(_dom, axis=1)
    return dfd


def annual_compiled_by_station(dfd):
    dfd = dfd.copy()
    dfd["ANIO"] = pd.to_datetime(dfd["FECHA"]).dt.year
    rows = []
    for (est, anio), g in dfd.groupby(["STATION", "ANIO"]):
        dias_anio = 366 if pd.Timestamp(anio, 12, 31).is_leap_year else 365
        cats = g["IAS_GLOBAL_CAT_DIA"]
        cont = cats.value_counts()
        c_buena = int(cont.get("Buena", 0))
        c_acept = int(cont.get("Aceptable", 0))
        c_mala = int(cont.get("Mala", 0))
        c_mm = int(cont.get("Muy mala", 0))
        c_ext = int(cont.get("Extremadamente mala", 0))
        rows.append({
            "STATION": est, "ANIO": anio, "DIAS_ANIO": dias_anio,
            "DIAS_IAS_BUENA": c_buena, "DIAS_IAS_ACEPTABLE": c_acept, "DIAS_IAS_MALA": c_mala,
            "DIAS_IAS_MUY_MALA": c_mm, "DIAS_IAS_EXTREMADAMENTE_MALA": c_ext,
        })
    return pd.DataFrame(rows).sort_values(["STATION", "ANIO"]).reset_index(drop=True)


def dias_buena_aceptable_por_estacion(res_comp_est, anio, estaciones=EST_ORDER_BASE):
    """Días con IAS Buena + Aceptable por estación (sin la virtual AMG), para
    el mapa de burbujas. Una estación es "suficiente" si tiene al menos
    suf_min_yearly(anio) días clasificados (75% del año), igual que el filtro
    exclude_sin_suf_anual de tu notebook."""
    salida = {}
    for est in estaciones:
        fila = res_comp_est[(res_comp_est["STATION"] == est) & (res_comp_est["ANIO"] == anio)]
        if fila.empty:
            salida[est] = {"dias": 0, "validos": 0, "suficiente": False}
            continue
        f = fila.iloc[0]
        validos = int(f["DIAS_IAS_BUENA"] + f["DIAS_IAS_ACEPTABLE"] + f["DIAS_IAS_MALA"]
                      + f["DIAS_IAS_MUY_MALA"] + f["DIAS_IAS_EXTREMADAMENTE_MALA"])
        salida[est] = {
            "dias": int(f["DIAS_IAS_BUENA"] + f["DIAS_IAS_ACEPTABLE"]),
            "validos": validos,
            "suficiente": validos >= suf_min_yearly(anio),
        }
    return salida


def dias_buena_aceptable_por_estacion_contaminante(dfd_all, anio, contaminante, estaciones=EST_ORDER_BASE):
    """Como dias_buena_aceptable_por_estacion, pero con la categoría IAS diaria
    de UN solo contaminante (columna IAS_{contaminante}_CAT_DIA; CO usa el
    máximo diario del promedio móvil de 8 h). "validos" = días con categoría."""
    col = f"IAS_{contaminante}_CAT_DIA"
    dfd = dfd_all.copy()
    dfd = dfd[pd.to_datetime(dfd["FECHA"]).dt.year == anio]
    salida = {}
    for est in estaciones:
        cats = dfd.loc[dfd["STATION"] == est, col].dropna() if col in dfd.columns else pd.Series(dtype=object)
        validos = int(len(cats))
        salida[est] = {
            "dias": int(cats.isin(["Buena", "Aceptable"]).sum()),
            "validos": validos,
            "suficiente": validos >= suf_min_yearly(anio),
        }
    return salida


def suf_min_yearly(anio):
    """75% de suficiencia anual: 274 días en año común, 275 en bisiesto."""
    return 275 if calendar.isleap(anio) else 274


# Límites NOM vigentes para 2024 ("año 3° en la NOM", igual que NOM_PRESETS[2024]
# en tu notebook). "metodo" describe cómo se evalúa cada renglón, siguiendo la
# metodología de la Tabla 9.2 del informe nacional: percentil 99 anual para los
# límites de 24h que se evalúan por percentil, promedio o máximo anual para el
# resto (igual que "tipo" en LIMITES_ANUALES de tu notebook).
FILAS_NOM_2024 = [
    {"id": "PM10_24H", "contaminante": "PM10", "simbolo": "PM₁₀", "periodo": "24 horas", "columna": "PM10_AVG_24H", "metodo": "percentil99", "limite": 60, "limite_txt": "Percentil 99 ≤ 60", "unidad": "µg/m³"},
    {"id": "PM10_ANUAL", "contaminante": "PM10", "simbolo": "PM₁₀", "periodo": "Anual", "columna": "PM10_AVG_24H", "metodo": "promedio", "limite": 28, "limite_txt": "Promedio ≤ 28", "unidad": "µg/m³"},
    {"id": "PM25_24H", "contaminante": "PM2.5", "simbolo": "PM₂.₅", "periodo": "24 horas", "columna": "PM2.5_AVG_24H", "metodo": "percentil99", "limite": 33, "limite_txt": "Percentil 99 ≤ 33", "unidad": "µg/m³"},
    {"id": "PM25_ANUAL", "contaminante": "PM2.5", "simbolo": "PM₂.₅", "periodo": "Anual", "columna": "PM2.5_AVG_24H", "metodo": "promedio", "limite": 10, "limite_txt": "Promedio ≤ 10", "unidad": "µg/m³"},
    {"id": "O3_1H", "contaminante": "O3", "simbolo": "O₃", "periodo": "1 hora", "columna": "O3_MAX_1H", "metodo": "maximo", "limite": 0.090, "limite_txt": "Máximo ≤ 0.090", "unidad": "ppm"},
    {"id": "O3_8H", "contaminante": "O3", "simbolo": "O₃", "periodo": "8 horas", "columna": "O3_MAX_8H", "metodo": "maximo", "limite": 0.060, "limite_txt": "Máximo de 8h ≤ 0.060", "unidad": "ppm"},
    {"id": "CO_1H", "contaminante": "CO", "simbolo": "CO", "periodo": "1 hora", "columna": "CO_MAX_1H", "metodo": "maximo", "limite": 26.0, "limite_txt": "Máximo ≤ 26", "unidad": "ppm"},
    {"id": "CO_8H", "contaminante": "CO", "simbolo": "CO", "periodo": "8 horas", "columna": "CO_MAX_8H", "metodo": "maximo", "limite": 9.0, "limite_txt": "Máximo de 8h ≤ 9", "unidad": "ppm"},
    {"id": "NO2_1H", "contaminante": "NO2", "simbolo": "NO₂", "periodo": "1 hora", "columna": "NO2_MAX_1H", "metodo": "maximo", "limite": 0.106, "limite_txt": "Máximo ≤ 0.106", "unidad": "ppm"},
    {"id": "NO2_ANUAL", "contaminante": "NO2", "simbolo": "NO₂", "periodo": "Anual", "columna": "NO2_AVG_24H", "metodo": "promedio", "limite": 0.021, "limite_txt": "Promedio ≤ 0.021", "unidad": "ppm"},
    {"id": "SO2_1H", "contaminante": "SO2", "simbolo": "SO₂", "periodo": "1 hora", "columna": "SO2_MAX_1H", "metodo": "percentil99", "limite": 0.075, "limite_txt": "Percentil 99 ≤ 0.075", "unidad": "ppm"},
    {"id": "SO2_24H", "contaminante": "SO2", "simbolo": "SO₂", "periodo": "24 horas", "columna": "SO2_AVG_24H", "metodo": "maximo", "limite": 0.040, "limite_txt": "Máximo ≤ 0.040", "unidad": "ppm"},
]


def calcular_cumplimiento_nom(dfd, anio, filas=FILAS_NOM_2024, estaciones=EST_ORDER_BASE):
    """
    Evalúa, por estación y por cada renglón de FILAS_NOM_2024, si se cumple
    el límite NOM correspondiente durante 'anio'. 'dfd' es la tabla diaria
    por estación (build_daily_table(dfh), sin la fila AMG).

    Regresa {fila_id: {estacion: {"status": ..., "valor": float|None}}}.
    status en {"cumple","no_cumple","DI","sin_datos"}. "sin_datos" es cuando
    la estación no tuvo NI UN día válido en todo el año para ese contaminante
    -- puede ser por no tener el equipo o por estar fuera de operación todo
    el año; esa distinción se resuelve del lado del informe con la tabla de
    equipamiento por estación (Tabla 2), no aquí.
    """
    minimo_anual = suf_min_yearly(anio)
    resultado = {}
    for fila in filas:
        resultado[fila["id"]] = {}
        for est in estaciones:
            serie = dfd.loc[dfd["STATION"] == est, fila["columna"]].dropna()
            dias_validos = len(serie)
            if dias_validos == 0:
                resultado[fila["id"]][est] = {"status": "sin_datos", "valor": None, "dias_validos": 0}
                continue
            if dias_validos < minimo_anual:
                resultado[fila["id"]][est] = {"status": "DI", "valor": None, "dias_validos": dias_validos}
                continue

            if fila["metodo"] == "percentil99":
                valor = float(np.percentile(serie, 99))
            elif fila["metodo"] == "promedio":
                valor = float(serie.mean())
            else:  # "maximo"
                valor = float(serie.max())

            status = "cumple" if valor <= fila["limite"] else "no_cumple"
            resultado[fila["id"]][est] = {"status": status, "valor": round(valor, 4), "dias_validos": dias_validos}
    return resultado


def resumen_cumplimiento_nom(dfd, anio, filas=FILAS_NOM_2024):
    """Convierte calcular_cumplimiento_nom() a una lista serializable a JSON:
    una entrada por renglón, con su metadata y los resultados por estación."""
    resultados = calcular_cumplimiento_nom(dfd, anio, filas=filas)
    salida = []
    for fila in filas:
        salida.append({
            "id": fila["id"],
            "contaminante": fila["contaminante"],
            "simbolo": fila["simbolo"],
            "periodo": fila["periodo"],
            "limite_txt": fila["limite_txt"],
            "unidad": fila["unidad"],
            "estaciones": resultados[fila["id"]],
        })
    return salida


# ---------------------------------------------------------------------------
# Orquestación: de un Excel de BD_{anio} a los campos que necesita el informe
# ---------------------------------------------------------------------------
def calcular_resumen_anual(ruta_excel, estacion="AMG"):
    """Corre el pipeline completo (Secciones 1-3) sobre un BD_{anio}.xlsx y
    regresa el diccionario de cifras que usa generar_informe.py para ese año."""
    dfh, anio = load_and_prepare_db(ruta_excel)
    dfh = calcular_columnas_rolling(dfh)

    dfd = build_daily_table(dfh)
    dfd_all = rebuild_amg_from_daily(dfd)
    dfd_all = compute_ias_daily(dfd_all)

    res_comp_est = annual_compiled_by_station(dfd_all)
    fila = res_comp_est.query("STATION == @estacion and ANIO == @anio")
    if fila.empty:
        raise ValueError(f"No se generaron resultados para la estación '{estacion}' en {anio}.")
    fila = fila.iloc[0]

    dias_buena_aceptable = int(fila["DIAS_IAS_BUENA"] + fila["DIAS_IAS_ACEPTABLE"])
    dias_estaciones = dias_buena_aceptable_por_estacion(res_comp_est, anio)
    dias_estaciones_contaminante = {
        pol: dias_buena_aceptable_por_estacion_contaminante(dfd_all, anio, pol) for pol in ("CO", "NO2", "SO2", "O3", "PM10", "PM2.5")
    }

    dfh = calcular_columnas_horarias_gases(dfh)
    dfh = calcular_ias_global_horario(dfh)
    horas_categoria = resumen_horas_categoria(agregar_amg_horario(dfh), estaciones=EST_ORDER_BASE + ["AMG"])
    violines_mensuales = {
        "CO": calcular_violines_mensuales(dfh, "CO", suavizado=rolling_8h),
        "NO2": calcular_violines_mensuales(dfh, "NO2"),
        "O3": calcular_violines_mensuales(dfh, "O3", suavizado=rolling_8h),
        # PM10: NowCast (promedio ponderado de 12 h) calculado sobre la serie del AMG
        "PM10": calcular_violines_mensuales(
            dfh, "PM10",
            suavizado=lambda s: pd.to_numeric(serie_nowcast_por_estacion(s.to_frame("PM10"), "PM10", 0), errors="coerce"),
        ),
        "PM2.5": calcular_violines_mensuales(
            dfh, "PM2.5",
            suavizado=lambda s: pd.to_numeric(serie_nowcast_por_estacion(s.to_frame("PM2.5"), "PM2.5", 1), errors="coerce"),
        ),
    }
    perfil_horario = {pol: calcular_perfil_horario(dfh, pol) for pol in ["PM10", "PM2.5", "O3", "NO2", "SO2", "CO"]}

    cumplimiento_nom = resumen_cumplimiento_nom(dfd, anio)

    return {
        "anio": anio,
        "dias_buena_aceptable": dias_buena_aceptable,
        "horas_categoria_calidad": horas_categoria,
        "cumplimiento_nom": cumplimiento_nom,
        "perfil_horario": perfil_horario,
        "violines_mensuales": violines_mensuales,
        "dias_buena_aceptable_estaciones": dias_estaciones,
        "dias_buena_aceptable_estaciones_contaminante": dias_estaciones_contaminante,
    }


def actualizar_historico(anio, url_o_id=None, ruta_excel=None, ruta_resumen=RUTA_RESUMEN_HISTORICO):
    """Descarga (si hace falta) BD_{anio}, calcula el resumen, y lo agrega
    a resumen_historico.json sin borrar los demás años."""
    if ruta_excel is None:
        ruta_excel = descargar_bd(anio, url_o_id=url_o_id)

    resultado = calcular_resumen_anual(ruta_excel)
    if resultado["anio"] != anio:
        print(f"Aviso: se pidió el año {anio} pero la base de datos corresponde a {resultado['anio']}.")

    ruta_resumen = Path(ruta_resumen)
    historico = json.loads(ruta_resumen.read_text(encoding="utf-8")) if ruta_resumen.exists() else {}
    entrada = historico.get(str(anio), {})
    entrada["dias_buena_aceptable"] = resultado["dias_buena_aceptable"]
    entrada["horas_categoria_calidad"] = resultado["horas_categoria_calidad"]
    entrada["cumplimiento_nom"] = resultado["cumplimiento_nom"]
    entrada["perfil_horario"] = resultado["perfil_horario"]
    entrada["violines_mensuales"] = resultado["violines_mensuales"]
    entrada["dias_buena_aceptable_estaciones"] = resultado["dias_buena_aceptable_estaciones"]
    entrada["dias_buena_aceptable_estaciones_contaminante"] = resultado["dias_buena_aceptable_estaciones_contaminante"]
    historico[str(anio)] = entrada

    ruta_resumen.parent.mkdir(parents=True, exist_ok=True)
    ruta_resumen.write_text(json.dumps(historico, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"[{anio}] {resultado['dias_buena_aceptable']} días Buena/Aceptable (AMG) -> {ruta_resumen}")
    return historico


def _main():
    parser = argparse.ArgumentParser(description="Calcula los datos anuales del informe desde BD_{anio}.")
    parser.add_argument("--anio", type=int, required=True)
    parser.add_argument("--sheet", type=str, default=None, help="URL o ID de Google Sheets (opcional)")
    parser.add_argument("--excel", type=str, default=None, help="Ruta a un .xlsx local ya descargado (opcional)")
    args = parser.parse_args()

    actualizar_historico(args.anio, url_o_id=args.sheet, ruta_excel=args.excel)


if __name__ == "__main__":
    _main()
