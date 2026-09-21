#@title Exportar resumen para el Informe Anual (PDF)
# ================================================================
# Pega esta celda al FINAL de tu notebook de diagnóstico, después de que
# `res_comp_est` ya esté calculado (Sección 3, salida de
# annual_compiled_by_station(dfd_all)).
#
# Qué hace:
#   - Toma de res_comp_est los días IAS "Buena" + "Aceptable" de la
#     estación AMG para el año que se acaba de procesar (`anio`).
#   - Actualiza (sin borrar años anteriores) un único archivo
#     resumen_historico.json compartido entre todos los años.
#   - generar_informe.py lee ese archivo para armar la Figura 2 y su
#     párrafo de resultados automáticamente.
#
# Cada vez que corras este notebook para un año distinto (cambiando
# anio_to_load y BD_{anio}), vuelve a correr esta celda al final: así
# vas construyendo el histórico de 2019 a 2024 poco a poco.
# ================================================================
import json
from pathlib import Path


def exportar_resumen_informe(anio: int, res_comp_est: pd.DataFrame, ruta_resumen: str, estacion: str = "AMG") -> dict:
    """
    Actualiza resumen_historico.json con las cifras del año `anio`.
    No sobreescribe los demás años ya guardados en el archivo.
    """
    fila = res_comp_est.query("STATION == @estacion and ANIO == @anio")
    if fila.empty:
        raise ValueError(
            f"No hay resultados de la estación '{estacion}' para el año {anio} en res_comp_est. "
            "Revisa que el pipeline de la Sección 3 ya se haya corrido para este año."
        )
    fila = fila.iloc[0]

    dias_buena_aceptable = int(fila["DIAS_IAS_BUENA"] + fila["DIAS_IAS_ACEPTABLE"])

    ruta = Path(ruta_resumen)
    if ruta.exists():
        historico = json.loads(ruta.read_text(encoding="utf-8"))
    else:
        historico = {}

    entrada = historico.get(str(anio), {})
    entrada.update({
        "dias_buena_aceptable": dias_buena_aceptable,
    })
    historico[str(anio)] = entrada

    ruta.parent.mkdir(parents=True, exist_ok=True)
    ruta.write_text(json.dumps(historico, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"[{anio}] {estacion}: {dias_buena_aceptable} días Buena/Aceptable -> guardado en {ruta}")
    return historico


# Ruta compartida en Drive: un nivel arriba de la carpeta de figuras de este
# año (PATH ya trae "/{anio_to_load}" al final), para que las entradas de
# todos los años queden juntas en un solo archivo.
RUTA_RESUMEN_INFORME = str(Path(PATH).parent / "resumen_historico.json")

historico_actualizado = exportar_resumen_informe(anio, res_comp_est, RUTA_RESUMEN_INFORME)
print(historico_actualizado)
