"""Genera instancias estáticas de Montserrat (Regular/SemiBold/Bold) a partir
de la variable font instalada en Windows, para poder incrustarlas en el PDF
con reportlab (que no soporta variable fonts)."""
from pathlib import Path

from fontTools import ttLib
from fontTools.varLib.instancer import instantiateVariableFont

SRC = Path(r"C:\Users\Mariana.Castellanos\AppData\Local\Microsoft\Windows\Fonts\Montserrat-VariableFont_wght.ttf")
OUT_DIR = Path(__file__).resolve().parent

WEIGHTS = {
    "Montserrat-Regular.ttf": ("Montserrat", 400),
    "Montserrat-SemiBold.ttf": ("Montserrat SemiBold", 600),
    "Montserrat-Bold.ttf": ("Montserrat Bold", 700),
}

for filename, (family_name, wght) in WEIGHTS.items():
    font = ttLib.TTFont(str(SRC))
    instantiateVariableFont(font, {"wght": wght}, inplace=True, updateFontNames=True)

    name_table = font["name"]
    for name_id in (1, 3, 4, 6, 16):
        name_table.setName(family_name.replace(" ", "") if name_id == 6 else family_name, name_id, 3, 1, 0x409)
        name_table.setName(family_name.replace(" ", "") if name_id == 6 else family_name, name_id, 1, 0, 0)
    name_table.setName("Regular", 2, 3, 1, 0x409)
    name_table.setName("Regular", 2, 1, 0, 0)

    out_path = OUT_DIR / filename
    font.save(str(out_path))
    print("generado:", out_path)
