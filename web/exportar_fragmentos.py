# -*- coding: utf-8 -*-
"""Exporta las plantillas de Excel a fragmentos HTML + imagenes para la app web.

Se ejecuta UNA vez en Windows (necesita Excel). Los archivos resultantes son
estaticos, asi la app Flask corre en cualquier lado (incluido Linux/Render).
"""
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from excel_a_html import convertir            # noqa: E402
from build_web import IDS_COTI, IDS_SUS       # noqa: E402

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEST = os.path.join(BASE, "webapp", "static", "doc")


def main():
    os.makedirs(DEST, exist_ok=True)
    print("Leyendo plantillas de Excel...")
    coti = convertir(os.path.join(BASE, "plantilla", "plantilla_cotizacion.xlsx"),
                     fila_fin=36, col_fin=13, ids=IDS_COTI, clase="xl")
    sus = convertir(os.path.join(BASE, "plantilla", "plantilla_sustento.xlsx"),
                    fila_fin=49, col_fin=14, ids=IDS_SUS, clase="xl")

    for nombre, html in (("cotizacion.html", coti), ("sustento.html", sus)):
        with open(os.path.join(DEST, nombre), "w", encoding="utf-8") as f:
            f.write(html)
        print("   ->", nombre, "(%.0f KB)" % (len(html) / 1024))

    for img in ("logo.png", "firma.png"):
        shutil.copyfile(os.path.join(BASE, "assets", img), os.path.join(DEST, img))
    print("   -> logo.png, firma.png")
    print("OK. Fragmentos en", DEST)


if __name__ == "__main__":
    main()
