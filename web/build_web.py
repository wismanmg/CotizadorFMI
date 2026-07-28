# -*- coding: utf-8 -*-
"""Genera 'Cotizador FMI (web).html': la app completa en un solo archivo,
con los documentos calcados de las plantillas de Excel.
"""
import base64
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from excel_a_html import convertir           # noqa: E402

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WEB = os.path.join(BASE, "web")
OUT = os.path.join(BASE, "dist_web", "Cotizador FMI (web).html")

# celdas que llena el JS  ->  id en el HTML
IDS_COTI = {
    "D4": "c_fecha", "D6": "c_cli", "D7": "c_con", "D8": "c_asu",
    "D11": "c_tick", "K11": "c_num",
    "B14": "c_grupo", "M14": "c_gsub",
    "M17": "c_cd", "L18": "c_pigv", "M18": "c_igv", "M19": "c_tot",
    "A1": "c_logo", "K31": "c_firma",
}
IDS_SUS = {
    "E5": "s_fecha", "E7": "s_num", "G12": "s_mes", "G13": "s_mot", "G15": "s_prov",
    "G17": "s_g1", "G18": "s_g2", "G19": "s_g3",
    "G24": "s_cli", "G25": "s_sede",
    "G27": "s_v1", "G28": "s_v2", "G29": "s_v3",
    "G35": "r_v1", "J35": "r_v2", "G36": "r_g1", "J36": "r_g2",
    "G37": "r_u1", "J37": "r_u2", "G39": "r_pct",
    "A1": "s_logo", "J44": "s_firma",     # J44:N47 = recuadro de firma
}


def datauri(p):
    with open(p, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode()


def main():
    print("Leyendo plantillas de Excel...")
    coti = convertir(os.path.join(BASE, "plantilla", "plantilla_cotizacion.xlsx"),
                     fila_fin=36, col_fin=13, ids=IDS_COTI, clase="xl")
    sus = convertir(os.path.join(BASE, "plantilla", "plantilla_sustento.xlsx"),
                    fila_fin=49, col_fin=14, ids=IDS_SUS, clase="xl")

    shell = open(os.path.join(WEB, "shell.html"), encoding="utf-8").read()
    html = (shell
            .replace("<!--COTI-->", coti)
            .replace("<!--SUS-->", sus)
            .replace("__LOGO__", datauri(os.path.join(BASE, "assets", "logo.png")))
            .replace("__FIRMA__", datauri(os.path.join(BASE, "assets", "firma.png"))))

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(html)
    print("OK ->", OUT, "(%.0f KB)" % (os.path.getsize(OUT) / 1024))


if __name__ == "__main__":
    main()
