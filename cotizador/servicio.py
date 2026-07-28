"""Orquesta la generacion de todos los documentos de una cotizacion."""
from __future__ import annotations
import os
import re

from .modelo import Cotizacion, quitar_acentos
from .gen_cotizacion import generar_cotizacion_excel
from .gen_sustento import generar_sustento_excel
from .gen_comparativo import generar_comparativo_excel
from .gen_pdf import generar_cotizacion_pdf, generar_sustento_pdf
from .gen_com import generar_con_excel, excel_disponible
from .control import registrar_en_bitacora
from . import config as cfgmod


def _slug(numero: str) -> str:
    s = quitar_acentos(numero).replace(" ", "").replace("/", "-")
    return re.sub(r"[^\w\-]", "", s)


def generar_todo(cot: Cotizacion, cfg: dict, opciones: dict, carpeta: str | None = None) -> dict:
    """Genera los documentos marcados en 'opciones'. Devuelve {tipo: ruta}."""
    carpeta = carpeta or cfgmod.carpeta_salidas(cfg)
    slug = _slug(cot.numero) or "cotizacion"
    sub = os.path.join(carpeta, slug)
    os.makedirs(sub, exist_ok=True)
    resultados = {}

    q_coti_x = opciones.get("cotizacion_excel", True)
    q_coti_p = opciones.get("cotizacion_pdf", True)
    q_sus = opciones.get("sustento", True)

    coti_x = os.path.join(sub, f"Cotizacion {cot.numero}.xlsx")
    coti_p = os.path.join(sub, f"Cotizacion {cot.numero}.pdf")
    sus_x = os.path.join(sub, f"Sustento de Gasto {cot.numero}.xlsx")
    sus_p = os.path.join(sub, f"Sustento de Gasto {cot.numero}.pdf")
    hecho = {"cx": False, "cp": False, "sx": False, "sp": False}

    # ---- Principal: Microsoft Excel (formato 100% identico + PDF calcado) ----
    if excel_disponible() and (q_coti_x or q_coti_p or q_sus):
        r = generar_con_excel(
            cot,
            xlsx_coti=coti_x if q_coti_x else None,
            pdf_coti=coti_p if q_coti_p else None,
            xlsx_sus=sus_x if q_sus else None,
            pdf_sus=sus_p if q_sus else None,
            cfg=cfg,
        )
        if r:
            if r.get("coti_xlsx"): resultados["Cotizacion (Excel)"] = r["coti_xlsx"]; hecho["cx"] = True
            if r.get("coti_pdf"): resultados["Cotizacion (PDF)"] = r["coti_pdf"]; hecho["cp"] = True
            if r.get("sus_xlsx"): resultados["Sustento (Excel)"] = r["sus_xlsx"]; hecho["sx"] = True
            if r.get("sus_pdf"): resultados["Sustento (PDF)"] = r["sus_pdf"]; hecho["sp"] = True

    # ---- Respaldo (sin Excel o si algo fallo): openpyxl + reportlab ----
    if q_coti_x and not hecho["cx"]:
        resultados["Cotizacion (Excel)"] = generar_cotizacion_excel(cot, cfg, coti_x)
    if q_coti_p and not hecho["cp"]:
        resultados["Cotizacion (PDF)"] = generar_cotizacion_pdf(cot, cfg, coti_p)
    if q_sus and not hecho["sx"]:
        resultados["Sustento (Excel)"] = generar_sustento_excel(cot, cfg, sus_x)
    if q_sus and not hecho["sp"]:
        resultados["Sustento (PDF)"] = generar_sustento_pdf(cot, cfg, sus_p)
    if opciones.get("comparativo", False) and cot.proveedores_comparados:
        resultados["Comparativo"] = generar_comparativo_excel(
            cot.asunto, cot.proveedores_comparados, cfg,
            os.path.join(sub, f"Comparativo {cot.numero}.xlsx"), cot.moneda)
    if opciones.get("control", True):
        bit = os.path.join(cfgmod.DATOS_DIR, "Bitacora_Cotizaciones_FMI.xlsx")
        resultados["Bitacora"] = registrar_en_bitacora(cot, bit)

    return resultados
