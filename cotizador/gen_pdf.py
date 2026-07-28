"""Genera los PDF de la Cotizacion FMI y del Sustento con reportlab."""
from __future__ import annotations
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, Paragraph,
                                Spacer, Image)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

from .modelo import Cotizacion
from .config import ruta_logo

AZUL = colors.HexColor("#1F3864")
AZUL_CLARO = colors.HexColor("#D9E1F2")
GRIS = colors.HexColor("#808080")
NARANJA = colors.HexColor("#C55A11")

_styles = getSampleStyleSheet()

def _p(txt, size=8, bold=False, align=TA_LEFT, color=colors.black):
    st = ParagraphStyle("x", parent=_styles["Normal"], fontName="Helvetica-Bold" if bold else "Helvetica",
                        fontSize=size, leading=size + 2, alignment=align, textColor=color)
    return Paragraph(str(txt).replace("\n", "<br/>"), st)


def _mon(v, moneda="S/."):
    try:
        return f"{moneda} {float(v):,.2f}"
    except (TypeError, ValueError):
        return str(v)


def _logo(cfg, max_w_mm=38, max_h_mm=15):
    """Devuelve un flowable Image del logo (escalado) o un texto de respaldo."""
    ruta = ruta_logo(cfg)
    if ruta:
        try:
            from PIL import Image as PILImage
            with PILImage.open(ruta) as im:
                w, h = im.size
            esc = min((max_w_mm * mm) / w, (max_h_mm * mm) / h)
            img = Image(ruta, width=w * esc, height=h * esc)
            img.hAlign = "CENTER"
            return img
        except Exception:
            pass
    return _p("PANORAMA\nOutsourcing", 13, True, TA_CENTER, NARANJA)


def _bloque_firma(cfg, emp):
    """Bloque de firma: imagen de firma (si hay) sobre el nombre y cargo."""
    from .config import ruta_firma
    filas = []
    ruta = ruta_firma(cfg)
    if ruta:
        try:
            from PIL import Image as PILImage
            with PILImage.open(ruta) as im:
                w, h = im.size
            esc = min((45 * mm) / w, (18 * mm) / h)
            img = Image(ruta, width=w * esc, height=h * esc)
            img.hAlign = "CENTER"
            filas.append([img])
        except Exception:
            pass
    filas.append([_p("_____________________________<br/><b>" + emp["firma_nombre"] +
                     "</b><br/>" + emp["firma_cargo"], 8, False, TA_CENTER)])
    t = Table(filas, colWidths=[80 * mm])
    t.hAlign = "CENTER"
    t.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"),
                           ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                           ("TOPPADDING", (0, 0), (-1, -1), 0)]))
    return t


def generar_cotizacion_pdf(cot: Cotizacion, cfg: dict, ruta: str) -> str:
    emp = cfg["empresa"]; cond = cfg["condiciones_comerciales"]; notas = cfg.get("notas", [])
    m = cot.moneda
    doc = SimpleDocTemplate(ruta, pagesize=A4, leftMargin=12 * mm, rightMargin=12 * mm,
                            topMargin=10 * mm, bottomMargin=10 * mm,
                            title=f"Cotizacion {cot.numero}")
    elems = []

    # Encabezado
    cab = Table([[_logo(cfg),
                  _p("COTIZACION DE SERVICIOS", 13, True, TA_CENTER),
                  Table([[_p("Codigo:", 7, True), _p(emp["codigo_cotizacion"], 7)],
                         [_p("Version:", 7, True), _p(emp["version_cotizacion"], 7)],
                         [_p("Pagina:", 7, True), _p("1 de 1", 7)]],
                        colWidths=[16 * mm, 22 * mm],
                        style=[("GRID", (0, 0), (-1, -1), 0.5, GRIS),
                               ("VALIGN", (0, 0), (-1, -1), "MIDDLE")])]],
                 colWidths=[40 * mm, 106 * mm, 40 * mm])
    cab.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.7, colors.black),
                             ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                             ("INNERGRID", (0, 0), (1, 0), 0.5, colors.black)]))
    elems += [cab, Spacer(1, 4)]

    # Datos generales
    datos = Table([[_p("Fecha:", 8, True), _p(cot.fecha, 8)],
                   [_p("Cliente:", 8, True), _p(cot.cliente, 8)],
                   [_p("Contacto:", 8, True), _p(cot.contacto, 8)],
                   [_p("Asunto :", 8, True), _p(cot.asunto, 8)]],
                  colWidths=[24 * mm, 162 * mm])
    datos.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    elems += [datos, Spacer(1, 3)]

    ref = Table([[_p("Referencia", 8, True, TA_LEFT, colors.white),
                  _p("Presupuesto Economico N°", 8, True, TA_CENTER, colors.white)],
                 [_p("N° Ticket / Solicitud: " + cot.referencia_ticket, 8),
                  _p(cot.numero, 9, True, TA_CENTER)]],
                colWidths=[110 * mm, 76 * mm])
    ref.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), GRIS),
                             ("GRID", (0, 0), (-1, -1), 0.5, GRIS),
                             ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    elems += [ref, Spacer(1, 6)]

    # Tabla de items
    data = [[_p("ITEM", 8, True, TA_CENTER, colors.white),
             _p("D E S C R I P C I O N", 8, True, TA_CENTER, colors.white),
             _p("U.M.", 8, True, TA_CENTER, colors.white),
             _p("CANT.", 8, True, TA_CENTER, colors.white),
             _p("PRECIO", 8, True, TA_CENTER, colors.white),
             _p("PRECIO PARCIAL", 8, True, TA_CENTER, colors.white),
             _p("SUB TOTAL", 8, True, TA_CENTER, colors.white)]]
    subtotal = cot.subtotal()
    data.append([_p("01.00", 8, True, TA_CENTER, AZUL),
                 _p((cot.asunto or "").upper(), 8, True, TA_LEFT, AZUL),
                 "", "", "", "", _p(_mon(subtotal, m), 8, True, TA_RIGHT)])
    for i, ln in enumerate(cot.lineas, 1):
        vu = ln.venta_unit(cot.margen, cot.modo_margen)
        vp = ln.venta_parcial(cot.margen, cot.modo_margen)
        data.append([_p(f"01.{i:02d}", 8, False, TA_CENTER),
                     _p(ln.descripcion, 8),
                     _p(ln.unidad, 8, False, TA_CENTER),
                     _p(_num(ln.cantidad), 8, False, TA_CENTER),
                     _p(_mon(vu, m), 8, False, TA_RIGHT),
                     _p(_mon(vp, m), 8, False, TA_RIGHT),
                     _p(_mon(vp, m), 8, False, TA_RIGHT)])
    tabla = Table(data, colWidths=[13 * mm, 83 * mm, 12 * mm, 12 * mm, 22 * mm, 22 * mm, 22 * mm])
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), AZUL),
        ("BACKGROUND", (0, 1), (-1, 1), AZUL_CLARO),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    elems += [tabla]

    # Totales
    tot = Table([[_p("COSTO DIRECTO", 9, True, TA_LEFT), _p(_mon(subtotal, m), 9, True, TA_RIGHT)],
                 [_p(f"IGV {int(cot.igv*100)}%", 9, True, TA_LEFT), _p(_mon(cot.igv_monto(), m), 9, True, TA_RIGHT)],
                 [_p("PRECIO CON IGV", 9, True, TA_LEFT, colors.white), _p(_mon(cot.total(), m), 9, True, TA_RIGHT, colors.white)]],
                colWidths=[142 * mm, 44 * mm])
    tot.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                             ("BACKGROUND", (0, 2), (-1, 2), AZUL),
                             ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    elems += [tot, Spacer(1, 8)]

    # Condiciones
    elems.append(_p("Condiciones Comerciales:", 9, True))
    condiciones = [("Validez de la Propuesta Economica:", cond["validez"]),
                   ("Forma de pago:", cond["forma_pago"]),
                   ("Plazo de Ejecucion y programacion:", cond["plazo_ejecucion"]),
                   ("Tiempo de ejecucion:", cond["tiempo_ejecucion"]),
                   ("Lugar de Entrega:", cond["lugar_entrega"]),
                   ("Orden a :", cond["orden_a"])]
    cc = Table([[_p(k, 8, True), _p(v, 8)] for k, v in condiciones], colWidths=[62 * mm, 124 * mm])
    cc.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    elems += [cc, Spacer(1, 6), _p("Notas:", 9, True)]
    for nota in notas:
        elems.append(_p("• " + nota, 8))

    elems += [Spacer(1, 22), _bloque_firma(cfg, emp)]

    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    doc.build(elems)
    return ruta


def generar_sustento_pdf(cot: Cotizacion, cfg: dict, ruta: str) -> str:
    emp = cfg["empresa"]; m = cot.moneda
    doc = SimpleDocTemplate(ruta, pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm,
                            topMargin=12 * mm, bottomMargin=12 * mm,
                            title=f"Sustento {cot.numero}")
    e = []
    caja = Table([[_p("Codigo:", 7, True), _p(emp["codigo_sustento"], 7)],
                  [_p("Version:", 7, True), _p(emp["version_sustento"], 7)],
                  [_p("Pagina:", 7, True), _p("1 de 1", 7)]],
                 colWidths=[16 * mm, 22 * mm],
                 style=[("GRID", (0, 0), (-1, -1), 0.5, GRIS),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE")])
    cab = Table([[_logo(cfg), _p("SUSTENTO DE INGRESOS Y GASTOS", 14, True, TA_CENTER), caja]],
                colWidths=[42 * mm, 100 * mm, 38 * mm])
    cab.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.7, colors.black),
                             ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    e += [cab, Spacer(1, 10)]
    e += [_p(f"<b>Fecha:</b> {cot.fecha}", 10),
          _p(f"<b>Presupuesto Economico de referencia:</b> {cot.numero}", 10), Spacer(1, 10)]

    def bloque(titulo, color, filas):
        t = Table([[_p(titulo, 11, True, TA_CENTER, colors.white)]], colWidths=[180 * mm])
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), color)]))
        e.append(t)
        body = Table(filas, colWidths=[90 * mm, 90 * mm])
        body.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.4, colors.lightgrey),
                                   ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
        e.append(body); e.append(Spacer(1, 8))

    prov = cot.proveedor_principal
    gasto = cot.costo_gasto()
    bloque("GASTO", NARANJA, [
        [_p("Proveedor :", 9, True), _p(prov.nombre or "(proveedor)", 9)],
        [_p("Motivo :", 9, True), _p(cot.asunto, 9)],
        [_p("Monto", 9), _p(_mon(gasto, m), 9, False, TA_RIGHT)],
        [_p("IGV", 9), _p(_mon(round(gasto * cot.igv, 2), m), 9, False, TA_RIGHT)],
        [_p("Total c/IGV", 9, True), _p(_mon(cot.gasto_con_igv(), m), 9, True, TA_RIGHT)]])

    venta = cot.subtotal()
    bloque("VENTA", AZUL, [
        [_p("Cliente :", 9, True), _p(cot.cliente, 9)],
        [_p("Local / Sede :", 9, True), _p(cot.sede or cot.asunto, 9)],
        [_p("Monto", 9), _p(_mon(venta, m), 9, False, TA_RIGHT)],
        [_p("IGV", 9), _p(_mon(cot.igv_monto(), m), 9, False, TA_RIGHT)],
        [_p("Monto c/IGV", 9, True), _p(_mon(cot.total(), m), 9, True, TA_RIGHT)]])

    bloque("RESUMEN", GRIS, [
        [_p("", 9), _p("s/IGV      c/IGV", 9, True, TA_RIGHT)],
        [_p("Total Venta", 9, True), _p(f"{_mon(venta, m)}      {_mon(cot.total(), m)}", 9, False, TA_RIGHT)],
        [_p("Total Gasto", 9, True), _p(f"{_mon(gasto, m)}      {_mon(cot.gasto_con_igv(), m)}", 9, False, TA_RIGHT)],
        [_p("Utilidad Obtenida", 9, True), _p(f"{_mon(cot.utilidad(), m)}      {_mon(round(cot.utilidad()*(1+cot.igv),2), m)}", 9, False, TA_RIGHT)],
        [_p("Utilidad Obtenida (%)", 9, True), _p(f"{cot.utilidad_pct()*100:.2f}%", 9, True, TA_RIGHT)]])

    e += [Spacer(1, 20), _bloque_firma(cfg, emp)]
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    doc.build(e)
    return ruta


def _num(x):
    x = float(x or 0)
    return int(x) if x == int(x) else round(x, 2)
