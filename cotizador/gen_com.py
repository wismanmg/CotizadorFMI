"""Generacion de Cotizacion y Sustento con Microsoft Excel (COM / pywin32).

Abre las plantillas .xlsx originales (validas, con logo y firma) DENTRO de Excel,
escribe solo los datos variables y exporta. Resultado: XLSX 100% identico y valido,
formulas recalculadas por Excel, y PDF calcado del Excel.

Requiere Microsoft Excel instalado. Si no esta disponible devuelve None y el
llamador usa los generadores openpyxl/reportlab de respaldo.
"""
from __future__ import annotations
import os
import sys

from .modelo import Cotizacion


def _base_dir() -> str:
    return getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def plantilla_coti() -> str:
    return os.path.join(_base_dir(), "plantilla", "plantilla_cotizacion.xlsx")


def plantilla_sustento() -> str:
    return os.path.join(_base_dir(), "plantilla", "plantilla_sustento.xlsx")


def excel_disponible() -> bool:
    try:
        import win32com.client  # noqa
    except Exception:
        return False
    return os.path.exists(plantilla_coti())


# ------------------------------------------------------------------ helpers
def _venta_unit(costo, m, modo):
    costo = float(costo or 0)
    if modo == "sobre_venta":
        d = (1 - m) or 1.0
        return costo / d
    return costo * (1 + m)


def _valor_fecha(cot: Cotizacion):
    """Excel/pywin32 solo acepta datetime (no date). Devuelve datetime o el texto."""
    import datetime as _dt
    f = getattr(cot, "fecha_dt", None)
    if isinstance(f, _dt.datetime):
        return f
    if isinstance(f, _dt.date):
        return _dt.datetime(f.year, f.month, f.day)
    return cot.fecha


def _factor_txt(m, modo):
    if modo == "sobre_venta":
        return True, round(1 - m, 6) or 1.0
    return False, round(1 + m, 6)


_PICTURE = 13  # msoPicture


def _asegurar_merge(ws, fila: int):
    """Garantiza que la descripcion ocupe B:H combinadas en esa fila."""
    try:
        celda = ws.Range(f"B{fila}:H{fila}")
        if not celda.MergeCells:
            celda.Merge()
    except Exception:
        pass


# La celda combinada B:H mide ~84 caracteres a 8pt; Excel no autoajusta
# el alto de celdas combinadas, asi que lo estimamos.
_CHARS_POR_LINEA = 84
_ALTO_LINEA = 11.5
_ALTO_MINIMO = 56.45


def _ajustar_alto(ws, fila: int, texto: str):
    """Da a la fila el alto necesario para que se vea toda la descripcion."""
    try:
        largo = len(texto or "")
        lineas = max(1, -(-largo // _CHARS_POR_LINEA))          # division hacia arriba
        lineas += (texto or "").count("\n")
        alto = max(_ALTO_MINIMO, lineas * _ALTO_LINEA + 10)
        ws.Rows(fila).RowHeight = alto
    except Exception:
        pass


def _asegurar_carpeta(ruta: str | None):
    """Excel no crea carpetas al guardar: hay que crearlas antes."""
    if ruta:
        carpeta = os.path.dirname(os.path.abspath(ruta))
        if carpeta:
            os.makedirs(carpeta, exist_ok=True)


def _reemplazar_shape(ws, shape, ruta_img: str):
    """Sustituye una imagen conservando su posicion y encajandola en su tamaño."""
    left, top, ancho, alto = shape.Left, shape.Top, shape.Width, shape.Height
    shape.Delete()
    nueva = ws.Shapes.AddPicture(os.path.abspath(ruta_img), False, True,
                                 left, top, -1, -1)   # -1 = tamaño original
    w0, h0 = nueva.Width, nueva.Height
    if w0 > 0 and h0 > 0:
        esc = min(ancho / w0, alto / h0)
        # sin esto Excel reajusta el alto al fijar el ancho y el escalado se aplica dos veces
        try:
            nueva.LockAspectRatio = 0     # msoFalse
        except Exception:
            pass
        nueva.Width = w0 * esc
        nueva.Height = h0 * esc
    # centra dentro del hueco que ocupaba la original
    nueva.Left = left + (ancho - nueva.Width) / 2
    nueva.Top = top + (alto - nueva.Height) / 2
    return nueva


def _aplicar_imagenes(ws, cfg: dict):
    """Reemplaza logo de empresa y/o firma solo si el usuario cargo los suyos.

    Heuristica sobre la plantilla:
      - imagenes en las primeras filas  -> logo de empresa (cabecera)
      - en la zona de firma (filas bajas): la mas ancha es la FIRMA, las demas
        son el logo pequeño del recuadro.
    """
    from .config import ruta_logo, ruta_firma, es_personalizado
    logo = ruta_logo(cfg) if es_personalizado(cfg, "logo") else None
    firma = ruta_firma(cfg) if es_personalizado(cfg, "firma") else None
    if not logo and not firma:
        return

    cabecera, zona_firma = [], []
    for i in range(1, ws.Shapes.Count + 1):
        s = ws.Shapes(i)
        if s.Type != _PICTURE:
            continue
        (cabecera if s.TopLeftCell.Row <= 5 else zona_firma).append(s)

    if logo:
        for s in cabecera:
            _reemplazar_shape(ws, s, logo)
    if zona_firma:
        zona_firma.sort(key=lambda s: s.Width, reverse=True)
        if firma:
            _reemplazar_shape(ws, zona_firma[0], firma)      # la mas ancha = firma
        if logo:
            for s in zona_firma[1:]:                          # logos pequeños
                _reemplazar_shape(ws, s, logo)


def _ajustar_pagina(ws, print_area: str):
    """Fija el area de impresion y fuerza que todo entre en UNA sola pagina."""
    ps = ws.PageSetup
    ps.PrintArea = print_area
    ps.Orientation = 1            # xlPortrait
    ps.Zoom = False               # obligatorio para que FitToPages tenga efecto
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1


class _Excel:
    """Context manager que abre Excel via COM y garantiza el cierre."""
    def __enter__(self):
        import pythoncom
        import win32com.client as win32
        self._pyc = pythoncom
        pythoncom.CoInitialize()
        self.app = win32.DispatchEx("Excel.Application")
        self.app.Visible = False
        self.app.DisplayAlerts = False
        return self.app

    def __exit__(self, *exc):
        try:
            self.app.Quit()
        except Exception:
            pass
        try:
            self._pyc.CoUninitialize()
        except Exception:
            pass
        return False


# ------------------------------------------------------------------ cotizacion
def _llenar_cotizacion(app, cot: Cotizacion, xlsx_out: str, pdf_out: str | None, cfg: dict | None = None):
    wb = app.Workbooks.Open(os.path.abspath(plantilla_coti()))
    try:
        ws = wb.Worksheets(1)
        lineas = [l for l in cot.lineas if (l.descripcion or l.costo_unit)] or cot.lineas[:1]
        n = len(lineas)
        base = 2
        if n > base:
            # Copiar+Insertar la fila modelo replica el merge B:H, bordes y ajuste de texto.
            # (Insert() a secas copia el formato pero NO las celdas combinadas.)
            extra = n - base
            ws.Rows(16).Copy()
            ws.Rows(f"17:{17 + extra - 1}").Insert(-4121)   # xlDown
            app.CutCopyMode = False
            for r in range(17, 17 + extra):                 # por si alguna quedo sin combinar
                _asegurar_merge(ws, r)
        elif n < base:
            ws.Rows("16:16").Delete()

        prim, ultimo = 15, 14 + n
        cd, igv, tot = ultimo + 1, ultimo + 2, ultimo + 3

        # limpia columnas de analisis/seguimiento internas (fuera del area de impresion)
        ws.Range(f"P14:AB{tot}").ClearContents()

        # D4 tiene formato de fecha larga: si hay fecha real la escribimos como fecha
        ws.Range("D4").Value = _valor_fecha(cot)
        ws.Range("D6").Value = cot.cliente
        ws.Range("D7").Value = cot.contacto
        ws.Range("D8").Value = cot.asunto
        ws.Range("D11").Value = cot.referencia_ticket
        ws.Range("K11").Value = cot.numero

        usa_div, val = _factor_txt(cot.margen, cot.modo_margen)
        for i, ln in enumerate(lineas):
            r = prim + i
            _asegurar_merge(ws, r)
            ws.Range(f"A{r}").Value = round(1.01 + i * 0.01, 2)
            ws.Range(f"B{r}").Value = ln.descripcion
            _ajustar_alto(ws, r, ln.descripcion)
            ws.Range(f"I{r}").Value = ln.unidad
            ws.Range(f"J{r}").Value = ln.cantidad
            ws.Range(f"O{r}").Value = round(float(ln.costo_unit or 0), 2)
            ws.Range(f"K{r}").Formula = f"=O{r}/{val}" if usa_div else f"=O{r}*{val}"
            ws.Range(f"L{r}").Formula = f"=K{r}*J{r}"
            ws.Range(f"M{r}").Formula = f"=J{r}*K{r}"

        ws.Range("M14").Formula = f"=SUM(L{prim}:L{ultimo})"
        ws.Range(f"M{cd}").Formula = "=M14"
        ws.Range(f"M{igv}").Formula = f"=M{cd}*L{igv}"
        ws.Range(f"M{tot}").Formula = f"=SUM(M{cd}:M{igv})"

        if cfg:
            _aplicar_imagenes(ws, cfg)
        _ajustar_pagina(ws, f"$A$1:$M${36 + (n - base)}")

        _asegurar_carpeta(xlsx_out)
        _asegurar_carpeta(pdf_out)
        wb.SaveAs(os.path.abspath(xlsx_out), FileFormat=51)
        pdf_ok = None
        if pdf_out:
            wb.ExportAsFixedFormat(0, os.path.abspath(pdf_out))
            pdf_ok = pdf_out if os.path.exists(pdf_out) else None
        return xlsx_out, pdf_ok
    finally:
        wb.Close(SaveChanges=False)


# ------------------------------------------------------------------ sustento
def _llenar_sustento(app, cot: Cotizacion, xlsx_out: str, pdf_out: str | None, cfg: dict | None = None):
    wb = app.Workbooks.Open(os.path.abspath(plantilla_sustento()))
    try:
        ws = wb.Worksheets(1)
        m = cot.margen
        gasto = round(sum(float(l.costo_unit or 0) * float(l.cantidad or 0)
                          for l in cot.lineas if l.es_gasto), 2)
        venta = round(sum(_venta_unit(l.costo_unit, m, cot.modo_margen) * float(l.cantidad or 0)
                          for l in cot.lineas), 2)

        # E5 debe ser fecha real: la plantilla calcula el mes con =TEXT(E5,"MMMM")
        ws.Range("E5").Value = _valor_fecha(cot)
        ws.Range("E7").Value = cot.numero
        ws.Range("G13").Value = cot.asunto
        ws.Range("G15").Value = cot.proveedor_principal.nombre or ""
        ws.Range("G17").Value = gasto
        ws.Range("G24").Value = cot.cliente
        ws.Range("G25").Value = cot.sede or cot.asunto
        ws.Range("G27").Value = venta

        if cfg:
            _aplicar_imagenes(ws, cfg)
        # Area de impresion: solo el bloque visible (evita 2da pagina por columnas vacias)
        _ajustar_pagina(ws, "$A$1:$N$49")

        _asegurar_carpeta(xlsx_out)
        _asegurar_carpeta(pdf_out)
        wb.SaveAs(os.path.abspath(xlsx_out), FileFormat=51)
        pdf_ok = None
        if pdf_out:
            wb.ExportAsFixedFormat(0, os.path.abspath(pdf_out))
            pdf_ok = pdf_out if os.path.exists(pdf_out) else None
        return xlsx_out, pdf_ok
    finally:
        wb.Close(SaveChanges=False)


# ------------------------------------------------------------------ API
def generar_con_excel(cot: Cotizacion, xlsx_coti=None, pdf_coti=None,
                      xlsx_sus=None, pdf_sus=None, cfg: dict | None = None) -> dict | None:
    """Genera los documentos pedidos en una sola sesion de Excel.
    Devuelve {'coti_xlsx':..,'coti_pdf':..,'sus_xlsx':..,'sus_pdf':..} o None si falla."""
    try:
        import win32com.client  # noqa
    except Exception:
        return None
    res = {}
    try:
        with _Excel() as app:
            # cada documento se maneja aparte: un fallo no arrastra al otro
            if xlsx_coti or pdf_coti:
                try:
                    cx, cp = _llenar_cotizacion(app, cot, xlsx_coti or _tmp("coti.xlsx"), pdf_coti, cfg)
                    res["coti_xlsx"] = cx if xlsx_coti else None
                    res["coti_pdf"] = cp
                    if not xlsx_coti and cx and os.path.exists(cx):
                        os.remove(cx)
                except Exception as e:
                    res["error_coti"] = str(e)
            if xlsx_sus or pdf_sus:
                try:
                    sx, sp = _llenar_sustento(app, cot, xlsx_sus or _tmp("sus.xlsx"), pdf_sus, cfg)
                    res["sus_xlsx"] = sx if xlsx_sus else None
                    res["sus_pdf"] = sp
                    if not xlsx_sus and sx and os.path.exists(sx):
                        os.remove(sx)
                except Exception as e:
                    res["error_sus"] = str(e)
        return res
    except Exception:
        return None


def _tmp(name: str) -> str:
    import tempfile
    return os.path.join(tempfile.gettempdir(), name)
