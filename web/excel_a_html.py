# -*- coding: utf-8 -*-
"""Convierte una hoja de Excel en HTML calcado: colores, bordes, fuentes, merges,
anchos y altos reales, y con el MISMO factor de reduccion que usa Excel al
imprimir en A4 ("ajustar a 1 pagina"). Asi la version web sale igual al original.
"""
import os
import pythoncom
import win32com.client as w

xlNone = -4142
xlLeft, xlCenter, xlRight = -4131, -4108, -4152
xlTop, xlCenterV, xlBottom = -4160, -4108, -4107
xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight = 7, 8, 9, 10
xlDouble = -4119

# ancho util de una A4 vertical con margenes de 6 mm  (en puntos)
ANCHO_UTIL_PT = (210 - 12) * 72 / 25.4        # ~561 pt


def bgr2hex(c):
    c = int(c)
    return "#%02X%02X%02X" % (c & 255, (c >> 8) & 255, (c >> 16) & 255)


def align_h(rng, valor):
    v = rng.HorizontalAlignment
    if v == xlCenter:
        return "center"
    if v == xlRight:
        return "right"
    if v == xlLeft:
        return "left"
    return "right" if isinstance(valor, (int, float)) else "left"


def align_v(v):
    return {xlTop: "top", xlBottom: "bottom", xlCenterV: "middle"}.get(v, "bottom")


def borde_css(rng, edge, k):
    b = rng.Borders(edge)
    if b.LineStyle == xlNone:
        return "none"
    g = 0.8 if b.Weight <= 2 else (1.4 if b.Weight == 3 else 2.0)
    est = "double" if b.LineStyle == xlDouble else "solid"
    return "%.2fpt %s %s" % (g * k, est, bgr2hex(b.Color))


def ancho_texto_pt(texto, pt, bold):
    """Ancho aproximado de un texto en Arial (suficiente para decidir desbordes)."""
    return len(texto) * pt * (0.58 if bold else 0.53)


def es_contable(fmt):
    return "*" in (fmt or "")


class Conv:
    def __init__(self, ws, f_ini, f_fin, c_ini, c_fin, ids, escala):
        self.ws, self.ids, self.k = ws, ids, escala
        self.f_ini, self.f_fin, self.c_ini, self.c_fin = f_ini, f_fin, c_ini, c_fin
        self.anchos = {c: ws.Columns(c).Width for c in range(c_ini, c_fin + 1)}
        self.ocupada = set()          # celdas absorbidas por merge o por desborde

    def _texto(self, rng):
        try:
            return rng.Text or ""
        except Exception:
            return ""

    def _vacia(self, r, c):
        if (r, c) in self.ocupada:
            return False
        cel = self.ws.Cells(r, c)
        if cel.MergeArea.Count > 1:
            return False
        return self._texto(cel).strip() == ""

    def _colspan_desborde(self, r, c, texto, pt, bold, ali):
        """Excel deja que el texto invada las celdas vacias de al lado."""
        if ali != "left" or not texto.strip():
            return 1
        necesario = ancho_texto_pt(texto, pt, bold)
        disponible = self.anchos[c]
        span = 1
        cc = c + 1
        while disponible < necesario and cc <= self.c_fin and self._vacia(r, cc):
            fondo_a = self.ws.Cells(r, c).Interior.Color
            fondo_b = self.ws.Cells(r, cc).Interior.Color
            if fondo_a != fondo_b:            # no invadir bloques de otro color
                break
            disponible += self.anchos[cc]
            self.ocupada.add((r, cc))
            span += 1
            cc += 1
        return span

    def celda(self, r, c):
        if (r, c) in self.ocupada:
            return None
        ws, k = self.ws, self.k
        rng = ws.Cells(r, c)
        ma = rng.MergeArea
        if (ma.Row, ma.Column) != (r, c):
            return None
        filas, cols = ma.Rows.Count, ma.Columns.Count
        for rr in range(ma.Row, ma.Row + filas):
            for cc in range(ma.Column, ma.Column + cols):
                if (rr, cc) != (r, c):
                    self.ocupada.add((rr, cc))

        valor, texto = rng.Value, self._texto(rng)
        f = rng.Font
        pt = (f.Size or 10)
        bold = bool(f.Bold)
        ali = align_h(rng, valor)
        wrap = bool(rng.WrapText)

        if cols == 1 and not wrap:
            cols_extra = self._colspan_desborde(r, c, texto, pt, bold, ali)
            cols = max(cols, cols_extra)

        est = [
            "font-family:Arial,Helvetica,sans-serif",
            "font-size:%.2fpt" % (pt * k),
            "color:%s" % bgr2hex(f.Color),
            "text-align:%s" % ali,
            "vertical-align:%s" % align_v(rng.VerticalAlignment),
            "white-space:%s" % ("normal" if wrap else "nowrap"),
            "overflow:hidden",
        ]
        if bold:
            est.append("font-weight:bold")
        if f.Italic:
            est.append("font-style:italic")
        try:
            if rng.Interior.ColorIndex != xlNone:
                est.append("background:%s" % bgr2hex(rng.Interior.Color))
        except Exception:
            pass
        for edge, lado in ((xlEdgeTop, "top"), (xlEdgeRight, "right"),
                           (xlEdgeBottom, "bottom"), (xlEdgeLeft, "left")):
            est.append("border-%s:%s" % (lado, borde_css(ma, edge, k)))

        attrs = ""
        if cols > 1:
            attrs += ' colspan="%d"' % cols
        if filas > 1:
            attrs += ' rowspan="%d"' % filas
        cid = self.ids.get("%s%d" % (chr(64 + c), r))
        if cid:
            attrs += ' id="%s"' % cid
            contenido = ""
        else:
            contenido = texto.replace("&", "&amp;").replace("<", "&lt;")
            # formato contable: simbolo pegado a la izquierda, numero a la derecha
            if contenido and es_contable(rng.NumberFormat) and isinstance(valor, (int, float)):
                p = contenido.strip().split(None, 1)
                if len(p) == 2:
                    contenido = ('<span style="float:left">%s</span>%s' % (p[0], p[1]))
        return '<td%s style="%s">%s</td>' % (attrs, ";".join(est), contenido or "&nbsp;")

    def html(self, clase):
        k = self.k
        cols = "".join('<col style="width:%.2fpt">' % (self.anchos[c] * k)
                       for c in range(self.c_ini, self.c_fin + 1))
        filas = []
        for r in range(self.f_ini, self.f_fin + 1):
            alto = self.ws.Rows(r).RowHeight * k
            tds = [td for td in (self.celda(r, c)
                                 for c in range(self.c_ini, self.c_fin + 1)) if td]
            filas.append('<tr data-fila="%d" style="height:%.2fpt">%s</tr>'
                         % (r, alto, "".join(tds)))
        ancho = sum(self.anchos.values()) * k
        return ('<table class="%s" style="width:%.2fpt"><colgroup>%s</colgroup>'
                '<tbody>%s</tbody></table>' % (clase, ancho, cols, "".join(filas)))


def convertir(ruta_xlsx, fila_fin, col_fin, ids, clase):
    pythoncom.CoInitialize()
    app = w.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    try:
        wb = app.Workbooks.Open(os.path.abspath(ruta_xlsx), ReadOnly=1)
        ws = wb.Worksheets(1)
        ancho_pt = sum(ws.Columns(c).Width for c in range(1, col_fin + 1))
        escala = min(1.0, ANCHO_UTIL_PT / ancho_pt)      # igual que "ajustar a 1 pagina"
        conv = Conv(ws, 1, fila_fin, 1, col_fin, ids, escala)
        html = conv.html(clase)
        wb.Close(False)
        print("   %s -> escala %.3f" % (os.path.basename(ruta_xlsx), escala))
        return html
    finally:
        app.Quit()
        pythoncom.CoUninitialize()
