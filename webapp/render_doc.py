# -*- coding: utf-8 -*-
"""Arma el HTML final de los documentos (con los datos ya puestos) y lo
convierte a PDF usando el navegador del sistema en modo silencioso.
Asi la descarga sale directa, sin abrir el dialogo de impresion.
"""
from __future__ import annotations

import base64
import os
import re
import subprocess
import tempfile
import uuid

BASE = os.path.dirname(os.path.abspath(__file__))
DOC = os.path.join(BASE, "static", "doc")

# navegadores que sirven para imprimir a PDF (el primero que exista)
NAVEGADORES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]

CSS = """
 @page{size:A4 portrait;margin:6mm;}
 html,body{margin:0;padding:0;background:#fff;}
 *{-webkit-print-color-adjust:exact !important;print-color-adjust:exact !important;}
 body{font-family:Arial,Helvetica,sans-serif;}
 table.xl{border-collapse:collapse;table-layout:fixed;}
 table.xl td{padding:0 1.5px;line-height:1.1;overflow:hidden;}
 img.logo{max-width:100%;max-height:100%;object-fit:contain;display:block;margin:auto;}
 .cfirma{position:relative;text-align:center;}
 .cfirma img.fi{max-width:62%;max-height:78%;object-fit:contain;display:block;margin:0 auto;}
"""


# ------------------------------------------------------------ utilidades HTML
def _uri(nombre: str) -> str:
    with open(os.path.join(DOC, nombre), "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode()


def _frag(nombre: str) -> str:
    with open(os.path.join(DOC, nombre), encoding="utf-8") as f:
        return f.read()


def _esc(v) -> str:
    return (str(v if v is not None else "")
            .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _poner(html: str, cid: str, valor: str) -> str:
    """Escribe el contenido dentro del <td id="cid">."""
    pat = re.compile(r'(<td[^>]*\bid="%s"[^>]*>).*?(</td>)' % re.escape(cid), re.S)
    return pat.sub(lambda m: m.group(1) + valor + m.group(2), html, count=1)


def _monto(n) -> str:
    try:
        n = float(n)
    except (TypeError, ValueError):
        return ""
    txt = "{:,.2f}".format(n)
    return '<span style="float:left">S/.</span>' + txt


def _contenido_td(td: str, valor: str) -> str:
    return re.sub(r"^(<td[^>]*>).*(</td>)$", lambda m: m.group(1) + valor + m.group(2),
                  td, flags=re.S)


# ------------------------------------------------------------ cotizacion
def render_cotizacion(d: dict, c: dict) -> str:
    html = _frag("cotizacion.html")

    # 1) fila modelo de items y limpieza de las filas 15 y 16
    m = re.search(r'<tr data-fila="15".*?</tr>', html, re.S)
    tpl = m.group(0) if m else ""
    for f in ("15", "16"):
        html = re.sub(r'<tr data-fila="%s".*?</tr>' % f, "", html, count=1, flags=re.S)

    # 2) filas reales
    filas = []
    for ln in c.get("lineas", []):
        tds = re.findall(r"<td[^>]*>.*?</td>", tpl, re.S)
        if len(tds) >= 7:
            vals = [ln["n"], _esc(ln["desc"]), _esc(ln["um"]), _esc(ln["cant"]),
                    _monto(ln["unit"]), _monto(ln["parcial"]), _monto(ln["parcial"])]
            nuevos = [_contenido_td(td, v) for td, v in zip(tds, vals)]
            fila = re.sub(r'\sdata-fila="\d+"', "", tpl)
            fila = re.sub(r"<td[^>]*>.*?</td>", lambda _m, it=iter(nuevos): next(it),
                          fila, flags=re.S)
            filas.append(fila)
    html = re.sub(r'(<tr data-fila="17")', "".join(filas) + r"\1", html, count=1)

    # 3) datos
    val = {
        "c_fecha": _esc(d.get("fecha")), "c_cli": _esc(d.get("cliente")),
        "c_con": _esc(d.get("contacto")), "c_asu": _esc(d.get("asunto")),
        "c_tick": _esc(d.get("ticket")), "c_num": _esc(d.get("numero")),
        "c_grupo": _esc(d.get("asunto")),
        "c_gsub": _monto(c["subtotal"]), "c_cd": _monto(c["subtotal"]),
        "c_pigv": "%d%%" % c["igv_pct"], "c_igv": _monto(c["igv_monto"]),
        "c_tot": _monto(c["total"]),
        "c_logo": '<img class="logo" src="%s">' % _uri("logo.png"),
        "c_firma": '<img class="fi" src="%s">' % _uri("firma.png"),
    }
    for k, v in val.items():
        html = _poner(html, k, v)
    html = html.replace('id="c_firma"', 'id="c_firma" class="cfirma"')
    return html


# ------------------------------------------------------------ sustento
def render_sustento(d: dict, c: dict) -> str:
    html = _frag("sustento.html")
    val = {
        "s_fecha": _esc(d.get("fecha")), "s_num": _esc(d.get("numero")),
        "s_mes": _esc(d.get("mes")), "s_mot": _esc(d.get("asunto")),
        "s_prov": _esc(d.get("proveedor")), "s_cli": _esc(d.get("cliente")),
        "s_sede": _esc(d.get("sede") or d.get("asunto")),
        "s_g1": _monto(c["gasto"]), "s_g2": _monto(c["gasto_igv"]),
        "s_g3": _monto(c["gasto_total"]),
        "s_v1": _monto(c["subtotal"]), "s_v2": _monto(c["igv_monto"]),
        "s_v3": _monto(c["total_sustento"]),
        "r_v1": _monto(c["subtotal"]), "r_v2": _monto(c["total_sustento"]),
        "r_g1": _monto(c["gasto"]), "r_g2": _monto(c["gasto_total"]),
        "r_u1": _monto(c["utilidad"]), "r_u2": _monto(c["utilidad_igv"]),
        "r_pct": ("%.2f%%" % c["utilidad_pct"]) if c["subtotal"] else "",
        "s_logo": '<img class="logo" src="%s">' % _uri("logo.png"),
        "s_firma": '<img class="fi" src="%s">' % _uri("firma.png"),
    }
    for k, v in val.items():
        html = _poner(html, k, v)
    html = html.replace('id="s_firma"', 'id="s_firma" class="cfirma"')
    return html


# ------------------------------------------------------------ PDF
def envolver(*documentos: str) -> str:
    cuerpo = '<div style="page-break-after:always">%s</div>'
    partes = [cuerpo % doc for doc in documentos[:-1]] + [documentos[-1]]
    return ("<!DOCTYPE html><html><head><meta charset='utf-8'><style>%s</style>"
            "</head><body>%s</body></html>" % (CSS, "".join(partes)))


def navegador() -> str | None:
    for exe in NAVEGADORES:
        if os.path.exists(exe):
            return exe
    return None


def html_a_pdf(html: str) -> bytes | None:
    """Convierte el HTML a PDF con Edge/Chrome en segundo plano."""
    exe = navegador()
    if not exe:
        return None
    tmp = tempfile.gettempdir()
    uid = uuid.uuid4().hex[:10]
    fhtml = os.path.join(tmp, "cotfmi_%s.html" % uid)
    fpdf = os.path.join(tmp, "cotfmi_%s.pdf" % uid)
    perfil = os.path.join(tmp, "cotfmi_perfil_%s" % uid)
    try:
        with open(fhtml, "w", encoding="utf-8") as f:
            f.write(html)
        subprocess.run(
            [exe, "--headless=new", "--disable-gpu", "--no-sandbox",
             "--no-pdf-header-footer", "--user-data-dir=" + perfil,
             "--print-to-pdf=" + fpdf, fhtml],
            timeout=90, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        if os.path.exists(fpdf):
            with open(fpdf, "rb") as f:
                return f.read()
        return None
    except Exception:
        return None
    finally:
        import shutil
        for p in (fhtml, fpdf):
            try:
                os.remove(p)
            except Exception:
                pass
        shutil.rmtree(perfil, ignore_errors=True)
