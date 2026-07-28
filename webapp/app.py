# -*- coding: utf-8 -*-
"""Cotizador FMI — aplicacion web (Flask).

Corre en cualquier PC con Python y se abre en el navegador.
Tambien se puede publicar en internet (Render, PythonAnywhere, etc.) para
entrar desde una laptop donde no se pueden instalar programas.

Uso local:   python app.py    ->   http://localhost:8000
"""
from __future__ import annotations

import datetime as dt
import json
import os
import sys
import tempfile

from flask import (Flask, jsonify, render_template, request,
                   send_from_directory)

BASE = os.path.dirname(os.path.abspath(__file__))
DOC = os.path.join(BASE, "static", "doc")
RAIZ = os.path.dirname(BASE)          # carpeta del proyecto (tiene cotizador/)
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

app = Flask(__name__)
# recarga la plantilla si se edita, sin tener que reiniciar el servidor
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.jinja_env.auto_reload = True

DIAS = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
         "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

CONFIG = os.path.join(BASE, "datos_web.json")
DEFAULTS = {
    "proximo_numero": 251, "anio": 26, "prefijo": "FMI",
    "margen": 0.15, "igv": 0.18,
    "firma_nombre": "Carlos Miguel Toledo Santa Cruz",
    "firma_cargo": "Jefe de Serv. Gener. - Chorrillos",
}


# --------------------------------------------------------------- config
def cargar_cfg() -> dict:
    cfg = dict(DEFAULTS)
    if os.path.exists(CONFIG):
        try:
            with open(CONFIG, encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception:
            pass
    return cfg


def guardar_cfg(cfg: dict) -> None:
    try:
        with open(CONFIG, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass          # en hosting de solo lectura simplemente no persiste


# --------------------------------------------------------------- calculo
def precio_venta(costo: float, margen: float, modo: str = "sobre_venta") -> float:
    """Regla de negocio: venta = costo / (1 - margen).  755 / 0.85 = 888.24"""
    costo = float(costo or 0)
    if modo == "sobre_venta":
        return costo / ((1 - margen) or 1.0)
    return costo * (1 + margen)


def calcular(datos: dict) -> dict:
    margen = float(datos.get("margen", 0.15))
    igv = float(datos.get("igv", 0.18))
    modo = datos.get("modo", "sobre_venta")

    lineas, subtotal, gasto = [], 0.0, 0.0
    for i, l in enumerate(datos.get("lineas", [])):
        desc = (l.get("desc") or "").strip()
        costo = float(l.get("costo") or 0)
        if not desc and not costo:
            continue
        cant = float(l.get("cant") or 0)
        vu = precio_venta(costo, margen, modo)
        parcial = vu * cant
        subtotal += parcial
        if l.get("gasto", True):
            gasto += costo * cant
        lineas.append({
            "n": "01.%02d" % (len(lineas) + 1), "desc": desc,
            "um": l.get("um") or "UND", "cant": l.get("cant"),
            "unit": round(vu, 2), "parcial": round(parcial, 2),
        })

    # Igual que las plantillas: la COTIZACION encadena formulas sobre el valor
    # sin redondear (1,214.71) y el SUSTENTO parte del monto ya redondeado
    # (1,214.70). Esa diferencia de un centavo existe en los archivos originales.
    sub_r = round(subtotal, 2)
    gasto = round(gasto, 2)
    igv_monto = round(subtotal * igv, 2)
    total_coti = round(subtotal * (1 + igv), 2)
    total_sus = round(sub_r * (1 + igv), 2)
    utilidad = round(sub_r - gasto, 2)
    return {
        "lineas": lineas,
        "subtotal": sub_r,
        "igv_pct": round(igv * 100),
        "igv_monto": igv_monto,
        "total": total_coti,
        "total_sustento": total_sus,
        "gasto": gasto,
        "gasto_igv": round(gasto * igv, 2),
        "gasto_total": round(gasto * (1 + igv), 2),
        "utilidad": utilidad,
        "utilidad_igv": round(utilidad * (1 + igv), 2),
        "utilidad_pct": round(utilidad / sub_r * 100, 2) if sub_r else 0,
    }


# --------------------------------------------------------------- rutas
def frag(nombre: str) -> str:
    with open(os.path.join(DOC, nombre), encoding="utf-8") as f:
        return f.read()


@app.route("/")
def index():
    cfg = cargar_cfg()
    hoy = dt.date.today()
    return render_template(
        "index.html",
        coti_html=frag("cotizacion.html"),
        sus_html=frag("sustento.html"),
        cfg=cfg,
        numero="%s - %04d-%d" % (cfg["prefijo"], cfg["proximo_numero"], cfg["anio"]),
        fecha=f"{DIAS[hoy.weekday()]}, {hoy.day} de {MESES[hoy.month - 1]} de {hoy.year}",
    )


@app.post("/api/calcular")
def api_calcular():
    """El calculo lo hace Python (misma regla que la app de escritorio)."""
    return jsonify(calcular(request.get_json(force=True) or {}))


@app.post("/api/correlativo")
def api_correlativo():
    cfg = cargar_cfg()
    cfg["proximo_numero"] = int(cfg["proximo_numero"]) + 1
    guardar_cfg(cfg)
    return jsonify({"numero": "%s - %04d-%d" % (cfg["prefijo"],
                                                cfg["proximo_numero"], cfg["anio"])})


def _fecha_dt(texto: str):
    """Convierte 'lunes, 27 de Julio de 2026' (o 27/07/2026) a date."""
    import re as _re
    t = (texto or "").strip().lower()
    m = _re.search(r"(\d{1,2})\s+de\s+([a-záéíóú]+)\s+de\s+(\d{4})", t)
    if m:
        mes = (m.group(2).replace("á", "a").replace("é", "e").replace("í", "i")
                          .replace("ó", "o").replace("ú", "u"))
        nombres = [x.lower() for x in MESES]
        if mes in nombres:
            return dt.date(int(m.group(3)), nombres.index(mes) + 1, int(m.group(1)))
    for f in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(t, f).date()
        except ValueError:
            pass
    return dt.date.today()


@app.post("/api/excel")
def api_excel():
    """Genera el Excel con las plantillas reales (via Microsoft Excel) y lo descarga."""
    import io
    import shutil
    from flask import send_file

    datos = request.get_json(force=True) or {}
    cual = datos.get("doc", "coti")
    try:
        from cotizador.modelo import Cotizacion, Linea, Proveedor
        from cotizador.gen_com import generar_con_excel, excel_disponible
    except Exception as e:
        return jsonify({"ok": False, "error": "No se pudo cargar el generador: %s" % e}), 500
    if not excel_disponible():
        return jsonify({"ok": False,
                        "error": "Para el Excel hace falta Microsoft Excel instalado. "
                                 "Podés descargar el PDF igual."}), 500

    cot = Cotizacion(
        numero=datos.get("numero") or "", fecha=datos.get("fecha") or "",
        fecha_dt=_fecha_dt(datos.get("fecha")),
        cliente=datos.get("cliente") or "", contacto=datos.get("contacto") or "",
        asunto=datos.get("asunto") or "",
        referencia_ticket=datos.get("ticket") or "", sede=datos.get("sede") or "",
        margen=float(datos.get("margen", 0.15)), modo_margen="sobre_venta",
        proveedor_principal=Proveedor(nombre=datos.get("proveedor") or ""),
    )
    cot.lineas = [
        Linea(l.get("desc") or "", l.get("um") or "UND", float(l.get("cant") or 0),
              float(l.get("costo") or 0), bool(l.get("gasto", True)))
        for l in datos.get("lineas", [])
        if (l.get("desc") or l.get("costo"))
    ] or [Linea("", "UND", 1, 0, True)]

    carpeta = os.path.join(tempfile.gettempdir(), "cotfmi_xls_%s" % os.getpid())
    os.makedirs(carpeta, exist_ok=True)
    fx = os.path.join(carpeta, "coti.xlsx")
    fs = os.path.join(carpeta, "sus.xlsx")
    try:
        r = generar_con_excel(cot,
                              xlsx_coti=fx if cual in ("coti", "ambos") else None,
                              xlsx_sus=fs if cual in ("sus", "ambos") else None)
        if not r or (r.get("error_coti") or r.get("error_sus")):
            return jsonify({"ok": False, "error": "Excel no pudo generar el archivo."}), 500

        num = (datos.get("numero") or "cotizacion").replace("/", "-").replace(" ", "")
        if cual == "ambos":
            import zipfile
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                z.write(fx, "Cotizacion %s.xlsx" % num)
                z.write(fs, "Sustento de Gasto %s.xlsx" % num)
            buf.seek(0)
            return send_file(buf, mimetype="application/zip", as_attachment=True,
                             download_name="Cotizacion y Sustento %s.zip" % num)

        ruta = fx if cual == "coti" else fs
        nombre = "Cotizacion" if cual == "coti" else "Sustento de Gasto"
        with open(ruta, "rb") as f:
            data = f.read()
        return send_file(
            io.BytesIO(data),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True, download_name="%s %s.xlsx" % (nombre, num))
    except Exception as e:
        return jsonify({"ok": False, "error": "No se pudo generar el Excel: %s" % e}), 500
    finally:
        shutil.rmtree(carpeta, ignore_errors=True)


@app.post("/api/leer-pdf")
def api_leer_pdf():
    """Lee la cotizacion del proveedor. Usa el MISMO motor que la app de
    escritorio: pdfplumber para PDFs con texto y OCR si viene escaneado."""
    archivo = request.files.get("pdf")
    if not archivo or not archivo.filename:
        return jsonify({"ok": False, "error": "No se recibió ningún archivo."}), 400
    if not archivo.filename.lower().endswith(".pdf"):
        return jsonify({"ok": False, "error": "El archivo debe ser un PDF."}), 400

    tmp = os.path.join(tempfile.gettempdir(),
                       "cotfmi_%d.pdf" % abs(hash(archivo.filename + str(dt.datetime.now()))))
    try:
        archivo.save(tmp)
        try:
            from cotizador.extraer_pdf import extraer
        except Exception as e:
            return jsonify({"ok": False,
                            "error": "No se pudo cargar el lector de PDF: %s" % e}), 500

        r = extraer(tmp)
        return jsonify({
            "ok": True,
            "mensaje": r.mensaje,
            "proveedor": r.nombre_proveedor or "",
            "escaneado": bool(r.es_escaneado),
            "ruc": r.ruc or "",
            "numero": r.numero_cotizacion or "",
            "total": r.total,
            "lineas": [{"desc": l.descripcion, "um": l.unidad or "UND",
                        "cant": l.cantidad or 1, "costo": l.costo_unit,
                        "gasto": True} for l in r.lineas],
        })
    except Exception as e:
        return jsonify({"ok": False, "error": "No se pudo leer el PDF: %s" % e}), 500
    finally:
        try:
            os.remove(tmp)
        except Exception:
            pass


@app.post("/api/pdf")
def api_pdf():
    """Genera el PDF en el servidor y lo devuelve como descarga directa
    (sin abrir el dialogo de impresion del navegador)."""
    import io
    from flask import send_file
    import render_doc as rd

    datos = request.get_json(force=True) or {}
    cual = datos.get("doc", "coti")
    calc = calcular(datos)

    docs = []
    if cual in ("coti", "ambos"):
        docs.append(rd.render_cotizacion(datos, calc))
    if cual in ("sus", "ambos"):
        docs.append(rd.render_sustento(datos, calc))
    if not docs:
        return jsonify({"ok": False, "error": "Documento no válido."}), 400

    pdf = rd.html_a_pdf(rd.envolver(*docs))
    if not pdf:
        return jsonify({"ok": False,
                        "error": "No se encontró Edge o Chrome para generar el PDF."}), 500

    num = (datos.get("numero") or "cotizacion").replace("/", "-").replace(" ", "")
    nombre = {"coti": "Cotizacion", "sus": "Sustento de Gasto",
              "ambos": "Cotizacion y Sustento"}[cual]
    return send_file(io.BytesIO(pdf), mimetype="application/pdf",
                     as_attachment=True, download_name="%s %s.pdf" % (nombre, num))


@app.get("/doc/<path:archivo>")
def doc(archivo):
    return send_from_directory(DOC, archivo)


@app.get("/salud")
def salud():
    return {"ok": True}


if __name__ == "__main__":
    puerto = int(os.environ.get("PORT", 8000))
    print("\n  Cotizador FMI (web)  ->  http://localhost:%d\n" % puerto)
    app.run(host="0.0.0.0", port=puerto, debug=False)
