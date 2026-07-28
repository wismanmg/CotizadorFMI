"""Extraccion de datos desde el PDF de un proveedor.

Estrategia robusta para PDFs heterogeneos:
  1. Extrae texto con pdfplumber (PDFs con texto real, ej. TYC).
  2. Si no hay texto (PDF escaneado, ej. JM GLASS) avisa que requiere carga manual.
  3. Intenta detectar RUC, N° de cotizacion, montos y una descripcion,
     devolviendo una propuesta que el usuario SIEMPRE revisa/edita.
"""
from __future__ import annotations
import re
from dataclasses import dataclass, field
from typing import List, Optional

try:
    import pdfplumber
except Exception:  # pragma: no cover
    pdfplumber = None

# Motor OCR (carga perezosa; opcional)
_OCR = None
_OCR_INTENTADO = False


def _get_ocr():
    """Devuelve una instancia unica de RapidOCR, o None si no esta disponible."""
    global _OCR, _OCR_INTENTADO
    if _OCR_INTENTADO:
        return _OCR
    _OCR_INTENTADO = True
    try:
        from rapidocr_onnxruntime import RapidOCR
        _OCR = RapidOCR()
    except Exception:
        _OCR = None
    return _OCR


def ocr_disponible() -> bool:
    return _get_ocr() is not None


def _ocr_texto(ruta_pdf: str, max_paginas: int = 3) -> str:
    """Renderiza el PDF a imagen y aplica OCR. Reconstruye lineas por posicion."""
    ocr = _get_ocr()
    if ocr is None:
        return ""
    try:
        import pypdfium2 as pdfium
        import numpy as np
    except Exception:
        return ""
    lineas_texto = []
    try:
        pdf = pdfium.PdfDocument(ruta_pdf)
        n = min(len(pdf), max_paginas)
        for i in range(n):
            img = pdf[i].render(scale=2.5).to_pil()
            resultado, _ = ocr(np.array(img))
            if not resultado:
                continue
            # cada item: [box(4 puntos), texto, confianza]
            cajas = []
            for item in resultado:
                box, texto = item[0], item[1]
                ys = [p[1] for p in box]; xs = [p[0] for p in box]
                cajas.append((sum(ys) / 4.0, min(xs), texto))
            cajas.sort(key=lambda c: (c[0], c[1]))
            # agrupa por filas (misma altura aproximada)
            filas = []
            for yc, x, txt in cajas:
                if filas and abs(yc - filas[-1][0]) < 12:
                    filas[-1][1].append((x, txt))
                else:
                    filas.append([yc, [(x, txt)]])
            for yc, partes in filas:
                partes.sort(key=lambda p: p[0])
                lineas_texto.append(" ".join(t for _, t in partes))
    except Exception:
        return ""
    return "\n".join(lineas_texto)


@dataclass
class LineaExtraida:
    descripcion: str = ""
    unidad: str = "UND"
    cantidad: float = 1.0
    costo_unit: float = 0.0


@dataclass
class ResultadoExtraccion:
    texto: str = ""
    es_escaneado: bool = False
    ruc: str = ""
    numero_cotizacion: str = ""
    nombre_proveedor: str = ""
    subtotal: Optional[float] = None
    total: Optional[float] = None
    lineas: List[LineaExtraida] = field(default_factory=list)
    mensaje: str = ""


_NUM = re.compile(r"(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})|\d+[.,]\d{2})")

# Lineas que NO son items: totales, datos de cabecera, condiciones comerciales,
# notas y pie de firma. Evita que el parser las convierta en lineas de la cotizacion.
_RUIDO = re.compile(
    r"(sub\s*total|i\.?g\.?v|total|costo\s*directo|precio\s*(con|sub)|valor\s*venta|"
    r"descuento|adelanto|saldo|adicional|r\.?u\.?c|fecha|cliente|contacto|asunto|"
    r"direccion|telefono|tel[eé]f|email|correo|moneda|nota|cta\b|banco|cuenta|son:|"
    r"presupuesto|item\b|descripci|unidad|u\.m\.|cant\b|"
    # condiciones comerciales
    r"condicion|comercial|validez|propuesta|calendario|forma\s*de\s*pago|factura\s*a\s*\d|"
    r"plazo|tiempo\s*de\s*ejec|lugar\s*de\s*entrega|orden\s*a\b|coordinacion|"
    r"seg[uú]n\s*requerimiento|referencia|ticket|solicitud|"
    # pie / firma
    r"firma|jefe\b|gerente|atentamente|saludos|outsourcing|panorama)", re.I)


def _limpiar_desc(texto: str) -> str:
    """Limpieza conservadora de la descripcion extraida."""
    t = texto or ""
    t = re.sub(r"[|]+", " ", t)
    t = re.sub(r"([,;.])(?=[A-Za-zÁ-úÑñ])", r"\1 ", t)      # espacio tras coma/punto
    t = re.sub(r"\b(?:S\s*/\.?|Gbl|und\.?)\s*$", "", t, flags=re.I)  # colas de tabla
    t = re.sub(r"(?:\s*S\s*/\.?\s*)+$", "", t)
    t = re.sub(r"^\d+[.,]\d+\s+", "", t)                    # numeracion de item inicial
    t = re.sub(r"\s{2,}", " ", t)
    return t.strip(" .-,")


def _a_float(txt: str) -> Optional[float]:
    if not txt:
        return None
    t = txt.strip().replace(" ", "")
    # normaliza formato peruano 1,234.56  o 1.234,56
    if t.count(",") and t.count(".") and t.rfind(",") > t.rfind("."):
        t = t.replace(".", "").replace(",", ".")
    else:
        t = t.replace(",", "")
    try:
        return round(float(t), 2)
    except ValueError:
        return None


def extraer(ruta_pdf: str) -> ResultadoExtraccion:
    res = ResultadoExtraccion()
    if pdfplumber is None:
        res.mensaje = "pdfplumber no esta instalado."
        return res

    textos = []
    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            for p in pdf.pages:
                textos.append(p.extract_text() or "")
    except Exception as e:
        res.mensaje = f"No se pudo abrir el PDF: {e}"
        return res

    texto = "\n".join(textos).strip()
    uso_ocr = False

    if len(texto) < 15:
        # PDF escaneado: intenta OCR automatico
        res.es_escaneado = True
        texto_ocr = _ocr_texto(ruta_pdf)
        if texto_ocr:
            texto = texto_ocr
            uso_ocr = True
        else:
            res.texto = texto
            res.mensaje = ("El PDF parece escaneado y el OCR no esta disponible. "
                           "Ingresa los datos del proveedor manualmente.")
            return res

    res.texto = texto

    # RUC (11 digitos)
    m = re.search(r"R\.?U\.?C\.?\s*(?:N[°ºo]?)?\s*[:.]?\s*(\d{11})", texto, re.I)
    if not m:
        m = re.search(r"\b(\d{11})\b", texto)
    if m:
        res.ruc = m.group(1)

    # N° de cotizacion / presupuesto: codigo con guion tipo 001-7401 o 135-26
    m = re.search(r"\b(\d{2,4}-\d{2,6})\b", texto)
    if m and m.group(1) != res.ruc:
        res.numero_cotizacion = m.group(1)

    # Totales: SUB TOTAL primero; TOTAL = ultima aparicion (tras IGV)
    subs = re.findall(r"SUB\s*TOTAL[^\d]{0,20}" + _NUM.pattern, texto, re.I)
    if subs:
        res.subtotal = _a_float(subs[-1])
    tots = re.findall(r"TOTAL[^\d]{0,20}" + _NUM.pattern, texto, re.I)
    if tots:
        res.total = _a_float(tots[-1])

    # Heuristica de lineas: filas con descripcion + un monto al final
    # En las tablas la descripcion suele ocupar varias lineas y solo la ultima
    # trae el importe: acumulamos las lineas sueltas y las unimos al item.
    buffer: list[str] = []
    for linea in texto.splitlines():
        l = linea.strip()
        if not l or len(l) < 4:
            continue
        if _RUIDO.search(l) or re.match(r"^\d{2}[.,]\d{2}\b", l):
            buffer.clear()          # lo anterior no pertenece a ningun item
            continue

        nums = _NUM.findall(l)
        resto = _NUM.sub("", l).strip(" .-|")
        if not nums:
            if sum(c.isalpha() for c in l) >= 6:
                buffer.append(l)
                if len(buffer) > 4:
                    buffer.pop(0)
            continue

        monto = _a_float(nums[-1])
        desc = _limpiar_desc(" ".join(buffer + ([resto] if resto else [])))
        buffer.clear()
        if monto and monto > 0 and sum(c.isalpha() for c in desc) >= 12:
            res.lineas.append(LineaExtraida(descripcion=desc, costo_unit=monto))

    # Si no detecto lineas pero si un subtotal/total, crea una linea unica
    if not res.lineas and (res.subtotal or res.total):
        base = res.subtotal or (round((res.total or 0) / 1.18, 2))
        res.lineas.append(LineaExtraida(descripcion="(Revisar descripcion)", costo_unit=base))

    origen = "OCR (PDF escaneado)" if uso_ocr else "texto del PDF"
    res.mensaje = (f"Datos leidos por {origen}. Se detectaron {len(res.lineas)} "
                   f"posibles lineas. Revisa y corrige antes de generar.")
    return res
