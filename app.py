"""Cotizador FMI - Panorama Outsourcing
Aplicacion de escritorio para generar cotizaciones a partir de las de proveedores.
"""
from __future__ import annotations
import os
import re
import sys
import subprocess
import datetime as dt
import traceback

import customtkinter as ctk
from tkinter import filedialog, messagebox

from cotizador import config as cfgmod
from cotizador.config import (cargar_config, guardar_config, siguiente_correlativo,
                              consumir_correlativo, guardar_logo, ruta_logo,
                              guardar_firma, ruta_firma, restaurar_logo,
                              restaurar_firma, es_personalizado)
from cotizador.modelo import Cotizacion, Linea, Proveedor
from cotizador.servicio import generar_todo
from cotizador.extraer_pdf import extraer
from cotizador.control import registrar_en_maestro

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
DIAS = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]


def fecha_larga(d: dt.date) -> str:
    return f"{DIAS[d.weekday()]}, {d.day} de {MESES[d.month-1].capitalize()} de {d.year}"


def parsear_fecha(texto: str) -> dt.date | None:
    """Convierte el texto del campo Fecha a un date real (o None si no se entiende)."""
    t = (texto or "").strip().lower()
    if not t:
        return None
    # formato largo: "jueves, 9 de julio de 2026"
    m = re.search(r"(\d{1,2})\s+de\s+([a-záéíóú]+)\s+de\s+(\d{4})", t)
    if m:
        dia, mes_txt, anio = int(m.group(1)), m.group(2), int(m.group(3))
        mes_txt = (mes_txt.replace("á", "a").replace("é", "e").replace("í", "i")
                          .replace("ó", "o").replace("ú", "u"))
        if mes_txt in MESES:
            return dt.date(anio, MESES.index(mes_txt) + 1, dia)
    # formatos numericos: 9/7/2026, 09-07-2026, 2026-07-09
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            return dt.datetime.strptime(t, fmt).date()
        except ValueError:
            continue
    return None


class FilaLinea:
    """Una fila editable de la tabla de lineas de venta."""
    def __init__(self, master, app, linea: Linea):
        self.app = app
        self.linea = linea
        self.desc = ctk.CTkEntry(master, width=340)
        self.desc.insert(0, linea.descripcion)
        self.um = ctk.CTkEntry(master, width=55); self.um.insert(0, linea.unidad)
        self.cant = ctk.CTkEntry(master, width=55); self.cant.insert(0, _fmt(linea.cantidad))
        self.costo = ctk.CTkEntry(master, width=85); self.costo.insert(0, _fmt(linea.costo_unit))
        # Si la linea cuenta como gasto del proveedor (para el Sustento) se decide
        # sola: las lineas del proveedor si, las propias (SCTR/EPPS) no.
        self.es_gasto = bool(linea.es_gasto)
        self.venta = ctk.CTkLabel(master, text="S/. 0.00", width=90, anchor="e")
        self.btn_del = ctk.CTkButton(master, text="✕", width=28, fg_color="#C0392B",
                                     hover_color="#922B21", command=self.eliminar)
        for w in (self.cant, self.costo):
            w.bind("<KeyRelease>", lambda e: app.recalcular())
        self.widgets = [self.desc, self.um, self.cant, self.costo, self.venta, self.btn_del]

    def grid(self, row):
        for c, w in enumerate(self.widgets):
            w.grid(row=row, column=c, padx=3, pady=2, sticky="w")

    def eliminar(self):
        self.app.eliminar_fila(self)

    def leer(self) -> Linea:
        return Linea(
            descripcion=self.desc.get().strip(),
            unidad=self.um.get().strip() or "UND",
            cantidad=_num(self.cant.get()),
            costo_unit=_num(self.costo.get()),
            es_gasto=self.es_gasto,
        )

    def destruir(self):
        for w in self.widgets:
            w.destroy()


def _num(s):
    try:
        return float(str(s).replace(",", "").strip() or 0)
    except ValueError:
        return 0.0


def _fmt(x):
    x = float(x or 0)
    return str(int(x)) if x == int(x) else f"{x:.2f}"


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.cfg = cargar_config()
        self.filas: list[FilaLinea] = []
        self.proveedores_comp: list[Proveedor] = []

        self.title("Cotizador FMI - Panorama Outsourcing")
        self.geometry("1000x760")
        self.minsize(900, 640)

        cont = ctk.CTkScrollableFrame(self)
        cont.pack(fill="both", expand=True, padx=10, pady=10)
        self.cont = cont

        self._encabezado(cont)
        self._datos_generales(cont)
        self._proveedor(cont)
        self._tabla_lineas(cont)
        self._margen_totales(cont)
        self._generar(cont)

        self._agregar_linea(Linea("", "UND", 1, 0, True))
        self.recalcular()

    # ---------------- Secciones ----------------
    def _titulo(self, master, txt):
        f = ctk.CTkFrame(master, fg_color="#1F3864")
        f.pack(fill="x", pady=(14, 4))
        ctk.CTkLabel(f, text=txt, font=ctk.CTkFont(size=14, weight="bold"),
                     text_color="white").pack(anchor="w", padx=10, pady=4)

    def _encabezado(self, m):
        f = ctk.CTkFrame(m, fg_color="transparent")
        f.pack(fill="x")
        ctk.CTkLabel(f, text="COTIZADOR FMI", font=ctk.CTkFont(size=22, weight="bold"),
                     text_color="#C55A11").pack(side="left", padx=4)
        ctk.CTkLabel(f, text="Panorama Outsourcing S.A.", font=ctk.CTkFont(size=13),
                     text_color="#1F3864").pack(side="left", padx=8)
        ctk.CTkButton(f, text="⚙ Configuracion", width=130, command=self.abrir_config).pack(side="right")
        ctk.CTkButton(f, text="🖼 Logo y firma", width=140, fg_color="#5B7DB1",
                      command=self.abrir_imagenes).pack(side="right", padx=6)
        self.lbl_logo = ctk.CTkLabel(f, text=self._estado_logo(), font=ctk.CTkFont(size=10),
                                     text_color="#555555")
        self.lbl_logo.pack(side="right", padx=6)

    def _campo(self, master, etiqueta, valor="", ancho=280, row=0, col=0, multilinea=False):
        cell = ctk.CTkFrame(master, fg_color="transparent")
        cell.grid(row=row, column=col, padx=8, pady=4, sticky="w")
        ctk.CTkLabel(cell, text=etiqueta, font=ctk.CTkFont(size=11, weight="bold")).pack(anchor="w")
        if multilinea:
            w = ctk.CTkTextbox(cell, width=ancho, height=52)
            w.insert("1.0", valor)
        else:
            w = ctk.CTkEntry(cell, width=ancho)
            w.insert(0, valor)
        w.pack(anchor="w")
        return w

    def _datos_generales(self, m):
        self._titulo(m, "1. Datos generales")
        g = ctk.CTkFrame(m, fg_color="transparent")
        g.pack(fill="x")
        hoy = dt.date.today()
        self.e_numero = self._campo(g, "N° Cotizacion", siguiente_correlativo(self.cfg), 200, 0, 0)
        self.e_fecha = self._campo(g, "Fecha", fecha_larga(hoy), 260, 0, 1)
        self.e_cliente = self._campo(g, "Cliente", "", 280, 0, 2)
        self.e_contacto = self._campo(g, "Contacto", "", 200, 1, 0)
        self.e_ref = self._campo(g, "N° Ticket / Solicitud", "Segun Requerimiento", 260, 1, 1)
        self.e_sede = self._campo(g, "Local / Sede", "", 280, 1, 2)
        g2 = ctk.CTkFrame(m, fg_color="transparent")
        g2.pack(fill="x")
        self.e_asunto = self._campo(g2, "Asunto / Descripcion del servicio", "", 760, 0, 0, multilinea=True)

    def _proveedor(self, m):
        self._titulo(m, "2. Proveedor principal (fuente del gasto)")
        g = ctk.CTkFrame(m, fg_color="transparent")
        g.pack(fill="x")
        self.e_prov_nom = self._campo(g, "Proveedor", "", 300, 0, 0)
        self.e_prov_ruc = self._campo(g, "RUC", "", 160, 0, 1)
        self.e_prov_num = self._campo(g, "N° Cotizacion proveedor", "", 220, 0, 2)
        bar = ctk.CTkFrame(m, fg_color="transparent")
        bar.pack(fill="x", pady=4)
        ctk.CTkButton(bar, text="📄  Cargar PDF de proveedor (extraer)", width=260,
                      fg_color="#C55A11", hover_color="#A0480E",
                      command=self.cargar_pdf).pack(side="left", padx=4)
        ctk.CTkButton(bar, text="＋ Proveedor al comparativo", width=210,
                      command=self.agregar_al_comparativo).pack(side="left", padx=4)
        self.lbl_comp = ctk.CTkLabel(bar, text="Comparativo: 0 proveedores")
        self.lbl_comp.pack(side="left", padx=10)

    def _tabla_lineas(self, m):
        self._titulo(m, "3. Lineas de la cotizacion")
        cab = ctk.CTkFrame(m, fg_color="#D9E1F2")
        cab.pack(fill="x")
        encabezados = [("Descripcion", 340), ("U.M.", 55), ("Cant.", 55),
                       ("Costo prov.", 85), ("Venta", 90), ("", 30)]
        for i, (t, w) in enumerate(encabezados):
            ctk.CTkLabel(cab, text=t, width=w, font=ctk.CTkFont(size=11, weight="bold")).grid(
                row=0, column=i, padx=3, pady=3, sticky="w")
        self.tabla = ctk.CTkFrame(m, fg_color="transparent")
        self.tabla.pack(fill="x")
        bar = ctk.CTkFrame(m, fg_color="transparent")
        bar.pack(fill="x", pady=4)
        ctk.CTkButton(bar, text="＋ Linea", width=110, command=lambda: self._agregar_linea(
            Linea("", "UND", 1, 0, True))).pack(side="left", padx=4)
        ctk.CTkButton(bar, text="＋ SCTR / EPPS / traslado", width=200, fg_color="#5B7DB1",
                      command=self.agregar_adicional).pack(side="left", padx=4)

    def _margen_totales(self, m):
        self._titulo(m, "4. Margen y totales")
        g = ctk.CTkFrame(m, fg_color="transparent")
        g.pack(fill="x")
        p = self.cfg["parametros"]
        cell = ctk.CTkFrame(g, fg_color="transparent"); cell.grid(row=0, column=0, padx=8, sticky="w")
        ctk.CTkLabel(cell, text="Margen %", font=ctk.CTkFont(size=11, weight="bold")).pack(anchor="w")
        self.e_margen = ctk.CTkEntry(cell, width=80)
        self.e_margen.insert(0, _fmt(p["margen_default"] * 100))
        self.e_margen.pack(); self.e_margen.bind("<KeyRelease>", lambda e: self.recalcular())
        cell2 = ctk.CTkFrame(g, fg_color="transparent"); cell2.grid(row=0, column=1, padx=8, sticky="w")
        ctk.CTkLabel(cell2, text="Modo de margen", font=ctk.CTkFont(size=11, weight="bold")).pack(anchor="w")
        self.opt_modo = ctk.CTkOptionMenu(cell2, width=220,
            values=["sobre venta (÷)  ej. 755/0.85", "sobre costo (×)  ej. 755×1.15"],
            command=lambda _: self.recalcular())
        self.opt_modo.pack()

        self.panel = ctk.CTkFrame(g, fg_color="#F2F2F2")
        self.panel.grid(row=0, column=2, padx=20, sticky="e")
        self.lbl_totales = ctk.CTkLabel(self.panel, text="", justify="left",
                                        font=ctk.CTkFont(size=13))
        self.lbl_totales.pack(padx=16, pady=8)

    def _generar(self, m):
        self._titulo(m, "5. Documentos a generar")
        g = ctk.CTkFrame(m, fg_color="transparent")
        g.pack(fill="x")
        self.chk = {
            "cotizacion_excel": ctk.CTkCheckBox(g, text="Cotizacion (Excel)"),
            "cotizacion_pdf": ctk.CTkCheckBox(g, text="Cotizacion (PDF)"),
            "sustento": ctk.CTkCheckBox(g, text="Sustento de Gasto (Excel+PDF)"),
            "comparativo": ctk.CTkCheckBox(g, text="Comparativo de proveedores"),
            "control": ctk.CTkCheckBox(g, text="Registrar en bitacora"),
        }
        for i, c in enumerate(self.chk.values()):
            c.grid(row=0, column=i, padx=8, pady=4, sticky="w")
            c.select()
        self.chk["comparativo"].deselect()

        bar = ctk.CTkFrame(m, fg_color="transparent")
        bar.pack(fill="x", pady=10)
        self.btn_generar = ctk.CTkButton(bar, text="✔  GENERAR DOCUMENTOS", height=44, width=280,
                      font=ctk.CTkFont(size=15, weight="bold"),
                      fg_color="#1F3864", hover_color="#152742",
                      command=self.generar)
        self.btn_generar.pack(side="left", padx=4)
        ctk.CTkButton(bar, text="📁 Abrir carpeta de salidas", width=200,
                      command=self.abrir_salidas).pack(side="left", padx=4)
        self.lbl_estado = ctk.CTkLabel(bar, text="", font=ctk.CTkFont(size=12),
                                       text_color="#1F3864")
        self.lbl_estado.pack(side="left", padx=12)

    # ---------------- Logica ----------------
    def _agregar_linea(self, linea: Linea):
        fila = FilaLinea(self.tabla, self, linea)
        self.filas.append(fila)
        self._reordenar()
        self.recalcular()

    def agregar_adicional(self):
        cat = self.cfg.get("lineas_adicionales_catalogo", [])
        if cat:
            a = cat[0]
            self._agregar_linea(Linea(a["descripcion"], a.get("unidad", "GLB"),
                                      a.get("cantidad", 1), a.get("costo_base", 0), es_gasto=False))
        else:
            self._agregar_linea(Linea("SCTR, EPPS y traslado de herramientas.", "GLB", 1, 0, False))

    def eliminar_fila(self, fila: FilaLinea):
        fila.destruir()
        self.filas.remove(fila)
        self._reordenar()
        self.recalcular()

    def _reordenar(self):
        for i, f in enumerate(self.filas):
            f.grid(i)

    def _modo(self):
        return "sobre_venta" if self.opt_modo.get().startswith("sobre venta") else "sobre_costo"

    def _construir_cotizacion(self, consumir=False) -> Cotizacion:
        margen = _num(self.e_margen.get()) / 100.0
        modo = self._modo()
        lineas = [f.leer() for f in self.filas if f.leer().descripcion]
        numero = self.e_numero.get().strip()
        if consumir:
            # Solo avanza contador si el numero coincide con el sugerido
            if numero == siguiente_correlativo(self.cfg):
                consumir_correlativo(self.cfg)
        prov = Proveedor(nombre=self.e_prov_nom.get().strip(),
                         ruc=self.e_prov_ruc.get().strip(),
                         numero_cotizacion=self.e_prov_num.get().strip(),
                         lineas=[l for l in lineas if l.es_gasto])
        txt_fecha = self.e_fecha.get().strip()
        cot = Cotizacion(
            numero=numero, fecha=txt_fecha, fecha_dt=parsear_fecha(txt_fecha),
            cliente=self.e_cliente.get().strip(), contacto=self.e_contacto.get().strip(),
            asunto=self.e_asunto.get("1.0", "end").strip(),
            referencia_ticket=self.e_ref.get().strip(), sede=self.e_sede.get().strip(),
            margen=margen, modo_margen=modo, igv=self.cfg["parametros"]["igv"],
            moneda=self.cfg["parametros"]["moneda"], proveedor_principal=prov, lineas=lineas,
            proveedores_comparados=self.proveedores_comp,
        )
        return cot

    def recalcular(self):
        try:
            cot = self._construir_cotizacion()
        except Exception:
            return
        m = cot.moneda
        for f in self.filas:
            vu = f.leer().venta_unit(cot.margen, cot.modo_margen)
            f.venta.configure(text=f"{m} {vu:,.2f}")
        txt = (f"Subtotal:   {m} {cot.subtotal():,.2f}\n"
               f"IGV 18%:   {m} {cot.igv_monto():,.2f}\n"
               f"TOTAL:      {m} {cot.total():,.2f}\n"
               f"Utilidad:    {m} {cot.utilidad():,.2f}  ({cot.utilidad_pct()*100:.1f}%)")
        self.lbl_totales.configure(text=txt)

    # ---------------- PDF proveedor ----------------
    def cargar_pdf(self):
        ruta = filedialog.askopenfilename(title="Selecciona el PDF del proveedor",
                                          filetypes=[("PDF", "*.pdf")])
        if not ruta:
            return
        res = extraer(ruta)
        if res.ruc and not self.e_prov_ruc.get().strip():
            self.e_prov_ruc.delete(0, "end"); self.e_prov_ruc.insert(0, res.ruc)
        if res.numero_cotizacion and not self.e_prov_num.get().strip():
            self.e_prov_num.delete(0, "end"); self.e_prov_num.insert(0, res.numero_cotizacion)
        if res.lineas:
            # limpia filas vacias
            for f in list(self.filas):
                if not f.leer().descripcion:
                    self.eliminar_fila(f)
            for le in res.lineas:
                self._agregar_linea(Linea(le.descripcion, le.unidad, le.cantidad, le.costo_unit, True))
        messagebox.showinfo("Extraccion de PDF", res.mensaje +
                            "\n\nRevisa y corrige las lineas antes de generar.")
        self.recalcular()

    def agregar_al_comparativo(self):
        cot = self._construir_cotizacion()
        prov = Proveedor(nombre=self.e_prov_nom.get().strip() or f"Proveedor {len(self.proveedores_comp)+1}",
                         numero_cotizacion=self.e_prov_num.get().strip(),
                         lineas=[l for l in cot.lineas if l.es_gasto])
        if not prov.lineas:
            messagebox.showwarning("Comparativo", "No hay lineas de gasto para agregar.")
            return
        self.proveedores_comp.append(prov)
        self.lbl_comp.configure(text=f"Comparativo: {len(self.proveedores_comp)} proveedores")
        self.chk["comparativo"].select()
        messagebox.showinfo("Comparativo",
                            f"Proveedor '{prov.nombre}' agregado al comparativo.\n"
                            "Cambia el proveedor/precios y agrega el siguiente.")

    # ---------------- Generar ----------------
    def generar(self):
        cot = self._construir_cotizacion()
        if not cot.cliente or not cot.asunto:
            messagebox.showwarning("Faltan datos", "Completa al menos Cliente y Asunto.")
            return
        if not [l for l in cot.lineas if l.descripcion]:
            messagebox.showwarning("Faltan datos", "Agrega al menos una linea con descripcion.")
            return
        opciones = {k: bool(c.get()) for k, c in self.chk.items()}
        # Corre en segundo plano para no congelar la ventana (Excel puede tardar unos segundos)
        self.btn_generar.configure(state="disabled")
        self.lbl_estado.configure(text="⏳ Generando documentos con Excel...")
        import threading
        threading.Thread(target=self._generar_worker, args=(opciones,), daemon=True).start()

    def _generar_worker(self, opciones):
        try:
            cot = self._construir_cotizacion(consumir=True)
            res = generar_todo(cot, self.cfg, opciones)
            err = None
        except Exception as e:
            res, err = None, f"{e}\n\n{traceback.format_exc()}"
        self.after(0, lambda: self._generar_listo(res, err))

    def _generar_listo(self, res, err):
        self.btn_generar.configure(state="normal")
        self.lbl_estado.configure(text="")
        if err:
            messagebox.showerror("Error al generar", err)
            return
        carpeta = os.path.dirname(next(iter(res.values())))
        self.e_numero.delete(0, "end"); self.e_numero.insert(0, siguiente_correlativo(self.cfg))
        detalle = "\n".join(f"• {k}" for k in res)
        if messagebox.askyesno("Documentos generados",
                               f"Se generaron:\n{detalle}\n\n¿Abrir la carpeta?"):
            self._abrir(carpeta)

    def abrir_salidas(self):
        self._abrir(cfgmod.carpeta_salidas(self.cfg))

    def _abrir(self, ruta):
        try:
            if sys.platform.startswith("win"):
                os.startfile(ruta)
            elif sys.platform == "darwin":
                subprocess.run(["open", ruta])
            else:
                subprocess.run(["xdg-open", ruta])
        except Exception as e:
            messagebox.showinfo("Carpeta", f"Salidas en:\n{ruta}\n({e})")

    # ---------------- Logo y firma ----------------
    def _estado_logo(self):
        l = "propio" if es_personalizado(self.cfg, "logo") else "original"
        f = "propia" if es_personalizado(self.cfg, "firma") else "original"
        return f"Logo: {l}  |  Firma: {f}"

    def abrir_imagenes(self):
        VentanaImagenes(self, self.cfg)

    # ---------------- Config ----------------
    def abrir_config(self):
        VentanaConfig(self, self.cfg)


class VentanaImagenes(ctk.CTkToplevel):
    """Permite cargar el logo de la empresa y la imagen de la firma."""

    def __init__(self, app: "App", cfg: dict):
        super().__init__(app)
        self.app = app
        self.cfg = cfg
        self.title("Logo y firma")
        self.geometry("560x520")
        self.grab_set()

        ctk.CTkLabel(self, text="Imagenes de los documentos",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(14, 2))
        ctk.CTkLabel(self, text="Se aplican a la Cotizacion y al Sustento de Gasto.",
                     font=ctk.CTkFont(size=11), text_color="#666666").pack(pady=(0, 8))

        self._previews = {}   # evita que el recolector borre las imagenes
        self.bloques = {}
        self._bloque("logo", "Logo de empresa", "Aparece en la cabecera del documento.")
        self._bloque("firma", "Logo de firma", "Aparece sobre el nombre, al pie del documento.")

        ctk.CTkButton(self, text="Cerrar", width=140, command=self.destroy).pack(pady=10)
        self._refrescar()

    def _bloque(self, que: str, titulo: str, ayuda: str):
        marco = ctk.CTkFrame(self)
        marco.pack(fill="x", padx=16, pady=8)
        ctk.CTkLabel(marco, text=titulo,
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=12, pady=(8, 0))
        ctk.CTkLabel(marco, text=ayuda, font=ctk.CTkFont(size=11),
                     text_color="#666666").pack(anchor="w", padx=12)

        fila = ctk.CTkFrame(marco, fg_color="transparent")
        fila.pack(fill="x", padx=12, pady=8)
        vista = ctk.CTkLabel(fila, text="(sin imagen)", width=190, height=62,
                             fg_color="#F2F2F2", corner_radius=6)
        vista.pack(side="left")
        botones = ctk.CTkFrame(fila, fg_color="transparent")
        botones.pack(side="left", padx=12)
        ctk.CTkButton(botones, text="📂 Cargar imagen", width=170,
                      command=lambda: self._cargar(que)).pack(pady=2)
        ctk.CTkButton(botones, text="↩ Usar la original", width=170, fg_color="#7A7A7A",
                      hover_color="#5E5E5E",
                      command=lambda: self._restaurar(que)).pack(pady=2)
        estado = ctk.CTkLabel(marco, text="", font=ctk.CTkFont(size=11))
        estado.pack(anchor="w", padx=12, pady=(0, 8))
        self.bloques[que] = {"vista": vista, "estado": estado}

    def _ruta(self, que: str):
        return ruta_logo(self.cfg) if que == "logo" else ruta_firma(self.cfg)

    def _refrescar(self):
        for que, w in self.bloques.items():
            ruta = self._ruta(que)
            propio = es_personalizado(self.cfg, que)
            w["estado"].configure(
                text=("● Usando tu imagen" if propio else "○ Usando la imagen original de la plantilla"),
                text_color="#1F7A1F" if propio else "#666666")
            img = self._miniatura(ruta)
            if img:
                self._previews[que] = img
                w["vista"].configure(image=img, text="")
            else:
                w["vista"].configure(image=None, text="(sin imagen)")
        self.app.lbl_logo.configure(text=self.app._estado_logo())

    def _miniatura(self, ruta):
        if not ruta or not os.path.exists(ruta):
            return None
        try:
            from PIL import Image
            im = Image.open(ruta)
            im.thumbnail((180, 56))
            return ctk.CTkImage(light_image=im, dark_image=im, size=im.size)
        except Exception:
            return None

    def _cargar(self, que: str):
        ruta = filedialog.askopenfilename(
            title=f"Selecciona la imagen ({'logo' if que == 'logo' else 'firma'})",
            filetypes=[("Imagenes", "*.png *.jpg *.jpeg *.bmp *.gif"), ("Todos", "*.*")])
        if not ruta:
            return
        try:
            (guardar_logo if que == "logo" else guardar_firma)(self.cfg, ruta)
        except Exception as e:
            messagebox.showerror("Imagen", f"No se pudo cargar:\n{e}", parent=self)
            return
        self._refrescar()

    def _restaurar(self, que: str):
        (restaurar_logo if que == "logo" else restaurar_firma)(self.cfg)
        self._refrescar()


class VentanaConfig(ctk.CTkToplevel):
    def __init__(self, app: App, cfg: dict):
        super().__init__(app)
        self.app = app; self.cfg = cfg
        self.title("Configuracion")
        self.geometry("560x560")
        self.grab_set()
        f = ctk.CTkScrollableFrame(self); f.pack(fill="both", expand=True, padx=12, pady=12)

        p = cfg["parametros"]; emp = cfg["empresa"]
        self.campos = {}
        def add(seccion, clave, etiqueta, valor):
            row = ctk.CTkFrame(f, fg_color="transparent"); row.pack(fill="x", pady=3)
            ctk.CTkLabel(row, text=etiqueta, width=220, anchor="w").pack(side="left")
            e = ctk.CTkEntry(row, width=280); e.insert(0, str(valor)); e.pack(side="left")
            self.campos[(seccion, clave)] = e

        ctk.CTkLabel(f, text="Parametros", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", pady=(4, 6))
        add("parametros", "margen_default", "Margen por defecto (0-1)", p["margen_default"])
        add("parametros", "igv", "IGV (0-1)", p["igv"])
        add("parametros", "prefijo_correlativo", "Prefijo correlativo", p["prefijo_correlativo"])
        add("parametros", "anio", "Año (2 digitos)", p["anio"])
        add("parametros", "proximo_numero", "Proximo numero", p["proximo_numero"])
        add("parametros", "moneda", "Moneda", p["moneda"])
        ctk.CTkLabel(f, text="Empresa / Firma", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", pady=(12, 6))
        add("empresa", "razon_social", "Razon social", emp["razon_social"])
        add("empresa", "firma_nombre", "Nombre firma", emp["firma_nombre"])
        add("empresa", "firma_cargo", "Cargo firma", emp["firma_cargo"])

        ctk.CTkButton(self, text="Guardar", command=self.guardar, height=40).pack(pady=8)

    def guardar(self):
        for (sec, clave), e in self.campos.items():
            val = e.get().strip()
            actual = self.cfg[sec][clave]
            if isinstance(actual, bool):
                val = val.lower() in ("true", "1", "si")
            elif isinstance(actual, int):
                try: val = int(float(val))
                except ValueError: pass
            elif isinstance(actual, float):
                try: val = float(val)
                except ValueError: pass
            self.cfg[sec][clave] = val
        guardar_config(self.cfg)
        self.app.e_numero.delete(0, "end")
        self.app.e_numero.insert(0, siguiente_correlativo(self.cfg))
        self.app.e_margen.delete(0, "end")
        self.app.e_margen.insert(0, _fmt(self.cfg["parametros"]["margen_default"] * 100))
        self.app.recalcular()
        messagebox.showinfo("Configuracion", "Cambios guardados.")
        self.destroy()


if __name__ == "__main__":
    try:
        App().mainloop()
    except Exception:
        traceback.print_exc()
        input("Error. Presiona Enter para salir...")
