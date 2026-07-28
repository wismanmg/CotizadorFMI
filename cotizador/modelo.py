"""Modelo de datos y logica de calculo de la Cotizacion FMI.

Reglas de negocio (validadas con el caso 0245-26):
  - Precio de venta de una linea = costo con margen aplicado.
      modo "sobre_venta":  venta = costo / (1 - margen)      (margen bruto)
      modo "sobre_costo":  venta = costo * (1 + margen)      (markup)
    Ej: 755 / (1 - 0.15) = 888.24   -> como en la cotizacion real.
  - Lineas de proveedor: SI cuentan como GASTO en el sustento.
  - Lineas adicionales (SCTR/EPPS/traslado): venta con margen, NO cuentan como gasto.
  - IGV 18% sobre el costo directo (subtotal de venta).
"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import List
import unicodedata


def redondear(x: float, dec: int = 2) -> float:
    return round(float(x) + 1e-9, dec)


def precio_venta(costo: float, margen: float, modo: str) -> float:
    """Aplica el margen a un costo y devuelve el precio de venta unitario."""
    costo = float(costo or 0)
    margen = float(margen or 0)
    if modo == "sobre_venta":
        if margen >= 1:
            margen = 0.0
        return redondear(costo / (1 - margen))
    # sobre_costo (markup)
    return redondear(costo * (1 + margen))


@dataclass
class Linea:
    """Una linea de la cotizacion (item)."""
    descripcion: str = ""
    unidad: str = "UND"
    cantidad: float = 1.0
    costo_unit: float = 0.0          # costo base (proveedor o base de la linea adicional)
    es_gasto: bool = True            # True = linea de proveedor (cuenta como gasto)
    margen: float | None = None      # None => usa el margen global de la cotizacion

    def venta_unit(self, margen_global: float, modo: str) -> float:
        m = self.margen if self.margen is not None else margen_global
        return precio_venta(self.costo_unit, m, modo)

    def venta_parcial(self, margen_global: float, modo: str) -> float:
        return redondear(self.venta_unit(margen_global, modo) * float(self.cantidad or 0))

    def costo_parcial(self) -> float:
        return redondear(float(self.costo_unit or 0) * float(self.cantidad or 0))


@dataclass
class Proveedor:
    nombre: str = ""
    ruc: str = ""
    numero_cotizacion: str = ""      # N° de cotizacion segun el proveedor
    lineas: List[Linea] = field(default_factory=list)


@dataclass
class Cotizacion:
    numero: str = ""                 # ej. "FMI - 0245-26"
    fecha: str = ""                  # texto ya formateado, ej "jueves, 9 de Julio de 2026"
    fecha_dt: object | None = None   # date real (Excel la formatea y calcula el mes)
    cliente: str = ""
    contacto: str = ""
    asunto: str = ""
    referencia_ticket: str = "Segun Requerimiento"
    sede: str = ""                   # local / sede (para el sustento)
    margen: float = 0.15
    modo_margen: str = "sobre_venta"
    igv: float = 0.18
    moneda: str = "S/."
    proveedor_principal: Proveedor = field(default_factory=Proveedor)
    lineas: List[Linea] = field(default_factory=list)   # todas las lineas de VENTA
    proveedores_comparados: List[Proveedor] = field(default_factory=list)

    # ---- calculos de venta ----
    def subtotal(self) -> float:
        return redondear(sum(l.venta_parcial(self.margen, self.modo_margen) for l in self.lineas))

    def igv_monto(self) -> float:
        return redondear(self.subtotal() * self.igv)

    def total(self) -> float:
        return redondear(self.subtotal() + self.igv_monto())

    # ---- calculos de gasto (proveedor) ----
    def costo_gasto(self) -> float:
        """Suma de costos que son gasto real de proveedor (sin IGV)."""
        return redondear(sum(l.costo_parcial() for l in self.lineas if l.es_gasto))

    def gasto_con_igv(self) -> float:
        return redondear(self.costo_gasto() * (1 + self.igv))

    # ---- utilidad ----
    def utilidad(self) -> float:
        return redondear(self.subtotal() - self.costo_gasto())

    def utilidad_pct(self) -> float:
        s = self.subtotal()
        return redondear(self.utilidad() / s, 4) if s else 0.0


def quitar_acentos(texto: str) -> str:
    if not texto:
        return ""
    nf = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in nf if not unicodedata.combining(c))
