"""Carga y guarda la configuracion editable de la app (datos/config.json)."""
from __future__ import annotations
import json
import os
import sys

def _base_dir() -> str:
    # Compatible con PyInstaller (onefile) y ejecucion normal
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

BASE_DIR = _base_dir()
DATOS_DIR = os.path.join(BASE_DIR, "datos")
CONFIG_PATH = os.path.join(DATOS_DIR, "config.json")


def _asegurar_config() -> None:
    """En el .exe (onefile) copia la config incluida junto al ejecutable si falta."""
    if os.path.exists(CONFIG_PATH):
        return
    os.makedirs(DATOS_DIR, exist_ok=True)
    bundle = getattr(sys, "_MEIPASS", None)
    origen = os.path.join(bundle, "datos", "config.json") if bundle else None
    if origen and os.path.exists(origen):
        import shutil
        shutil.copyfile(origen, CONFIG_PATH)


def cargar_config() -> dict:
    _asegurar_config()
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def guardar_config(cfg: dict) -> None:
    os.makedirs(DATOS_DIR, exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def siguiente_correlativo(cfg: dict) -> str:
    """Devuelve el correlativo sugerido sin consumirlo, ej 'FMI - 0246-26'."""
    p = cfg["parametros"]
    return f"{p['prefijo_correlativo']} - {int(p['proximo_numero']):04d}-{p['anio']}"


def consumir_correlativo(cfg: dict) -> str:
    """Devuelve el correlativo actual y avanza el contador (persistiendo)."""
    numero = siguiente_correlativo(cfg)
    cfg["parametros"]["proximo_numero"] = int(cfg["parametros"]["proximo_numero"]) + 1
    guardar_config(cfg)
    return numero


def _ruta_imagen(cfg: dict, clave: str) -> str | None:
    """Resuelve la ruta absoluta de una imagen de config (logo_path / firma_path)."""
    rel = (cfg.get("empresa", {}).get(clave) or "").strip()
    if not rel:
        return None
    if os.path.isabs(rel):
        return rel if os.path.exists(rel) else None
    # busca junto al ejecutable/proyecto y, en el .exe, en el bundle
    candidatos = [os.path.join(BASE_DIR, rel)]
    bundle = getattr(sys, "_MEIPASS", None)
    if bundle:
        candidatos.append(os.path.join(bundle, rel))
    for c in candidatos:
        if os.path.exists(c):
            return c
    return None


def ruta_logo(cfg: dict) -> str | None:
    """Ruta absoluta del logo de empresa, o None."""
    return _ruta_imagen(cfg, "logo_path")


def ruta_firma(cfg: dict) -> str | None:
    """Ruta absoluta de la imagen de firma, o None."""
    return _ruta_imagen(cfg, "firma_path")


def _guardar_imagen(cfg: dict, ruta_origen: str, nombre: str,
                    clave_path: str, clave_flag: str) -> str:
    """Copia la imagen elegida a assets/<nombre>.<ext>, marca personalizada y guarda."""
    import shutil
    assets = os.path.join(BASE_DIR, "assets")
    os.makedirs(assets, exist_ok=True)
    ext = os.path.splitext(ruta_origen)[1].lower() or ".png"
    destino = os.path.join(assets, nombre + ext)
    shutil.copyfile(ruta_origen, destino)
    cfg["empresa"][clave_path] = os.path.join("assets", nombre + ext).replace("\\", "/")
    cfg["empresa"][clave_flag] = True
    guardar_config(cfg)
    return destino


def guardar_logo(cfg: dict, ruta_origen: str) -> str:
    """Carga un logo de empresa propio (se usara en Cotizacion y Sustento)."""
    return _guardar_imagen(cfg, ruta_origen, "logo_usuario",
                           "logo_path", "logo_personalizado")


def guardar_firma(cfg: dict, ruta_origen: str) -> str:
    """Carga una imagen de firma propia (reemplaza la firma de las plantillas)."""
    return _guardar_imagen(cfg, ruta_origen, "firma_usuario",
                           "firma_path", "firma_personalizada")


def restaurar_logo(cfg: dict) -> None:
    """Vuelve al logo original de las plantillas."""
    cfg["empresa"]["logo_path"] = "assets/logo.png"
    cfg["empresa"]["logo_personalizado"] = False
    guardar_config(cfg)


def restaurar_firma(cfg: dict) -> None:
    """Vuelve a la firma original de las plantillas."""
    cfg["empresa"]["firma_path"] = "assets/firma.png"
    cfg["empresa"]["firma_personalizada"] = False
    guardar_config(cfg)


def es_personalizado(cfg: dict, que: str) -> bool:
    """que: 'logo' | 'firma'."""
    clave = "logo_personalizado" if que == "logo" else "firma_personalizada"
    return bool(cfg.get("empresa", {}).get(clave, False))


def carpeta_salidas(cfg: dict) -> str:
    carpeta = cfg.get("carpeta_salidas", "salidas")
    if not os.path.isabs(carpeta):
        carpeta = os.path.join(BASE_DIR, carpeta)
    os.makedirs(carpeta, exist_ok=True)
    return carpeta
