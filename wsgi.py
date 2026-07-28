import os
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Asegura que tanto la carpeta actual como la subcarpeta 'CotizadorFMI' esten en sys.path
for p in [BASE_DIR, os.path.join(BASE_DIR, "CotizadorFMI")]:
    if os.path.exists(p) and p not in sys.path:
        sys.path.insert(0, p)

try:
    from webapp.app import app
except ModuleNotFoundError:
    try:
        from CotizadorFMI.webapp.app import app
    except ModuleNotFoundError:
        import app as _app_module
        app = _app_module.app

if __name__ == "__main__":
    app.run()
