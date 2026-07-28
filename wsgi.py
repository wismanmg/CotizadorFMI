import os
import sys
import importlib.util

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

app_path = os.path.join(BASE_DIR, "webapp", "app.py")

if os.path.exists(app_path):
    # Carga directamente el archivo webapp/app.py por ruta de archivo
    spec = importlib.util.spec_from_file_location("webapp_app", app_path)
    webapp_module = importlib.util.module_from_spec(spec)
    sys.modules["webapp_app"] = webapp_module
    spec.loader.exec_module(webapp_module)
    app = webapp_module.app
else:
    from webapp.app import app

if __name__ == "__main__":
    app.run()
