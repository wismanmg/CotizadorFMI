import os
import sys
import importlib.util

# Garantiza que la carpeta raiz este en sys.path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

# Carga la aplicacion web Flask desde webapp/app.py
webapp_app_path = os.path.join(BASE_DIR, "webapp", "app.py")
spec = importlib.util.spec_from_file_location("flask_webapp", webapp_app_path)
webapp_module = importlib.util.module_from_spec(spec)
sys.modules["flask_webapp"] = webapp_module
spec.loader.exec_module(webapp_module)

# Expone el objeto app de Flask para Gunicorn
app = webapp_module.app

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port)
