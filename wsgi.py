import os
import sys

# Agrega la carpeta raiz al PATH de Python
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from webapp.app import app

if __name__ == "__main__":
    app.run()
