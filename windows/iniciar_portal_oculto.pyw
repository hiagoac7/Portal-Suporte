# -*- coding: utf-8 -*-

import os
import sys
import runpy

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DIR_APP = os.path.join(RAIZ, "app")
APP = os.path.join(DIR_APP, "app.py")

try:
    with open(APP, "rb") as _f:
        _dados = _f.read()
    _limpo = _dados.replace(b"\x00", b"").rstrip() + b"\n"
    if _limpo != _dados:
        with open(APP, "wb") as _f:
            _f.write(_limpo)
except Exception:
    pass

try:
    sys.path.insert(0, DIR_APP)
    from marca_streamlit import aplicar as _marcar
    _marcar()
except Exception:
    pass

os.chdir(RAIZ)
sys.argv = [
    "streamlit", "run", APP,
    "--server.port", "8501",
    "--server.address", "0.0.0.0",
    "--server.headless", "true",
]
runpy.run_module("streamlit", run_name="__main__")
