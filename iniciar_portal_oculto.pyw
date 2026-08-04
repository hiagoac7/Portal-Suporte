# -*- coding: utf-8 -*-
# Lancador OCULTO do Portal de Suporte para a inicializacao com o Windows.
#
# Por que existe: a sincronizacao (OneDrive/backup) costuma colar BYTES NULOS
# no fim do app.py, e isso faz o Python recusar o arquivo com o erro
# "source code string cannot contain null bytes". Este lancador SANEIA o
# app.py (remove os bytes nulos) ANTES de subir o Streamlit -- assim o portal
# nunca falha ao iniciar junto com o Windows.
#
# Roda com pythonw.exe (sem console), entao nenhuma janela aparece.

import os
import sys
import runpy

BASE = os.path.dirname(os.path.abspath(__file__))
APP = os.path.join(BASE, "app.py")

# 1) Saneamento anti-corrupcao (remove bytes nulos / lixo do final do arquivo)
try:
    with open(APP, "rb") as _f:
        _dados = _f.read()
    _limpo = _dados.replace(b"\x00", b"").rstrip() + b"\n"
    if _limpo != _dados:
        with open(APP, "wb") as _f:
            _f.write(_limpo)
except Exception:
    # Se o saneamento falhar, ainda assim tentamos iniciar o portal.
    pass

# 2) Inicia o Streamlit NESTE mesmo processo (equivale a "python -m streamlit run app.py")
os.chdir(BASE)
sys.argv = [
    "streamlit", "run", APP,
    "--server.port", "8501",
    "--server.address", "0.0.0.0",
    "--server.headless", "true",
]
runpy.run_module("streamlit", run_name="__main__")
