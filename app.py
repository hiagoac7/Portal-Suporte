# -*- coding: utf-8 -*-
"""REDIRECIONADOR — compatibilidade com o caminho antigo.

Na reorganização de 09/2026 o portal passou para PortalSuporte\\app\\app.py.
Atalhos, Tarefas Agendadas e configurações do VS Code criados ANTES disso
continuam chamando este caminho, e quebravam com:

    [Errno 2] No such file or directory: '...\\PortalSuporte\\app.py'

Este arquivo existe só para esses chamadores antigos: ele executa o app de
verdade, informando o caminho correto, para que `__file__` (e portanto as
pastas app\\ e dados\\) resolvam certo lá dentro.

Pode ser apagado depois de rodar `windows\\INSTALAR INICIO AUTOMATICO.bat`,
que reaponta o início automático para o lugar novo.
"""
import os
import runpy
import sys

_RAIZ = os.path.dirname(os.path.abspath(__file__))
_REAL = os.path.join(_RAIZ, "app", "app.py")

if not os.path.exists(_REAL):
    raise SystemExit("Nao encontrei o portal em %s" % _REAL)

if os.path.dirname(_REAL) not in sys.path:
    sys.path.insert(0, os.path.dirname(_REAL))

runpy.run_path(_REAL, run_name="__main__")
