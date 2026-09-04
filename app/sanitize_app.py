# -*- coding: utf-8 -*-

import os
import shutil
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
ALVO = os.path.join(BASE, "app.py")
BACKUP_DIR = os.path.join(os.path.dirname(BASE), "dados", "backups")


def sanear(caminho):
    """Remove lixo do final do arquivo. Retorna True se alterou algo.

    Trata os dois tipos de corrupcao ja observados no fim do app.py:
    (1) bytes nulos no final; (2) caractere UTF-8 cortado pela metade no final,
    que faz o Streamlit falhar com 'unexpected end of data'.
    """
    with open(caminho, "rb") as f:
        original = f.read()

    limpo = original.replace(b"\x00", b"")

    while limpo:
        try:
            limpo.decode("utf-8")
            break
        except UnicodeDecodeError:
            limpo = limpo[:-1]

    limpo = limpo.rstrip() + b"\n"

    if limpo != original:
        with open(caminho, "wb") as f:
            f.write(limpo)
        return True
    return False


def _compila(caminho):
    """True se o arquivo le como UTF-8 e compila sem erro de sintaxe."""
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            fonte = f.read()
        compile(fonte, caminho, "exec")
        return True
    except Exception:
        return False


def _backups_validos_mais_recentes():
    """Backups que parecem fonte do app (app*.py / app.py.bak*), do mais
    recente para o mais antigo."""
    if not os.path.isdir(BACKUP_DIR):
        return []
    candidatos = []
    for nome in os.listdir(BACKUP_DIR):
        n = nome.lower()
        if not n.startswith("app"):
            continue
        if not (n.endswith(".py") or ".bak" in n):
            continue
        if "quebrado" in n or "corrupt" in n or "broken" in n:
            continue
        caminho = os.path.join(BACKUP_DIR, nome)
        if os.path.isfile(caminho):
            candidatos.append((os.path.getmtime(caminho), caminho))
    candidatos.sort(reverse=True)
    return [c for _, c in candidatos]


def garantir_integridade(caminho):
    """Garante que o app.py compila. Se nao, restaura o backup valido mais
    recente. Retorna string de status legivel."""
    if _compila(caminho):
        return "OK"

    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    try:
        shutil.copy2(caminho, os.path.join(BACKUP_DIR, "app_QUEBRADO_%s.py" % ts))
    except Exception:
        pass

    for backup in _backups_validos_mais_recentes():
        if _compila(backup):
            try:
                shutil.copy2(backup, caminho)
            except Exception as e:
                return "FALHA ao restaurar %s: %s" % (os.path.basename(backup), e)
            if _compila(caminho):
                return "RESTAURADO de %s" % os.path.basename(backup)
    return "CRITICO: nenhum backup valido encontrado para restaurar"


if __name__ == "__main__":
    try:
        if sanear(ALVO):
            print("[sanitize] app.py saneado: lixo removido do final.")
        else:
            print("[sanitize] app.py OK: nada a limpar.")

        status = garantir_integridade(ALVO)
        if status == "OK":
            print("[sanitize] integridade OK: app.py compila normalmente.")
        elif status.startswith("RESTAURADO"):
            print("[sanitize] *** app.py estava QUEBRADO -> %s. "
                  "O conteudo quebrado foi salvo em backups/app_QUEBRADO_*.py" % status)
        else:
            print("[sanitize] *** ATENCAO: %s" % status)
    except FileNotFoundError:
        print("[sanitize] app.py nao encontrado em", ALVO)
    except Exception as e:
        print("[sanitize] falha ao sanear/validar app.py:", e)
