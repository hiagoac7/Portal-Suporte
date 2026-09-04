# -*- coding: utf-8 -*-
"""Aplica a marca TLM no HTML que o Streamlit serve.

POR QUE ISTO EXISTE
-------------------
O `page_title` do `st.set_page_config` só chega ao navegador DEPOIS que a
sessão WebSocket abre — cerca de 2 segundos. Até lá o Streamlit serve
`<title>Streamlit</title>`, e o Chrome preenche esse intervalo com o último
título que ele conhece daquela URL. Resultado: a aba pisca o nome antigo
("IT ServiceDesk Pro") antes de virar "TLM Service Desk".

A correção é marcar o próprio `index.html` do Streamlit, para que o primeiro
byte servido já venha com o título e o ícone certos. Não há mais intervalo.

O patch é idempotente e se reaplica sozinho: os lançadores chamam este script
antes de subir o servidor, então um `pip install -U streamlit` (que reescreve
o index.html original) é corrigido na próxima inicialização.

Uso:  python marca_streamlit.py
"""
import os
import re
import shutil

TITULO = "TLM Service Desk"
BASE = os.path.dirname(os.path.abspath(__file__))
ICONE = os.path.join(BASE, "tlm_icon.png")



ANTI_SALTO = """<!--tlm:inicio-->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<!-- Favicon com IMPRESSAO DIGITAL na URL. O Chrome guarda favicon em cache
     de forma muito agressiva: nao respeita Ctrl+Shift+R nem troca de
     arquivo com o mesmo nome. Com ?v=<hash do PNG> a URL muda quando o
     icone muda, e o navegador e obrigado a buscar de novo. -->
<link rel="icon" type="image/png" href="./favicon.png?v=__TLM_HASH__">
<link rel="shortcut icon" type="image/png" href="./favicon.png?v=__TLM_HASH__">
<!-- Só preconnect: a folha das fontes continua vindo pelo app. Um
     <link rel=stylesheet> aqui BLOQUEIA a renderização — se o Google Fonts
     estiver lento ou bloqueado na rede, o portal inteiro fica esperando. -->
<style id="tlm-anti-salto">
/* fundo já na cor certa: evita o lampejo branco antes da primeira pintura */
html, body { background: #0d0d0c; color: #f5f3ee; }
@media (prefers-color-scheme: light) {
  html, body { background: #f6f4f0; color: #14130f; }
}
@media (min-width: 768px) {
  /* Enquanto a sidebar não existe: reserva os 300px dela e já usa o layout
     largo. Quando ela monta, o seletor deixa de casar e nada se move. */
  [data-testid="stAppViewContainer"]:not(:has(> [data-testid="stSidebar"])) > div {
      margin-left: 300px;
  }
  [data-testid="stAppViewContainer"]:not(:has(> [data-testid="stSidebar"]))
      [data-testid="stMainBlockContainer"] {
      max-width: none;
  }
}
</style>
<!--tlm:fim-->"""

def _static_do_streamlit():
    import streamlit
    return os.path.join(os.path.dirname(os.path.abspath(streamlit.__file__)), "static")


def aplicar(verboso=False):
    def diz(msg):
        if verboso:
            print(msg)

    try:
        static = _static_do_streamlit()
    except Exception as e:
        diz("[!] Streamlit nao encontrado: %s" % e)
        return False

    index = os.path.join(static, "index.html")
    if not os.path.exists(index):
        diz("[!] index.html nao encontrado em %s" % static)
        return False

    mudou = False

    try:
        with open(index, encoding="utf-8") as f:
            html = f.read()

        try:
            import streamlit as _st_v
            _versao = getattr(_st_v, "__version__", "x")
        except Exception:
            _versao = "x"
        orig = "%s.orig-%s" % (index, _versao)
        if not os.path.exists(orig) and "<!--tlm:inicio-->" not in html:
            shutil.copy2(index, orig)
            diz("[i] Backup do index.html original (%s) criado." % _versao)

        novo = re.sub(r"<title>[^<]*</title>", "<title>%s</title>" % TITULO, html, count=1)

        novo = re.sub(r"<!--tlm:inicio-->.*?<!--tlm:fim-->", "", novo, flags=re.S)

        try:
            import hashlib
            with open(ICONE, "rb") as _fi:
                _hash = hashlib.md5(_fi.read()).hexdigest()[:10]
        except Exception:
            _hash = "0"
        bloco = ANTI_SALTO.replace("__TLM_HASH__", _hash)

        novo = novo.replace("</head>", bloco + "\n</head>", 1)

        if novo != html:
            with open(index, "w", encoding="utf-8") as f:
                f.write(novo)
            mudou = True
            diz("[OK] Titulo do HTML: %s" % TITULO)
        else:
            diz("[=] Titulo ja estava aplicado.")
    except Exception as e:
        diz("[!] Falha ao ajustar o titulo: %s" % e)

    try:
        if os.path.exists(ICONE):
            destino = os.path.join(static, "favicon.png")
            backup = destino + ".orig"
            if os.path.exists(destino) and not os.path.exists(backup):
                shutil.copy2(destino, backup)

            try:
                from PIL import Image
                import io as _io
                img = Image.open(ICONE).convert("RGBA").resize((192, 192), Image.LANCZOS)
                buf = _io.BytesIO()
                img.save(buf, "PNG")
                dados_novos = buf.getvalue()
            except Exception:
                with open(ICONE, "rb") as f:
                    dados_novos = f.read()

            atual = b""
            if os.path.exists(destino):
                with open(destino, "rb") as f:
                    atual = f.read()

            if atual != dados_novos:
                with open(destino, "wb") as f:
                    f.write(dados_novos)
                mudou = True
                diz("[OK] Favicon servido trocado pelo selo TLM.")
            else:
                diz("[=] Favicon ja estava aplicado.")
        else:
            diz("[!] %s nao encontrado — favicon nao alterado." % ICONE)
    except Exception as e:
        diz("[!] Falha ao trocar o favicon: %s" % e)

    try:
        import json
        mf = os.path.join(static, "manifest.json")
        if os.path.exists(mf):
            with open(mf, encoding="utf-8") as f:
                dados = json.load(f)
            if dados.get("name") != TITULO:
                if not os.path.exists(mf + ".orig"):
                    shutil.copy2(mf, mf + ".orig")
                dados["name"] = TITULO
                dados["short_name"] = "TLM"
                with open(mf, "w", encoding="utf-8") as f:
                    json.dump(dados, f, ensure_ascii=False, indent=2)
                mudou = True
                diz("[OK] manifest.json renomeado.")
            else:
                diz("[=] manifest.json ja estava aplicado.")
    except Exception as e:
        diz("[!] Falha no manifest.json: %s" % e)

    return mudou


if __name__ == "__main__":
    aplicar(verboso=True)
    print("\nPronto. Se o portal estiver aberto, feche a aba e recarregue com Ctrl+Shift+R.")
