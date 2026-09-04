import streamlit as st
import time
import pandas as pd
import os
import html
import hmac
import uuid
import re
import unicodedata
from collections import Counter
from datetime import datetime
import shutil
import streamlit.components.v1 as components

def _garantir_pacote(pkg_import: str, pkg_pip: str | None = None):
    """Instala o pacote silenciosamente se não estiver disponível."""
    import importlib, subprocess, sys
    try:
        importlib.import_module(pkg_import)
    except ImportError:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", pkg_pip or pkg_import, "--quiet"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

@st.cache_resource(show_spinner=False)
def _garantir_dependencias_extras():
    """Garante que os pacotes opcionais estejam instalados — executado uma única
    vez por sessão do servidor (cache_resource). NÃO importa as bibliotecas no
    topo: os imports pesados ficam lazy dentro das funções que os usam, para não
    atrasar o primeiro carregamento da página inicial."""
    _garantir_pacote("openpyxl")

_garantir_dependencias_extras()


@st.cache_resource
def get_server_start_time():
    return time.time()

server_start = get_server_start_time()

@st.cache_data(show_spinner=False)
def carregar_dados_cacheados(caminho, colunas):
    return carregar_dados(caminho, colunas)


@st.cache_data(show_spinner=False)
def _filtrar_periodo_cacheado(df, mes_sel, ano_sel):
    """Filtra um DataFrame pelo mês/ano. Cacheado para evitar re-parsear
    as datas (pd.to_datetime sobre todo o dataset) a cada rerun do dashboard."""
    if df.empty:
        return df
    dc = df.copy()
    dc["__dt"] = pd.to_datetime(dc["Data"], format="%d/%m/%Y %H:%M", errors="coerce")
    return dc[
        (dc["__dt"].dt.month == mes_sel) & (dc["__dt"].dt.year == ano_sel)
    ].drop(columns=["__dt"]).reset_index(drop=True)


@st.cache_data(show_spinner=False)
def _gerar_export_periodo_zip(df_b_all, df_c_all, mes_sel, ano_sel, mes_label):
    """Gera o ZIP/XLSX do período filtrado. Cacheado por mês/ano + dados —
    evita reconstruir o workbook a cada rerun do dashboard (só recalcula
    quando o período ou os dados mudam)."""
    import io as _io, zipfile as _zf

    def _pre_filter(df):
        if df.empty:
            return df
        dc = df.copy()
        dc["__dt"] = pd.to_datetime(dc["Data"], format="%d/%m/%Y %H:%M", errors="coerce")
        return dc[(dc["__dt"].dt.month == mes_sel) & (dc["__dt"].dt.year == ano_sel)].drop(columns=["__dt"])

    _df_exp_ch = _pre_filter(df_b_all)
    _df_exp_cp = _pre_filter(df_c_all)

    def _xlsx_periodo(frames):
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        _hf = PatternFill("solid", fgColor="1E3A5F")
        _hfont = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        _thin = Side(style="thin", color="CCCCCC")
        _brd  = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
        wb2 = Workbook(); wb2.remove(wb2.active)
        for sname, df2 in frames.items():
            ws2 = wb2.create_sheet(title=sname[:31])
            if df2.empty:
                ws2.append(["(sem dados neste período)"]); continue
            ws2.append(list(df2.columns))
            for c in ws2[1]:
                c.font=_hfont; c.fill=_hf
                c.alignment=Alignment(horizontal="center",vertical="center")
                c.border=_brd
            ws2.row_dimensions[1].height = 26
            ef = PatternFill("solid", fgColor="EAF2FB")
            of = PatternFill("solid", fgColor="FFFFFF")
            for ri, row2 in enumerate(df2.itertuples(index=False), start=2):
                ws2.append(list(row2))
                for c in ws2[ri]:
                    c.fill=ef if ri%2==0 else of
                    c.font=Font(name="Arial",size=9)
                    c.border=_brd
            for ci, col2 in enumerate(ws2.columns, start=1):
                w = min(max(len(str(c.value or "")) for c in col2)+4, 52)
                ws2.column_dimensions[get_column_letter(ci)].width = w
            ws2.freeze_panes = "A2"
        b2 = _io.BytesIO(); wb2.save(b2); b2.seek(0)
        return b2.read()

    _buf_exp = _io.BytesIO()
    with _zf.ZipFile(_buf_exp, "w", _zf.ZIP_DEFLATED) as _z:
        if not _df_exp_ch.empty or not _df_exp_cp.empty:
            _frames_periodo = {}
            if not _df_exp_ch.empty:
                _frames_periodo["Chamados"] = _df_exp_ch
            if not _df_exp_cp.empty:
                _frames_periodo["Compras"] = _df_exp_cp
            _z.writestr(
                f"dados_{mes_label.lower()}_{ano_sel}.xlsx",
                _xlsx_periodo(_frames_periodo)
            )
    _buf_exp.seek(0)
    return _buf_exp.read()


_DIR_APP = os.path.dirname(os.path.abspath(__file__))

_RAIZ  = os.path.dirname(_DIR_APP)
_DADOS = os.path.join(_RAIZ, "dados")
os.makedirs(_DADOS, exist_ok=True)
_PORTAL_ICON = os.path.join(_DIR_APP, "tlm_icon.png")
_PORTAL_ICON_FALLBACK = os.path.join(_DIR_APP, "portal_logo.png")


def _icone_pagina():
    """Selo TLM para a aba. NUNCA devolve emoji.

    Antes esta função caía num 🛠️ quando o PNG não era encontrado. O
    problema é que o Chrome guarda favicon em cache de forma agressiva —
    Ctrl+Shift+R não o recarrega — então uma única sessão servida com o
    emoji continuava aparecendo por semanas depois de o arquivo existir.
    Se os dois PNGs sumirem, desenhamos o selo na hora em vez de mostrar
    um ícone genérico."""
    for caminho in (_PORTAL_ICON, _PORTAL_ICON_FALLBACK):
        if os.path.exists(caminho):
            return caminho
    try:
        from PIL import Image, ImageDraw
        img = Image.new("RGBA", (128, 128), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        d.rounded_rectangle([2, 2, 126, 126], radius=24, fill="#FFB700")
        d.text((30, 52), "TLM", fill="#0C0C0C")
        img.save(_PORTAL_ICON)
        return _PORTAL_ICON
    except Exception:
        return None
@st.cache_resource(show_spinner=False)
def _marcar_streamlit_uma_vez():
    try:
        from marca_streamlit import aplicar
        aplicar()
    except Exception:
        pass
    return True

_marcar_streamlit_uma_vez()

st.set_page_config(
    page_title="TLM Service Desk",
    page_icon=_icone_pagina(),
    layout="wide",
)

_primeira_carga = "_app_booted" not in st.session_state
st.session_state["_app_booted"] = True
_fade_css = ("""
[data-testid="stMain"] { animation: antiFoucReveal .32s ease-out both; }
@keyframes antiFoucReveal { from { opacity: 0; } to { opacity: 1; } }
""" if _primeira_carga else "")
_FOUC_BG = "#0d0d0c"
_FOUC_FG = "#f5f3ee"
try:
    if getattr(getattr(st, "context", None), "theme", None) and st.context.theme.type == "light":
        _FOUC_BG, _FOUC_FG = "#f6f4f0", "#14130f"
except Exception:
    pass
st.markdown(f"""
<style id="anti-fouc-critical">
html, body, .stApp,
[data-testid="stAppViewContainer"], [data-testid="stMain"] {{
    background-color: {_FOUC_BG} !important;
    color: {_FOUC_FG} !important;
    /* Pilha de fontes definida JÁ na primeira pintura: enquanto a Geist baixa,
       o texto aparece numa fonte de sistema próxima — elimina o "flash de
       letras sem estilo" (FOUT) e o reflow quando a Geist troca. */
    font-family: 'Archivo', system-ui, -apple-system, 'Segoe UI', Roboto,
                 'Helvetica Neue', Arial, sans-serif !important;
}}
/* Suaviza a substituição da fonte sem piscar */
html {{ -webkit-font-smoothing: antialiased; text-rendering: optimizeLegibility; }}

</style>
""", unsafe_allow_html=True)

@st.cache_data(show_spinner=False)
def _minify_css(css: str) -> str:
    """Minifica o CSS (remove comentários e indentação) para reduzir o payload
    enviado ao navegador a cada rerun. Resultado cacheado — custo pago 1x."""
    import re as _re
    css = _re.sub(r'/\*.*?\*/', '', css, flags=_re.S)
    css = _re.sub(r'(?m)^[ \t]+', '', css)
    css = _re.sub(r'\n{2,}', '\n', css)
    return css



_NEUTRO = {
    "n0": ("#0d0d0c", "#f6f4f0"), "n1": ("#141413", "#ffffff"),
    "n2": ("#1a1a18", "#efece6"), "n3": ("#232320", "#e2ded5"),
    "n4": ("#2f2f2b", "#d3cdc2"), "n5": ("#3d3c37", "#bdb7ac"),
    "n6": ("#55534c", "#a29d94"), "n7": ("#78746c", "#8e8981"),
    "n8": ("#a8a49b", "#6b665d"), "n9": ("#c5c1b8", "#4b473f"),
    "n10":("#e2dfd8", "#2b2924"), "n11":("#f5f3ee", "#14130f"),
}
_NEUTRO_TXT = {
    "n0": ("#3d3c37", "#14130f"), "n1": ("#45443e", "#14130f"),
    "n2": ("#4d4b45", "#14130f"), "n3": ("#6f6b62", "#14130f"),
    "n4": ("#615e56", "#2b2924"), "n5": ("#6f6b62", "#8a857c"),
    "n6": ("#7d786f", "#7d786f"), "n7": ("#948f85", "#6b665d"),
    "n8": ("#b3aea4", "#5c584f"), "n9": ("#cbc7be", "#3d3a34"),
    "n10":("#e2dfd8", "#26241f"), "n11":("#f5f3ee", "#14130f"),
}
_FAM = {
    "acc-d": ("#8a6300", "#7a5600", "#c48d00"),
    "acc":   ("#ffb700", "#8a6000", "#ffb700"),
    "acc-l": ("#ffcc4d", "#a37400", "#ffcc4d"),
    "acc-xl":("#fff0c2", "#6b4a00", "#fff0c2"),
    "run-d": ("#1d4a8a", "#17457f", "#2f6fc0"),
    "run":   ("#6ba7f5", "#1f5fbf", "#3b82f6"),
    "run-l": ("#a8ccfa", "#2a6ac9", "#a8ccfa"),
    "ok-d":  ("#13402a", "#0e5c2f", "#1a7a45"),
    "ok":    ("#4ec27c", "#157a3c", "#1f9950"),
    "ok-l":  ("#9ee0b8", "#12693a", "#9ee0b8"),
    "crit-d":("#5c1b22", "#8f1a28", "#b02030"),
    "crit":  ("#f5707f", "#c02434", "#dc2f42"),
    "crit-l":("#f9aab3", "#a81f2f", "#f9aab3"),
}
def _hx2rgb(h):
    return "%d,%d,%d" % tuple(int(h[i:i+2], 16) for i in (1, 3, 5))

def _tema_atual() -> str:
    """Tema escolhido pelo usuário no menu Configurações > Aparência do
    Streamlit. Usar o mecanismo nativo faz o chrome dos widgets (radio,
    select, tooltip) e o CSS do portal trocarem juntos — um alternador
    próprio só mudaria o CSS e deixaria os widgets no tema errado."""
    try:
        t = getattr(getattr(st, "context", None), "theme", None)
        if t is not None and getattr(t, "type", None) == "light":
            return "claro"
    except Exception:
        pass
    return "escuro"

_SOMBRA = {
    "shc-1": ("rgba(0,0,0,.62)", "rgba(28,25,20,.155)"),
    "shc-2": ("rgba(0,0,0,.42)", "rgba(28,25,20,.105)"),
    "shc-3": ("rgba(0,0,0,.26)", "rgba(28,25,20,.070)"),
    "shc-4": ("rgba(0,0,0,.14)", "rgba(28,25,20,.045)"),
}

_SUP = {
    "sup-a":  ("#1a1a18", "#ffffff"), "sup-b":  ("#141413", "#f8f6f3"),
    "sup2-a": ("#141413", "#ffffff"), "sup2-b": ("#0d0d0c", "#f5f2ee"),
}

_TINTA_ACC = [
    ('ta000', 0.0, 0.0),
    ('ta015', 0.015, 0.065),
    ('ta030', 0.03, 0.102),
    ('ta040', 0.04, 0.123),
    ('ta045', 0.045, 0.133),
    ('ta050', 0.05, 0.143),
    ('ta055', 0.055, 0.152),
    ('ta060', 0.06, 0.161),
    ('ta070', 0.07, 0.178),
    ('ta080', 0.08, 0.194),
    ('ta090', 0.09, 0.209),
    ('ta100', 0.1, 0.224),
    ('ta110', 0.11, 0.238),
    ('ta120', 0.12, 0.252),
    ('ta130', 0.13, 0.265),
    ('ta140', 0.14, 0.279),
    ('ta150', 0.15, 0.291),
    ('ta160', 0.16, 0.304),
    ('ta170', 0.17, 0.316),
    ('ta180', 0.18, 0.328),
    ('ta200', 0.2, 0.351),
    ('ta220', 0.22, 0.374),
    ('ta240', 0.24, 0.395),
    ('ta250', 0.25, 0.406),
    ('ta260', 0.26, 0.417),
    ('ta280', 0.28, 0.437),
    ('ta300', 0.3, 0.457),
    ('ta320', 0.32, 0.477),
    ('ta350', 0.35, 0.505),
    ('ta360', 0.36, 0.515),
    ('ta380', 0.38, 0.533),
    ('ta400', 0.4, 0.551),
    ('ta420', 0.42, 0.569),
    ('ta440', 0.44, 0.586),
    ('ta450', 0.45, 0.595),
    ('ta480', 0.48, 0.621),
    ('ta500', 0.5, 0.637),
    ('ta520', 0.52, 0.654),
    ('ta550', 0.55, 0.678),
    ('ta600', 0.6, 0.717),
    ('ta620', 0.62, 0.733),
    ('ta650', 0.65, 0.756),
    ('ta700', 0.7, 0.793),
    ('ta720', 0.72, 0.808),
    ('ta750', 0.75, 0.829),
    ('ta800', 0.8, 0.865),
    ('ta820', 0.82, 0.879),
    ('ta880', 0.88, 0.92),
    ('ta950', 0.95, 0.967),
]

@st.cache_data(show_spinner=False)
def _tokens_vars(tema: str) -> str:
    """Declarações :root para o tema pedido (sem a tag <style>)."""
    esc = (tema != "claro")
    L = []
    for k, (d, c) in _NEUTRO.items():
        v = d if esc else c
        L.append("--%s:%s;" % (k, v))
        L.append("--rgb-%s:%s;" % (k, _hx2rgb(v)))
    for k, (d, c) in _NEUTRO_TXT.items():
        v = d if esc else c
        L.append("--t-%s:%s;" % (k, v))
        L.append("--rgb-t-%s:%s;" % (k, _hx2rgb(v)))
    for k, (d, ct, cf) in _FAM.items():
        vt = d if esc else ct
        vf = d if esc else cf
        L.append("--c-%s:%s;" % (k, vt))
        L.append("--f-%s:%s;" % (k, vf))
        L.append("--rgb-%s:%s;" % (k, _hx2rgb(vf)))
        L.append("--rgb-c-%s:%s;" % (k, _hx2rgb(vt)))
    L.append("--rgb-lift:%s;" % ("255,255,255" if esc else "20,19,15"))
    for k, (d, c) in _SOMBRA.items():
        L.append("--%s:%s;" % (k, d if esc else c))
    L.append("--campo:%s;"    % ("#1a1a18" if esc else "#ffffff"))
    L.append("--campo-bd:%s;" % ("#2f2f2b" if esc else "#d9d4c9"))
    L.append("--veu:%s;" % ("rgba(0,0,0,.62)" if esc else "rgba(28,25,20,.42)"))
    L.append("--campo-txt:%s;" % ("var(--c-acc-l)" if esc else "var(--t-n11)"))
    for _nt, _ae, _ac in _TINTA_ACC:
        L.append("--%s:rgba(255,183,0,%s);" % (_nt, _ae if esc else _ac))
    for k, (d, c) in _SUP.items():
        v = d if esc else c
        L.append("--%s:%s;" % (k, v))
        L.append("--rgb-%s:%s;" % (k, _hx2rgb(v)))
    L.append("--act:%s;"     % ("var(--c-acc)" if esc else "#0c0c0c"))
    L.append("--act-ink:%s;" % ("#0c0c0c" if esc else "#ffffff"))
    L.append("--act-hover:%s;" % ("#ffc633" if esc else "#2a2a2a"))
    L.append("--brand:var(--c-acc);--brand-ink:#0c0c0c;")
    return ":root{" + "".join(L) + "}"

def _cor(token: str, alpha=None, tema: str = None) -> str:
    """Resolve um token de cor para hex/rgba literal.

    Canvas (Chart.js) e Plotly desenham fora do CSS: recebem a string de cor
    direto, entao `var(--x)` chega como texto invalido e vira preto (ou erro
    de validacao no Plotly). Tudo que sai do CSS para essas bibliotecas passa
    por aqui."""
    esc = (tema or _tema_atual()) != "claro"
    t = token.lstrip("-")
    for pref in ("rgb-t-", "rgb-c-", "rgb-", "t-", "c-", "f-"):
        if t.startswith(pref):
            t = t[len(pref):]
            kind = pref
            break
    else:
        kind = ""
    if t in _SOMBRA:
        return _SOMBRA[t][0 if esc else 1]
    if t.startswith("ta"):
        for _nt, _ae, _ac in _TINTA_ACC:
            if _nt == t:
                return "rgba(255,183,0,%s)" % (_ae if esc else _ac)
    if t in _SUP:
        hx = _SUP[t][0 if esc else 1]
        if alpha is None:
            return hx
        r, g, b = (int(hx[i:i+2], 16) for i in (1, 3, 5))
        return "rgba(%d,%d,%d,%s)" % (r, g, b, alpha)
    if t in _NEUTRO:
        tabela = _NEUTRO_TXT if kind in ("t-", "rgb-t-") else _NEUTRO
        hx = tabela[t][0 if esc else 1]
    elif t in _FAM:
        d, ct, cf = _FAM[t]
        hx = d if esc else (ct if kind in ("c-", "rgb-c-") else cf)
    else:
        hx = "#f5f3ee" if esc else "#14130f"
    if alpha is None:
        return hx
    r, g, b = (int(hx[i:i+2], 16) for i in (1, 3, 5))
    return "rgba(%d,%d,%d,%s)" % (r, g, b, alpha)


def _resolver_vars(txt: str, tema: str = None) -> str:
    """Troca todo var(--x) / rgba(var(--rgb-x),a) por cor literal no texto."""
    txt = re.sub(r"rgba\(\s*var\(--rgb-([a-zA-Z0-9_-]+)\)\s*,\s*([\d.]+)\s*\)",
                 lambda m: _cor(m.group(1), m.group(2), tema), txt)
    return re.sub(r"var\(\s*(--[a-zA-Z0-9_-]+)\s*\)",
                  lambda m: _cor(m.group(1), None, tema), txt)



_SVG_PATHS = {
    "chamado":   '<rect x="2.5" y="4.5" width="15" height="11" rx="2"/><path d="M10 8v4M8 10h4"/>',
    "carrinho":  '<circle cx="8" cy="16.6" r="1.2"/><circle cx="14.6" cy="16.6" r="1.2"/>'
                 '<path d="M2 3.2h2.1l2.2 9.1a1.3 1.3 0 0 0 1.3 1h7.5a1.3 1.3 0 0 0 1.3-1L17.6 6H5"/>',
    "historico": '<path d="M3.2 10a6.8 6.8 0 1 0 2-4.8"/><path d="M3.2 3.2v3.4h3.4"/>'
                 '<path d="M10 6.6V10l2.5 1.6"/>',
    "escudo":    '<path d="M10 2.6 4.2 5v4.4c0 3.5 2.4 6.5 5.8 7.9 3.4-1.4 5.8-4.4 5.8-7.9V5z"/>'
                 '<path d="M8.2 10.1h3.6v3.1H8.2z"/><path d="M8.9 10.1V9a1.1 1.1 0 0 1 2.2 0v1.1"/>',
    "exportar":  '<path d="M11.4 2.6H6A1.5 1.5 0 0 0 4.5 4.1v11.8A1.5 1.5 0 0 0 6 17.4h8a1.5 1.5 0 0 0 1.5-1.5V6.7z"/>'
                 '<path d="M11.4 2.6v4.1h4.1"/><path d="M10 9.4v5M8 12.4l2 2 2-2"/>',
    "monitor":   '<rect x="2.5" y="3.5" width="15" height="10" rx="1.6"/><path d="M7 17h6M10 13.5V17"/>',
    "mouse":     '<rect x="6" y="2.6" width="8" height="14.8" rx="4"/><path d="M10 6v3"/>',
    "rede":      '<path d="M2.6 7.4a11.6 11.6 0 0 1 14.8 0"/><path d="M5.4 10.5a7.7 7.7 0 0 1 9.2 0"/>'
                 '<path d="M8.1 13.5a3.8 3.8 0 0 1 3.8 0"/><circle cx="10" cy="16.2" r=".9"/>',
    "email":     '<rect x="2.5" y="4.6" width="15" height="10.8" rx="1.6"/><path d="m2.9 5.7 7.1 5 7.1-5"/>',
    "impressora":'<path d="M6 7.4V3h8v4.4"/><rect x="2.5" y="7.4" width="15" height="6" rx="1.6"/>'
                 '<rect x="6" y="12" width="8" height="5.4" rx="1"/>',
    "sistema":   '<rect x="2.5" y="3.5" width="15" height="13" rx="2"/><path d="M2.5 7.4h15"/>'
                 '<circle cx="5.2" cy="5.5" r=".65"/><circle cx="7.3" cy="5.5" r=".65"/>',
    "telefone":  '<rect x="6" y="2.6" width="8" height="14.8" rx="2"/><path d="M9 15h2"/>',
    "chave":     '<circle cx="7" cy="8" r="3.4"/><path d="m9.4 10.4 6.1 6.1M13.4 14.4l1.5-1.5M15 16l1.4-1.4"/>',
    "duvida":    '<circle cx="10" cy="10" r="7.4"/>'
                 '<path d="M8.2 8a1.9 1.9 0 1 1 2.6 1.8c-.5.2-.8.7-.8 1.2v.4"/><circle cx="10" cy="14" r=".8"/>',
    "ok":        '<circle cx="10" cy="10" r="7.4"/><path d="m6.6 10.2 2.3 2.3 4.5-4.6"/>',
    "info":      '<circle cx="10" cy="10" r="7.4"/><path d="M10 9.2v5"/><circle cx="10" cy="6.4" r=".9"/>',
    "clipe":     '<path d="M14.4 9.1 9.6 13.9a3 3 0 0 1-4.3-4.3l5.5-5.5a2.1 2.1 0 0 1 2.9 2.9l-5.5 5.5'
                 'a1 1 0 0 1-1.5-1.5l5-5"/>',
    "raio":      '<path d="M11.2 2.5 4.8 11h4.2l-.8 6.5L15.2 9H11z"/>',
    "alerta":    '<path d="M10 3.2 2.6 16.2h14.8z"/><path d="M10 8v3.4"/><circle cx="10" cy="13.9" r=".85"/>',
    "engrenagem":'<circle cx="10" cy="10" r="2.6"/><path d="M15.6 12.2a1.2 1.2 0 0 0 .24 1.32l.05.04a1.5 1.5 0 1 1-2.12 2.12l-.04-.05'
                 'a1.2 1.2 0 0 0-1.32-.24 1.2 1.2 0 0 0-.73 1.1v.12a1.5 1.5 0 1 1-3 0v-.06a1.2 1.2 0 0 0-.79-1.1 1.2 1.2 0 0 0-1.32.24'
                 'l-.04.05a1.5 1.5 0 1 1-2.12-2.12l.05-.04a1.2 1.2 0 0 0 .24-1.32 1.2 1.2 0 0 0-1.1-.73H3.5a1.5 1.5 0 1 1 0-3h.06'
                 'a1.2 1.2 0 0 0 1.1-.79 1.2 1.2 0 0 0-.24-1.32l-.05-.04A1.5 1.5 0 1 1 6.49 4.2l.04.05a1.2 1.2 0 0 0 1.32.24h.06'
                 'a1.2 1.2 0 0 0 .73-1.1V3.3a1.5 1.5 0 1 1 3 0v.06a1.2 1.2 0 0 0 .73 1.1 1.2 1.2 0 0 0 1.32-.24l.04-.05a1.5 1.5 0 1 1 2.12 2.12'
                 'l-.05.04a1.2 1.2 0 0 0-.24 1.32v.06a1.2 1.2 0 0 0 1.1.73h.12a1.5 1.5 0 1 1 0 3h-.06a1.2 1.2 0 0 0-1.1.73z"/>',
    "ampulheta": '<path d="M6 2.6h8M6 17.4h8"/><path d="M6.6 2.6v3.1c0 1.3 3.4 3 3.4 4.3s-3.4 3-3.4 4.3v3.1"/>'
                 '<path d="M13.4 2.6v3.1c0 1.3-3.4 3-3.4 4.3s3.4 3 3.4 4.3v3.1"/>',
    "relogio":   '<circle cx="10" cy="10" r="7.4"/><path d="M10 5.6V10l3 1.9"/>',
    "caixa":     '<path d="M17 6.6 10 3 3 6.6v6.8L10 17l7-3.6z"/><path d="M3 6.6 10 10.3l7-3.7M10 10.3V17"/>',
    "grafico":   '<path d="M2.8 16.6h14.4"/><path d="M5.6 14V8.4M10 14V3.6M14.4 14v-4.2"/>',
    "calendario":'<rect x="2.8" y="4.4" width="14.4" height="12.2" rx="1.8"/><path d="M2.8 8h14.4M6.6 2.6v3.4M13.4 2.6v3.4"/>',
    "predio":    '<path d="M3.2 17V4.4A1.4 1.4 0 0 1 4.6 3h6.2a1.4 1.4 0 0 1 1.4 1.4V17"/>'
                 '<path d="M12.2 8.4h3.2A1.4 1.4 0 0 1 16.8 9.8V17"/><path d="M2.4 17h15.2"/>'
                 '<path d="M6 6.2h1.4M6 9.4h1.4M6 12.6h1.4M9.4 6.2h1.4M9.4 9.4h1.4M9.4 12.6h1.4"/>',
    "pessoa":    '<circle cx="10" cy="6.4" r="3"/><path d="M3.8 17c0-2.9 2.8-4.6 6.2-4.6s6.2 1.7 6.2 4.6"/>',
    "alvo":      '<circle cx="10" cy="10" r="7.2"/><circle cx="10" cy="10" r="3.9"/><circle cx="10" cy="10" r=".9"/>',
    "prancheta": '<rect x="4.2" y="3.6" width="11.6" height="13.2" rx="1.6"/>'
                 '<path d="M7.6 3.6V2.8a1 1 0 0 1 1-1h2.8a1 1 0 0 1 1 1v.8z"/><path d="M7.4 9h5.2M7.4 12.2h3.4"/>',
    "trofeu":    '<path d="M6.4 3h7.2v4.2a3.6 3.6 0 0 1-7.2 0z"/><path d="M6.4 4.2H4.2v1a2.6 2.6 0 0 0 2.2 2.5"/>'
                 '<path d="M13.6 4.2h2.2v1a2.6 2.6 0 0 1-2.2 2.5"/><path d="M10 10.8v3.4M7 17h6l-.6-2.8H7.6z"/>',
    "filtro":    '<path d="M2.8 3.8h14.4l-5.6 6.6V16l-3.2 1.4v-7z"/>',
    "busca":     '<circle cx="8.8" cy="8.8" r="5.4"/><path d="m12.8 12.8 4 4"/>',
    "cadeado":   '<rect x="4.4" y="8.6" width="11.2" height="8.4" rx="1.8"/>'
                 '<path d="M6.8 8.6V6.4a3.2 3.2 0 0 1 6.4 0v2.2"/>',
    "sync":      '<path d="M16.6 8.6A6.8 6.8 0 0 0 4.5 5.6"/><path d="M3.4 11.4a6.8 6.8 0 0 0 12.1 3"/>'
                 '<path d="M3.4 3.4v3.4h3.4M16.6 16.6v-3.4h-3.4"/>',
    "baixar":    '<path d="M3.4 12.8V15a1.8 1.8 0 0 0 1.8 1.8h9.6A1.8 1.8 0 0 0 16.6 15v-2.2"/>'
                 '<path d="M10 3.2v8.6M6.6 8.6 10 12l3.4-3.4"/>',
    "planilha":  '<rect x="3" y="3.4" width="14" height="13.2" rx="1.6"/><path d="M3 8h14M3 12h14M8 3.4v13.2M13 3.4v13.2"/>',
    "documento": '<path d="M11.6 2.8H6.2A1.6 1.6 0 0 0 4.6 4.4v11.2a1.6 1.6 0 0 0 1.6 1.6h7.6a1.6 1.6 0 0 0 1.6-1.6V6.4z"/>'
                 '<path d="M11.6 2.8v3.6h3.8"/><path d="M7.6 10.4h4.8M7.6 13.2h3.2"/>',
    "pasta":     '<path d="M2.8 5.6A1.6 1.6 0 0 1 4.4 4h3.1l1.6 2h6.5a1.6 1.6 0 0 1 1.6 1.6v6.8a1.6 1.6 0 0 1-1.6 1.6H4.4a1.6 1.6 0 0 1-1.6-1.6z"/>',
    "lapis":     '<path d="M13.4 3.4a1.9 1.9 0 0 1 2.7 2.7L7 15.2l-3.6 1 1-3.6z"/><path d="m12.2 4.6 3.2 3.2"/>',
    "salvar":    '<path d="M3.4 5.2A1.8 1.8 0 0 1 5.2 3.4h8.2l3.2 3.2v8.2a1.8 1.8 0 0 1-1.8 1.8H5.2a1.8 1.8 0 0 1-1.8-1.8z"/>'
                 '<path d="M6.4 3.4v4.4h6V3.4M6.4 16.6v-4.4h7.2v4.4"/>',
    "ferramenta":'<path d="M13.6 2.8a4.2 4.2 0 0 0-4 5.5l-6 6a1.7 1.7 0 0 0 2.4 2.4l6-6a4.2 4.2 0 0 0 5-5.6l-2.6 2.6-2.4-.6-.6-2.4z"/>',
}


def _ico(nome, tam=16, cor="currentColor", peso=1.6, extra=""):
    """SVG inline pronto para HTML."""
    d = _SVG_PATHS.get(nome, _SVG_PATHS["duvida"])
    return (
        '<svg viewBox="0 0 20 20" width="{t}" height="{t}" fill="none" '
        'stroke="{c}" stroke-width="{w}" stroke-linecap="round" '
        'stroke-linejoin="round" style="flex:none;vertical-align:-.18em;{e}" '
        'aria-hidden="true">{d}</svg>'
    ).format(t=tam, c=cor, w=peso, d=d, e=extra)


def _ico_uri(nome, peso=1.6):
    """Mesmo ícone como data-URI, para usar em mask-image no CSS.

    Em mask-image a cor vem do background-color do elemento, então o ícone
    acompanha o tema e o estado (ativo/hover) sem precisar de duplicata."""
    from urllib.parse import quote
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20" fill="none" '
        'stroke="#000" stroke-width="%s" stroke-linecap="round" stroke-linejoin="round">%s</svg>'
        % (peso, _SVG_PATHS.get(nome, _SVG_PATHS["duvida"]))
    )
    return "url(\"data:image/svg+xml,%s\")" % quote(svg, safe="")


_ICONES_NAV = ["chamado", "carrinho", "historico", "escudo", "exportar"]

_ICONE_CATEGORIA = {
    "Hardware": "monitor", "Periféricos": "mouse", "Internet/Rede": "rede",
    "E-mail/Login": "email", "Impressora": "impressora", "Sistema/SW": "sistema",
    "Telefonia": "telefone", "Acesso/Permissão": "chave", "Outro": "duvida",
    "Infra/Rede": "rede", "Licença/SW": "sistema", "Suprimentos": "impressora",
    "Mobilidade": "telefone", "Energia": "raio",
}


def _css_icones_nav():
    """Regras que põem o ícone dentro de cada item do menu lateral."""
    regras = ["""
[data-testid="stSidebar"] div[role="radiogroup"] > label [data-testid="stMarkdownContainer"]{
    display:flex !important; align-items:center !important; gap:10px !important;
}
[data-testid="stSidebar"] div[role="radiogroup"] > label [data-testid="stMarkdownContainer"]::before{
    content:""; width:17px; height:17px; flex:none;
    background-color: var(--t-n8);
    -webkit-mask: var(--ico-nav) center/contain no-repeat;
            mask: var(--ico-nav) center/contain no-repeat;
    transition: background-color .2s ease;
}
[data-testid="stSidebar"] div[role="radiogroup"] > label:hover [data-testid="stMarkdownContainer"]::before,
[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) [data-testid="stMarkdownContainer"]::before{
    background-color: var(--c-acc);
}
/* O ponto decorativo à direita sai: com ícone à esquerda e barra de ativo,
   ele só disputava espaço com o rótulo, que quebrava em duas linhas. */
[data-testid="stSidebar"] div[role="radiogroup"] > label::after{ display:none !important; }
[data-testid="stSidebar"] div[role="radiogroup"] > label p{
    font-size:.8rem !important; letter-spacing:.005em !important;
}
"""]
    for i, nome in enumerate(_ICONES_NAV, start=1):
        regras.append(
            '[data-testid="stSidebar"] div[role="radiogroup"] > label:nth-of-type(%d)'
            '{--ico-nav:%s;}' % (i, _ico_uri(nome))
        )
    return "\n".join(regras)


def _tokens_css(tema: str) -> str:
    return ("<style id=\"tlm-tokens\">" + _tokens_vars(tema)
            + _css_icones_nav() + "</style>")

@st.cache_data(show_spinner=False)
def _get_global_css(tema: str = "escuro") -> str:
    """CSS global cacheado — montado uma vez por tema."""
    return _tokens_css(tema) + """
<style>
/* Fontes carregadas via <link> (preconnect) fora deste bloco — um @import aqui
   bloquearia a aplicação de TODO o CSS até a fonte baixar, causando FOUC
   (texto sem estilo antes do design carregar). */

/* ── CORREÇÕES ANTI-MODO CLARO (FORÇAR DARK THEME) ── */
[data-testid="stAppViewContainer"] {
    background: linear-gradient(180deg, var(--n0) 0%, var(--n0) 48%, var(--n0) 100%) !important;
    color: var(--t-n11) !important; /* var(--txt-1) */
}
[data-testid="stHeader"] {
    background: transparent !important;
    
    
}
[data-testid="stSidebar"] {
    background-color: var(--n1) !important; /* var(--bg-1) */
}
/* Força títulos, textos e labels nativos a não ficarem pretos */
.st-emotion-cache-16txtl3, .st-emotion-cache-10trblm, h1, h2, h3, h4, h5, h6, p, label, span {
    color: var(--t-n11) !important;
}
/* Evita que inputs nativos fiquem com texto escuro no modo claro */
input, textarea, select {
    color: var(--t-n11) !important;
    -webkit-text-fill-color: var(--t-n11) !important;
}

/* ── TOKENS ── */
:root {
    --bg-0:       var(--n0);
    --bg-1:       var(--n1);
    --bg-2:       var(--n1);
    --bg-card:    var(--n2);
    --panel:      rgba(var(--rgb-n2),.92);
    --panel-2:    rgba(var(--rgb-n1),.85);
    --bg-hover:   var(--n3);
    --border:     var(--ta080);
    --border-soft:rgba(var(--rgb-n9),.10);
    --border-hi:  var(--ta550);
    --accent:     var(--c-acc);
    --accent-dim: var(--ta140);
    --accent-soft:var(--ta070);
    --violet:     var(--c-run-l);
    --violet-soft:rgba(var(--rgb-run-l),.08);
    --green:      var(--c-ok);
    --amber:      var(--c-run);
    --red:        var(--c-crit-l);
    --txt-1:      var(--n11);
    --txt-2:      var(--n8);
    --txt-3:      var(--n5);
    --font:       'Archivo', sans-serif;
    --mono:       'IBM Plex Mono', monospace;
    --radius:     14px;
    /* Typography scale */
    --text-xs:    0.70rem;
    --text-sm:    0.78rem;
    --text-md:    0.86rem;
    /* Shadows */
    --shadow:     0 20px 50px var(--shc-1), 0 0 0 1px var(--ta040);
    --shadow-soft:0 1px 2px var(--shc-4), 0 10px 26px var(--shc-2);
    --glow-accent: 0 0 0 rgba(0,0,0,0);
    --glow-violet: 0 0 0 rgba(0,0,0,0);
    --glow-green: 0 0 0 rgba(0,0,0,0);
    --glow-red: 0 0 0 rgba(0,0,0,0);
}

/* ── RESET / BASE ── */
.stApp {
    background:
        linear-gradient(var(--ta015) 1px, transparent 1px),
        linear-gradient(90deg, var(--ta015) 1px, transparent 1px);
    background-size: 44px 44px;
    font-family: var(--font);
}
.stApp::before {
    content: "";
    position: fixed;
    inset: 0;
    pointer-events: none;
    background:
        linear-gradient(180deg, rgba(var(--rgb-lift),.03), transparent 18%),
        linear-gradient(90deg, transparent, var(--ta080), transparent);
    opacity: .34;
    mask-image: linear-gradient(180deg, rgba(0,0,0,.9), transparent 82%);
}
.main .block-container {
    padding: 2.6rem 2.6rem 3.8rem;
    max-width: 1240px;
}
h1,h2,h3,h4,h5 { font-family: var(--font) !important; color: var(--txt-1) !important; }
p, li { color: var(--txt-2); }
hr {
    border: 0 !important;
    height: 1px !important;
    background: linear-gradient(90deg, transparent, var(--ta450), transparent) !important;
    box-shadow: none;
    margin: 22px 0 !important;
}
/* ── SCROLLBARS GLOBAIS — azul ciano combinando com a sidebar ── */
/* Firefox */
*, html, body {
    scrollbar-width: thin;
    scrollbar-color: var(--ta550) transparent;
}

/* Webkit (Chrome, Edge, Safari) — todos os elementos */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: rgba(var(--rgb-n1),.40); border-radius: 99px; }
::-webkit-scrollbar-thumb {
    background: linear-gradient(180deg,
        var(--f-acc) 0%,
        var(--ta650) 60%,
        var(--ta350) 100%);
    border-radius: 99px;
    border: none;
    box-shadow: none;
    transition: background .2s, box-shadow .2s;
}
::-webkit-scrollbar-thumb:hover {
    background: linear-gradient(180deg,
        var(--f-acc-l) 0%,
        var(--f-acc) 50%,
        var(--ta800) 100%);
    box-shadow: none;
}
::-webkit-scrollbar-corner { background: transparent; }

/* Sidebar scrollbar — mesma cor, ligeiramente mais fina */
[data-testid="stSidebar"] ::-webkit-scrollbar { width: 3px; }
[data-testid="stSidebar"] ::-webkit-scrollbar-track { background: transparent; }
[data-testid="stSidebar"] ::-webkit-scrollbar-thumb {
    background: linear-gradient(180deg,
        var(--f-acc) 0%,
        var(--ta600) 100%);
    border-radius: 99px;
    border: none;
    box-shadow: none;
}
[data-testid="stSidebar"] ::-webkit-scrollbar-thumb:hover {
    background: var(--f-acc);
    box-shadow: none;
}
[data-testid="stSidebar"] > div {
    scrollbar-width: thin;
    scrollbar-color: var(--ta550) transparent;
}

/* Dataframes, tabelas e containers internos do Streamlit */
[data-testid="stDataFrame"] *,
[data-testid="stDataFrame"] iframe,
.stDataFrame *,
.dvn-scroller,
.glideDataEditor,
[class*="scrollable"],
[class*="overflow"],
.stTabContent,
.element-container * {
    scrollbar-width: thin !important;
    scrollbar-color: var(--ta550) transparent !important;
}
[data-testid="stDataFrame"] ::-webkit-scrollbar,
.dvn-scroller::-webkit-scrollbar,
.glideDataEditor ::-webkit-scrollbar { width: 5px !important; height: 5px !important; }
[data-testid="stDataFrame"] ::-webkit-scrollbar-track,
.dvn-scroller::-webkit-scrollbar-track { background: rgba(var(--rgb-n1),.40) !important; border-radius: 99px !important; }
[data-testid="stDataFrame"] ::-webkit-scrollbar-thumb,
.dvn-scroller::-webkit-scrollbar-thumb {
    background: linear-gradient(180deg, var(--f-acc) 0%, var(--ta550) 100%) !important;
    border-radius: 99px !important;
    box-shadow: none;
}
[data-testid="stDataFrame"] ::-webkit-scrollbar-thumb:hover,
.dvn-scroller::-webkit-scrollbar-thumb:hover {
    background: var(--f-acc-l) !important;
    box-shadow: none;
}

/* Injeta o estilo dentro dos iframes de dataframe via CSS variable trick */
iframe { color-scheme: dark; }
footer, .stDeployButton, [data-testid="stAppDeployButton"], [data-testid="stDecoration"] { display: none !important; visibility: hidden !important; }
#MainMenu { display: inline-flex !important; visibility: visible !important; opacity: .55; }
#MainMenu:hover { opacity: 1; }

/* ── SIDEBAR / NAV SHELL ── */
[data-testid="stSidebar"] {
    position: relative !important;
    height: 100vh !important;
    top: 0 !important;
    background: linear-gradient(180deg, var(--sup2-a) 0%, var(--sup2-b) 54%, var(--n0) 100%) !important;
    border-right: 1px solid var(--ta130) !important;
    box-shadow: inset -1px 0 0 rgba(var(--rgb-lift),.03), 18px 0 60px var(--shc-3) !important;
    transition: transform .22s ease, width .22s ease, box-shadow .22s ease !important;
}
[data-testid="stSidebar"]::before {
    content: "";
    position: absolute;
    inset: 0;
    pointer-events: none;
    background-image:
        linear-gradient(var(--ta045) 1px, transparent 1px),
        linear-gradient(90deg, var(--ta045) 1px, transparent 1px);
    background-size: 28px 28px;
    mask-image: linear-gradient(180deg, rgba(0,0,0,.85), rgba(0,0,0,.08));
}
[data-testid="stSidebar"] > div {
    padding: 0 16px 28px !important;
    height: 100% !important;
    overflow-y: auto !important;
    overflow-x: hidden !important;
}

/* Sidebar collapse control */
[data-testid="collapsedControl"] {
    transition: transform .18s ease, background .18s ease, border-color .18s ease, box-shadow .18s ease !important;
    border-radius: 12px !important;
    border: 1px solid var(--ta220) !important;
    background: rgba(var(--rgb-n1),.72) !important;
    box-shadow: 0 10px 28px var(--shc-2), 0 0 0 1px var(--ta060) !important;
    
}
[data-testid="collapsedControl"]:hover {
    transform: translateY(-1px) scale(1.02) !important;
    border-color: var(--ta450) !important;
    box-shadow: 0 16px 40px var(--shc-2);
}

/* ── MAIN NAVIGATION AS PRODUCT-GRADE ACTION ROWS (ATUALIZADO PARA MICROINTERAÇÕES) ── */
[data-testid="stSidebar"] .stRadio {
    margin-top: 14px !important;
}
[data-testid="stSidebar"] div[role="radiogroup"] {
    gap: 10px !important;
    display: flex !important;
    flex-direction: column !important;
}

/* Seletor universal mapeado direto no DOM nativo do Streamlit */
[data-testid="stSidebar"] div[role="radiogroup"] > label {
    position: relative !important;
    display: flex !important;
    align-items: center !important;
    width: 100% !important;
    min-height: 50px !important;
    margin: 0 !important;
    padding: 4px 14px !important;
    border-radius: 10px !important;
    background: linear-gradient(135deg, rgba(var(--rgb-sup-a),.72), rgba(var(--rgb-sup-b),.52)) !important;
    border: 1px solid rgba(var(--rgb-n9),.10) !important;
    box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.035) !important;
    cursor: pointer !important;
    /* Transições de alta performance */
    transition: transform 0.25s cubic-bezier(0.4, 0, 0.2, 1), 
                border-color 0.25s cubic-bezier(0.4, 0, 0.2, 1), 
                background 0.25s cubic-bezier(0.4, 0, 0.2, 1), 
                box-shadow 0.25s cubic-bezier(0.4, 0, 0.2, 1) !important;
}

/* Oculta a bolinha padrão do Streamlit. O 1º filho do <label> é um <span>
   com o input escondido, então mira-se o irmão anterior ao rótulo. */
/* O `:not([data-testid])` é essencial: no Streamlit 1.56 o contêiner do texto
   É o primeiro filho, e sem essa exclusão a regra escondia o próprio rótulo
   (item de menu ficava vazio). Com ela, funciona no 1.56 e no 1.62. */
[data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child:not([data-testid]),
[data-testid="stSidebar"] div[role="radiogroup"] label
    div:has(> [data-testid="stMarkdownContainer"]) > div:first-child:not([data-testid]) {
    display: none !important;
}
/* Cinto de segurança: o rótulo e seu ícone nunca somem. */
[data-testid="stSidebar"] div[role="radiogroup"] > label [data-testid="stMarkdownContainer"]{
    display: flex !important;
}

[data-testid="stSidebar"] div[role="radiogroup"] > label [data-testid="stMarkdownContainer"] {
    width: 100% !important;
    padding: 0 !important;
}

[data-testid="stSidebar"] div[role="radiogroup"] > label p {
    margin: 0 !important;
    color: var(--t-n11) !important;
    font-size: 0.86rem !important;
    font-weight: 650 !important;
    line-height: 1.2 !important;
    letter-spacing: 0.01em !important;
    text-transform: uppercase !important;
    transition: color 0.2s ease, text-shadow 0.2s ease !important;
}

/* ── MICROINTERAÇÃO: HOVER ── */
[data-testid="stSidebar"] div[role="radiogroup"] > label:hover {
    transform: translateX(6px) !important; /* Deslocamento físico lateral fluido */
    border-color: var(--ta380) !important;
    background: linear-gradient(135deg, var(--ta140), rgba(var(--rgb-n2),.74)) !important;
    box-shadow: 0 12px 28px var(--shc-2), inset 0 1px 0 rgba(var(--rgb-lift),.05) !important;
}

[data-testid="stSidebar"] div[role="radiogroup"] > label:hover p {
    color: var(--t-n11) !important;
    text-shadow: none;
}

/* ── MICROINTERAÇÃO: CLIQUE ATIVO (SENSACÃO TÁTIL) ── */
[data-testid="stSidebar"] div[role="radiogroup"] > label:active {
    transform: translateX(4px) scale(0.985) !important;
    filter: brightness(0.9);
}

/* ── ESTADO SELECIONADO / ATIVO ── */
[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) {
    border-color: var(--ta600) !important;
    background: 
        linear-gradient(135deg, var(--ta220), rgba(var(--rgb-run-l),.12)), 
        linear-gradient(135deg, rgba(var(--rgb-sup-a),.95), rgba(var(--rgb-sup-b),.80)) !important;
    box-shadow: 0 0 0 1px var(--ta160), 0 16px 38px var(--ta140) !important;
}

[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) p {
    color: var(--t-n11) !important;
    font-weight: 800 !important;
    text-shadow: none;
}

/* Indicador Vertical Esquerdo (Neon) */
[data-testid="stSidebar"] div[role="radiogroup"] > label::before {
    content: "" !important;
    position: absolute !important;
    left: 0 !important;
    top: 25% !important;
    bottom: 25% !important;
    width: 3px !important;
    border-radius: 0 4px 4px 0 !important;
    background: transparent !important;
    transform: scaleY(0) !important;
    transition: transform 0.25s cubic-bezier(0.4, 0, 0.2, 1), background 0.2s ease !important;
}

[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked)::before {
    background: var(--f-acc) !important;
    box-shadow: none;
    transform: scaleY(1) !important;
}

/* Dot Indicator Direito */
[data-testid="stSidebar"] div[role="radiogroup"] > label::after {
    content: "" !important;
    position: absolute !important;
    right: 16px !important;
    top: 50% !important;
    width: 6px !important;
    height: 6px !important;
    border-radius: 999px !important;
    background: rgba(var(--rgb-n8),.35) !important;
    transform: translateY(-50%) !important;
    transition: background 0.2s ease, box-shadow 0.2s ease, transform 0.2s ease !important;
}

[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked)::after {
    background: var(--f-acc) !important;
    box-shadow: none;
    transform: translateY(-50%) scale(1.2) !important;
}

/* ── CONTAINERS / CARDS ──────────────────────────────────────────────
   O Streamlit 1.56 nao emite mais `stVerticalBlockBorderWrapper`: um
   `st.container(border=True)` vira um `stVerticalBlock` comum dentro de um
   `stLayoutWrapper`, sem nenhum atributo que o distinga de um container sem
   borda. Por isso cada cartao recebe `key="cartao_*"` no Python, e o gancho
   aqui e a classe `st-key-cartao*` que o proprio Streamlit gera a partir da
   chave — estavel entre versoes, ao contrario do hash do emotion.
   O seletor antigo fica junto para versoes mais velhas. */
[data-testid="stVerticalBlockBorderWrapper"],
[class*="st-key-cartao"],
[data-testid="stForm"],
[data-testid="stExpander"] {
    position: relative !important;
    background: var(--n1) !important;
    border: 1px solid var(--n3) !important;
    border-radius: 10px !important;
    /* Elevacao em dois degraus: um contorno curto que assenta a caixa na
       pagina e uma sombra larga que a levanta. Os dois tokens mudam com o
       tema — no claro o alfa cai para menos de um terco. */
    box-shadow: 0 1px 2px var(--shc-4), 0 8px 22px var(--shc-3) !important;
    transition: border-color .18s ease, box-shadow .18s ease !important;
}
[data-testid="stVerticalBlockBorderWrapper"]:hover,
[class*="st-key-cartao"]:hover {
    border-color: var(--ta420) !important;
    box-shadow: 0 1px 2px var(--shc-4), 0 12px 30px var(--shc-2) !important;
}

/* ════════════════════════════════════════════════════════
   FORM INPUTS — PREMIUM, SELETORES data-baseweb DIRETOS
   ════════════════════════════════════════════════════════ */

/* Labels */
.stTextInput label, .stTextArea label, .stSelectbox label,
.stNumberInput label, .stFileUploader label,
.stMultiSelect label, .stSlider label {
    color: var(--c-acc-xl) !important;
    font-size: .82rem !important;
    font-family: var(--font) !important;
    font-weight: 700 !important;
    text-transform: none !important;
    letter-spacing: .01em !important;
    margin-bottom: 5px !important;
}

/* ── base-input: container real do st.text_input ── */
div[data-baseweb="base-input"] {
    background: linear-gradient(160deg, var(--n1) 0%, var(--n1) 100%) !important;
    border: 1px solid var(--ta220) !important;
    border-radius: 10px !important;
    box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.04), inset 0 -1px 0 var(--shc-2), 0 4px 18px var(--shc-2) !important;
    transition: border-color .22s ease, box-shadow .22s ease, transform .18s ease !important;
}
div[data-baseweb="base-input"]:hover {
    border-color: var(--ta400) !important;
    transform: translateY(-1px) !important;
}
div[data-baseweb="base-input"]:focus-within {
    border-color: var(--ta720) !important;
    box-shadow: 0 0 0 3px var(--ta130), inset 0 1px 0 var(--ta100) !important;
    transform: translateY(-2px) !important;
}
div[data-baseweb="base-input"] input {
    background: transparent !important;
    color: var(--campo-txt) !important;
    font-family: var(--font) !important;
    font-size: .90rem !important;
    font-weight: 500 !important;
    caret-color: var(--c-acc) !important;
    -webkit-text-fill-color: var(--c-acc-l) !important;
}
div[data-baseweb="base-input"] input::placeholder {
    color: rgba(var(--rgb-t-n9),.68) !important;
    -webkit-text-fill-color: rgba(var(--rgb-t-n9),.68) !important;
    font-style: normal !important;
    font-size: .85rem !important;
}

/* ── textarea ── */
div[data-baseweb="textarea"] {
    background: linear-gradient(160deg, var(--n1) 0%, var(--n1) 100%) !important;
    border: 1px solid var(--ta220) !important;
    border-radius: 10px !important;
    box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.04), 0 4px 18px var(--shc-2) !important;
    transition: border-color .22s ease, box-shadow .22s ease, transform .18s ease !important;
}
div[data-baseweb="textarea"]:hover {
    border-color: var(--ta400) !important;
    transform: translateY(-1px) !important;
}
div[data-baseweb="textarea"]:focus-within {
    border-color: var(--ta720) !important;
    box-shadow: 0 0 0 3px var(--ta130), inset 0 1px 0 var(--ta100) !important;
    transform: translateY(-2px) !important;
}
div[data-baseweb="textarea"] textarea {
    background: transparent !important;
    color: var(--campo-txt) !important;
    -webkit-text-fill-color: var(--campo-txt) !important;
    font-family: var(--font) !important;
    font-size: .88rem !important;
    caret-color: var(--c-acc) !important;
}

/* ── select: container visual ── */
div[data-baseweb="select"] > div:first-child {
    background: linear-gradient(160deg, var(--n1) 0%, var(--n1) 100%) !important;
    border: 1px solid var(--ta220) !important;
    border-radius: 10px !important;
    box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.04), 0 4px 18px var(--shc-2) !important;
    transition: border-color .22s ease, box-shadow .22s ease, transform .18s ease !important;
}
div[data-baseweb="select"] > div:first-child:hover {
    border-color: var(--ta400) !important;
    transform: translateY(-1px) !important;
}
div[data-baseweb="select"]:focus-within > div:first-child {
    border-color: var(--ta720) !important;
    box-shadow: 0 0 0 3px var(--ta130), inset 0 1px 0 var(--ta100) !important;
    transform: translateY(-2px) !important;
}
div[data-baseweb="select"] span,
div[data-baseweb="select"] div[class*="single"] {
    color: var(--campo-txt) !important;
    -webkit-text-fill-color: var(--campo-txt) !important;
    font-family: var(--font) !important;
    font-size: .88rem !important;
    font-weight: 500 !important;
}
div[data-baseweb="select"] svg {
    fill: var(--ta520) !important;
    transition: transform .22s ease, fill .2s ease !important;
}
div[data-baseweb="select"]:focus-within svg {
    fill: var(--f-acc) !important;
    transform: rotate(180deg) !important;
}

/* ── Dropdown menu ── */
ul[data-baseweb="menu"] {
    background: linear-gradient(160deg, var(--n1), var(--n1)) !important;
    border: 1px solid var(--ta260) !important;
    border-radius: 12px !important;
    box-shadow: 0 0 0 1px var(--ta060), 0 24px 60px var(--shc-1);
    padding: 5px !important;
    
    overflow: hidden !important;
}
li[role="option"] {
    border-radius: 8px !important;
    color: var(--t-n9) !important;
    font-family: var(--font) !important;
    font-size: .87rem !important;
    font-weight: 500 !important;
    transition: background .14s ease, color .14s ease, padding-left .14s ease !important;
    padding: 9px 14px !important;
}
li[role="option"]:hover {
    background: linear-gradient(90deg, var(--ta160), var(--ta060)) !important;
    color: var(--t-n11) !important;
    padding-left: 18px !important;
}
li[role="option"][aria-selected="true"] {
    background: linear-gradient(90deg, var(--ta180), var(--ta060)) !important;
    color: var(--c-acc-l) !important;
    font-weight: 700 !important;
    border-left: 2px solid var(--ta750) !important;
}

/* ── FILTER PANEL v2 — TOGGLE CHIP DESIGN ── */
.fpanel {
    position: relative;
    background: linear-gradient(145deg, rgba(var(--rgb-sup2-a),.98), rgba(var(--rgb-sup2-b),.96));
    border: 1px solid var(--ta140);
    border-radius: 14px;
    overflow: hidden;
    box-shadow: 0 0 0 1px var(--ta040), 0 28px 64px var(--shc-1);
    margin-bottom: 0;
}
.fpanel::before {
    content: "";
    position: absolute;
    top: 0; left: 0; right: 0; height: 1px;
    background: linear-gradient(90deg, transparent 3%, var(--ta550) 38%, rgba(var(--rgb-run-l),.38) 62%, transparent 97%);
}
.fpanel-violet { border-color: rgba(var(--rgb-run-l),.16); }
.fpanel-violet::before {
    background: linear-gradient(90deg, transparent 3%, rgba(var(--rgb-run-l),.55) 38%, var(--ta320) 62%, transparent 97%);
}
.fpanel-hdr {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 11px 20px 10px;
    border-bottom: 1px solid var(--ta060);
    background: linear-gradient(90deg, var(--ta055), rgba(var(--rgb-run-l),.025), transparent);
}
.fpanel-title {
    display: flex;
    align-items: center;
    gap: 9px;
    font-family: var(--mono);
    font-size: .65rem;
    font-weight: 800;
    letter-spacing: .16em;
    text-transform: uppercase;
    color: var(--accent);
}
.fpanel-dot {
    width: 7px; height: 7px;
    border-radius: 50%;
    background: var(--accent);
    box-shadow: none;
    flex-shrink: 0;
    animation: none;
}
@keyframes pulse-dot {
    0%, 100% { opacity: 1; box-shadow: none; }
    50% { opacity: .55; box-shadow: none; }
}
.fpanel-meta { display: flex; align-items: center; gap: 8px; }
.fpanel-badge {
    font-family: var(--mono);
    font-size: .62rem;
    font-weight: 700;
    letter-spacing: .07em;
    color: var(--txt-3);
    padding: 3px 10px;
    border: 1px solid var(--ta100);
    border-radius: 999px;
    background: var(--ta040);
    transition: all .2s ease;
}
.fpanel-badge-warn {
    color: var(--c-crit-l);
    border-color: rgba(var(--rgb-crit-l),.28);
    background: rgba(var(--rgb-crit-l),.06);
    box-shadow: none;
}
.fpanel-body {
    padding: 16px 18px 18px;
    background: transparent;
}

/* ── GROUP LAYOUT ── */
.fgroup {
    display: flex;
    flex-direction: column;
    gap: 0;
}
.fgroup-hdr {
    display: flex;
    align-items: center;
    gap: 6px;
    font-family: var(--mono);
    font-size: .60rem;
    font-weight: 800;
    letter-spacing: .14em;
    text-transform: uppercase;
    color: var(--t-n6);
    margin-bottom: 9px;
    padding-bottom: 7px;
    border-bottom: 1px solid var(--ta060);
}
.fgroup-icon { font-size: .75rem; line-height: 1; }
.fchip-row {
    display: flex;
    flex-direction: column;
    gap: 4px;
}
.fchip-row-wrap {
    flex-direction: row !important;
    flex-wrap: wrap !important;
    gap: 4px !important;
    align-content: flex-start;
}
.fgroup-row2 {
    margin-top: 14px;
    padding-top: 14px;
    border-top: 1px solid var(--ta060);
}

/* ── CHIP (wraps each st.checkbox) ── */
.fchip {
    position: relative;
    border-radius: 7px;
    transition: background .15s ease, border-color .15s ease, box-shadow .18s ease, transform .12s ease;
}
/* ON state — chip lit */
.fchip-on {
    background: linear-gradient(135deg, rgba(var(--chip-rgb, 255,183,0), .14), rgba(var(--chip-rgb, 255,183,0), .06));
    background: color-mix(in srgb, var(--chip-color, var(--c-acc)) 12%, transparent);
    border: 1px solid color-mix(in srgb, var(--chip-color, var(--c-acc)) 40%, transparent);
    box-shadow: 0 0 0 0 transparent;
}
/* OFF state — chip dimmed/inactive */
.fchip-off {
    background: rgba(var(--rgb-n1),.45);
    border: 1px solid rgba(var(--rgb-n9),.07);
    opacity: .52;
}
.fchip-off:hover { opacity: .75; border-color: rgba(var(--rgb-n9),.18); }
.fchip-on:hover {
    border-color: color-mix(in srgb, var(--chip-color, var(--c-acc)) 70%, transparent) !important;
    box-shadow: none;
    transform: translateY(-1px);
}

/* ── HIJACK st.checkbox inside .fchip ── */
.fchip .stCheckbox {
    margin: 0 !important;
    padding: 0 !important;
}
.fchip .stCheckbox > label {
    display: flex !important;
    align-items: center !important;
    gap: 8px !important;
    padding: 7px 10px 7px 9px !important;
    margin: 0 !important;
    cursor: pointer !important;
    width: 100% !important;
    border-radius: 6px !important;
}
/* Hide native Streamlit checkbox box; make entire chip clickable */
.fchip .stCheckbox [data-testid="stCheckbox"] {
    width: 100% !important;
}
.fchip .stCheckbox > label > span:first-child {
    /* the actual checkbox input area */
    width: 14px !important;
    height: 14px !important;
    border-radius: 4px !important;
    border: 1.5px solid color-mix(in srgb, var(--chip-color, var(--c-acc)) 50%, rgba(var(--rgb-n9),.3)) !important;
    background: rgba(var(--rgb-n0),.6) !important;
    flex-shrink: 0 !important;
    transition: background .15s ease, border-color .15s ease !important;
}
.fchip-on .stCheckbox > label > span:first-child {
    background: color-mix(in srgb, var(--chip-color, var(--c-acc)) 85%, white 15%) !important;
    border-color: var(--chip-color, var(--c-acc)) !important;
    box-shadow: none;
}
.fchip .stCheckbox > label > div {
    font-family: var(--mono) !important;
    font-size: .69rem !important;
    font-weight: 700 !important;
    letter-spacing: .06em !important;
    text-transform: uppercase !important;
    color: color-mix(in srgb, var(--chip-color, var(--c-acc)) 90%, white 10%) !important;
    transition: color .15s ease !important;
    line-height: 1.2 !important;
}
.fchip-off .stCheckbox > label > div {
    color: var(--t-n5) !important;
}
/* Remove Streamlit's default focus ring clutter */
.fchip .stCheckbox > label:focus-within {
    outline: none !important;
    box-shadow: none !important;
}

/* ── SEARCH + DATE INPUTS inside fpanel-body ── */
.filter-search-wrap .stTextInput > div {
    background: rgba(var(--rgb-n0),.75) !important;
    border: 1px solid var(--ta140) !important;
    border-radius: 8px !important;
    height: 40px !important;
    transition: border-color .2s ease, box-shadow .2s ease !important;
}
.filter-search-wrap .stTextInput > div:focus-within {
    border-color: var(--ta520) !important;
    box-shadow: 0 0 0 3px var(--ta090);
}
.filter-search-wrap .stTextInput input::placeholder { color: var(--t-n5) !important; }

/* date inputs in filter row */
.fpanel-body [data-testid="stDateInput"] > div {
    background: rgba(var(--rgb-n0),.75) !important;
    border: 1px solid var(--ta140) !important;
    border-radius: 8px !important;
    height: 40px !important;
}
.fpanel-body [data-testid="stDateInput"] > div:focus-within {
    border-color: var(--ta520) !important;
    box-shadow: 0 0 0 3px var(--ta090) !important;
}

/* Reset button inside filter */
.filter-reset-btn .stButton > button {
    background: rgba(var(--rgb-crit-l),.06) !important;
    border: 1px solid rgba(var(--rgb-crit-l),.22) !important;
    border-radius: 8px !important;
    color: rgba(var(--rgb-c-crit-l),.7) !important;
    font-family: var(--mono) !important;
    font-size: .66rem !important;
    font-weight: 700 !important;
    letter-spacing: .08em !important;
    height: 40px !important;
    padding: 0 10px !important;
    transition: all .18s ease !important;
    width: 100% !important;
}
.filter-reset-btn .stButton > button:hover {
    background: rgba(var(--rgb-crit-l),.14) !important;
    border-color: rgba(var(--rgb-crit-l),.45) !important;
    color: var(--c-crit-l) !important;
    box-shadow: none;
}

/* ══════════════════════════════════════════════════════════
   FILTER PANEL v3 — MULTISELECT TECH OVERHAUL
   ══════════════════════════════════════════════════════════ */

/* ── Container do painel de filtros ── */
[data-testid="stVerticalBlockBorderWrapper"]:has(.stMultiSelect) {
    background: linear-gradient(145deg, rgba(var(--rgb-sup2-a),.98), rgba(var(--rgb-sup2-b),.96)) !important;
    border: 1px solid var(--ta180) !important;
    border-radius: 14px !important;
    padding: 4px !important;
    box-shadow: 0 0 0 1px var(--ta050), inset 0 1px 0 var(--ta080), 0 24px 56px var(--shc-1) !important;
    position: relative !important;
    overflow: hidden !important;
}
[data-testid="stVerticalBlockBorderWrapper"]:has(.stMultiSelect)::before {
    content: "" !important;
    position: absolute !important;
    top: 0; left: 0; right: 0; height: 1px !important;
    background: linear-gradient(90deg,
        transparent 2%,
        var(--ta650) 35%,
        rgba(var(--rgb-run-l),.45) 65%,
        transparent 98%) !important;
    pointer-events: none !important;
}

/* ── Label dos multiselects: estilo monospace neon ── */
.stMultiSelect label {
    font-family: 'IBM Plex Mono', ui-monospace, monospace !important;
    font-size: .62rem !important;
    font-weight: 800 !important;
    letter-spacing: .15em !important;
    text-transform: uppercase !important;
    color: var(--t-n6) !important;
    margin-bottom: 5px !important;
    display: flex !important;
    align-items: center !important;
    gap: 5px !important;
}

/* ── Wrapper externo do multiselect ── */
.stMultiSelect > div {
    background: rgba(var(--rgb-n0),.82) !important;
    border: 1px solid var(--ta160) !important;
    border-radius: 9px !important;
    min-height: 38px !important;
    box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.03), inset 0 -1px 0 var(--shc-2) !important;
    transition: border-color .2s ease, box-shadow .22s ease, background .2s ease !important;
    position: relative !important;
}
.stMultiSelect > div:focus-within {
    border-color: var(--ta550) !important;
    background: rgba(var(--rgb-n1),.92) !important;
    box-shadow: 0 0 0 3px var(--ta100), inset 0 1px 0 var(--ta060) !important;
}

/* ── Área interna com as tags selecionadas ── */
.stMultiSelect [data-baseweb="select"] > div:first-child {
    background: transparent !important;
    padding: 4px 8px !important;
    gap: 4px !important;
    flex-wrap: wrap !important;
    min-height: 36px !important;
    height: auto !important;
    align-items: center !important;
    overflow: visible !important;
}
/* Texto do placeholder ("Choose options") e do valor: o baseweb costuma
   herdar um line-height fixo pensado pra fonte padrão; sem isso, com
   'Archivo'/'IBM Plex Mono', o topo das letras fica cortado pelo container. */
.stMultiSelect [data-baseweb="select"] > div:first-child > div {
    line-height: 1.5 !important;
    overflow: visible !important;
    display: flex !important;
    align-items: center !important;
}
.stMultiSelect [data-baseweb="select"] [class*="placeholder"] {
    line-height: 1.5 !important;
    overflow: visible !important;
    color: var(--t-n6) !important;
    font-family: 'Archivo', sans-serif !important;
    font-size: .8rem !important;
}

/* ── Tags individuais (itens selecionados) ── */
.stMultiSelect [data-baseweb="tag"] {
    background: linear-gradient(135deg,
        var(--ta180) 0%,
        var(--ta080) 100%) !important;
    border: 1px solid var(--ta380) !important;
    border-radius: 6px !important;
    height: 24px !important;
    padding: 0 8px 0 9px !important;
    box-shadow: inset 0 1px 0 var(--ta120) !important;
    transition: border-color .15s ease, box-shadow .15s ease, transform .12s ease !important;
    max-width: 160px !important;
    position: relative !important;
}
.stMultiSelect [data-baseweb="tag"]:hover {
    border-color: var(--ta650) !important;
    box-shadow: none;
    transform: translateY(-1px) !important;
}

/* Texto dentro da tag */
.stMultiSelect [data-baseweb="tag"] span:first-child {
    font-family: 'IBM Plex Mono', ui-monospace, monospace !important;
    font-size: .63rem !important;
    font-weight: 700 !important;
    letter-spacing: .07em !important;
    text-transform: uppercase !important;
    color: var(--c-acc-l) !important;
    white-space: nowrap !important;
    overflow: hidden !important;
    text-overflow: ellipsis !important;
}

/* Botão X (remover) da tag */
.stMultiSelect [data-baseweb="tag"] [role="presentation"] {
    color: rgba(var(--rgb-c-acc),.5) !important;
    margin-left: 4px !important;
    font-size: .7rem !important;
    transition: color .15s ease !important;
    background: transparent !important;
    border: none !important;
    cursor: pointer !important;
    display: flex !important;
    align-items: center !important;
}
.stMultiSelect [data-baseweb="tag"] [role="presentation"]:hover {
    color: var(--c-crit-l) !important;
}

/* Ícone X via SVG override — make it visible */
.stMultiSelect [data-baseweb="tag"] svg {
    fill: var(--ta550) !important;
    width: 10px !important;
    height: 10px !important;
    transition: fill .15s ease !important;
}
.stMultiSelect [data-baseweb="tag"]:hover svg {
    fill: var(--f-crit-l) !important;
}

/* Placeholder "Todos / Todas" */
.stMultiSelect [data-baseweb="select"] input::placeholder {
    color: var(--t-n5) !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: .75rem !important;
    letter-spacing: .06em !important;
}
.stMultiSelect [data-baseweb="select"] input {
    color: var(--t-n11) !important;
    font-family: 'Archivo', sans-serif !important;
    font-size: .8rem !important;
    caret-color: var(--c-acc) !important;
    background: transparent !important;
    border: none !important;
    outline: none !important;
    min-width: 40px !important;
}

/* Seta dropdown */
.stMultiSelect [data-baseweb="select"] svg[data-icon="chevron-down"],
.stMultiSelect [data-baseweb="select"] > div:last-child svg {
    fill: var(--ta400) !important;
    transition: fill .2s ease, transform .2s ease !important;
}
.stMultiSelect > div:focus-within svg[data-icon="chevron-down"],
.stMultiSelect > div:focus-within > div:last-child svg {
    fill: var(--ta800) !important;
    transform: rotate(180deg) !important;
}

/* ── Dropdown menu flutuante ── */
[data-baseweb="popover"] [data-baseweb="menu"],
[data-baseweb="menu"] {
    background: linear-gradient(145deg, rgba(var(--rgb-sup2-a),.99), rgba(var(--rgb-sup2-b),.98)) !important;
    border: 1px solid var(--ta220) !important;
    border-radius: 10px !important;
    box-shadow: 0 0 0 1px var(--ta060), 0 24px 60px var(--shc-1);
    padding: 4px !important;
    
    overflow: hidden !important;
}

/* Opções individuais do dropdown */
[data-baseweb="menu"] li,
[data-baseweb="option"] {
    background: transparent !important;
    border-radius: 6px !important;
    padding: 7px 12px !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: .72rem !important;
    font-weight: 600 !important;
    letter-spacing: .05em !important;
    color: var(--t-n8) !important;
    text-transform: uppercase !important;
    transition: background .14s ease, color .14s ease !important;
    cursor: pointer !important;
    border: 1px solid transparent !important;
    margin-bottom: 2px !important;
}
[data-baseweb="menu"] li:hover,
[data-baseweb="option"]:hover {
    background: var(--ta100) !important;
    border-color: var(--ta220) !important;
    color: var(--t-n11) !important;
}
/* Opção já selecionada */
[data-baseweb="option"][aria-selected="true"] {
    background: linear-gradient(135deg, var(--ta180), var(--ta060)) !important;
    border-color: var(--ta300) !important;
    color: var(--c-acc-l) !important;
}

/* ── Botão ↺ reset — estilo neon vermelho ── */
[data-testid="stVerticalBlockBorderWrapper"]:has(.stMultiSelect) .stButton > button {
    background: rgba(var(--rgb-crit-l),.06) !important;
    border: 1px solid rgba(var(--rgb-crit-l),.28) !important;
    border-radius: 8px !important;
    color: rgba(var(--rgb-c-crit-l),.75) !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: .85rem !important;
    font-weight: 700 !important;
    height: 38px !important;
    width: 100% !important;
    padding: 0 !important;
    transition: all .18s ease !important;
    box-shadow: none !important;
}
[data-testid="stVerticalBlockBorderWrapper"]:has(.stMultiSelect) .stButton > button:hover {
    background: rgba(var(--rgb-crit-l),.14) !important;
    border-color: rgba(var(--rgb-crit-l),.55) !important;
    color: var(--c-crit-l) !important;
    box-shadow: none;
    transform: none !important;
}

/* ── Date inputs dentro do painel de filtros ── */
[data-testid="stVerticalBlockBorderWrapper"]:has(.stMultiSelect) [data-testid="stDateInput"] > div,
[data-testid="stVerticalBlockBorderWrapper"]:has(.stMultiSelect) [data-baseweb="input"] {
    background: rgba(var(--rgb-n0),.82) !important;
    border: 1px solid var(--ta160) !important;
    border-radius: 9px !important;
    min-height: 38px !important;
    box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.03) !important;
    transition: border-color .2s ease, box-shadow .2s ease !important;
}
[data-testid="stVerticalBlockBorderWrapper"]:has(.stMultiSelect) [data-testid="stDateInput"] > div:focus-within,
[data-testid="stVerticalBlockBorderWrapper"]:has(.stMultiSelect) [data-baseweb="input"]:focus-within {
    border-color: var(--ta500) !important;
    box-shadow: 0 0 0 3px var(--ta090);
}

/* ── Busca livre abaixo do painel ── */
.stTextInput:has(#f_txt_mgr), .stTextInput:has(#f_txt_cp) {
    margin-top: 6px !important;
}
.stTextInput:has(#f_txt_mgr) > div,
.stTextInput:has(#f_txt_cp) > div {
    background: rgba(var(--rgb-n0),.85) !important;
    border: 1px solid var(--ta140) !important;
    border-radius: 9px !important;
    box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.025) !important;
    transition: border-color .2s ease, box-shadow .2s ease !important;
}
.stTextInput:has(#f_txt_mgr) > div:focus-within,
.stTextInput:has(#f_txt_cp) > div:focus-within {
    border-color: var(--ta520) !important;
    box-shadow: 0 0 0 3px var(--ta090);
}
.stTextInput:has(#f_txt_mgr) label,
.stTextInput:has(#f_txt_cp) label {
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: .62rem !important;
    font-weight: 800 !important;
    letter-spacing: .15em !important;
    text-transform: uppercase !important;
    color: var(--t-n6) !important;
}

/* Keep old filter-panel-wrap styles for any remaining usage */
.filter-panel-wrap {
    position: relative;
    background: linear-gradient(145deg, rgba(var(--rgb-n1),.97), rgba(var(--rgb-n1),.94));
    border: 1px solid var(--ta160);
    border-radius: 14px;
    padding: 0;
    overflow: hidden;
    box-shadow: 0 0 0 1px var(--ta040), 0 24px 56px var(--shc-2);
}
.filter-group-label {
    display: flex;
    align-items: center;
    gap: 7px;
    font-family: var(--mono);
    font-size: .64rem;
    font-weight: 700;
    letter-spacing: .12em;
    text-transform: uppercase;
    color: var(--t-n7);
    margin-bottom: 8px;
    padding-bottom: 6px;
    border-bottom: 1px solid var(--ta050);
}
.filter-group-icon {
    width: 18px; height: 18px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    border-radius: 5px;
    font-size: .7rem;
}
.fgi-prio  { background: rgba(var(--rgb-run),.12);  border: 1px solid rgba(var(--rgb-run),.22); }
.fgi-status{ background: rgba(var(--rgb-ok),.12);   border: 1px solid rgba(var(--rgb-ok),.22); }
.fgi-setor { background: rgba(var(--rgb-run-l),.12);  border: 1px solid rgba(var(--rgb-run-l),.22); }
.fgi-busca { background: var(--ta100);   border: 1px solid var(--ta220); }

/* ── BUTTONS ── */
.stButton > button[kind="primary"] {
    position: relative !important;
    overflow: hidden !important;
    background: linear-gradient(135deg, var(--f-acc) 0%, var(--f-acc) 55%, var(--f-acc) 100%) !important;
    border: 1px solid rgba(var(--rgb-acc-l),.45) !important;
    border-radius: 11px !important;
    color: var(--t-n1) !important;
    font-family: var(--mono) !important;
    font-weight: 700 !important;
    font-size: .8rem !important;
    letter-spacing: .10em !important;
    text-transform: uppercase !important;
    padding: 4px 0 !important;
    box-shadow: 0 8px 22px var(--ta280), inset 0 1px 0 rgba(var(--rgb-lift),.35) !important;
    transition: transform .18s ease, box-shadow .22s ease, filter .2s ease !important;
}
/* brilho sutil que desliza ao passar o mouse */
.stButton > button[kind="primary"]::after {
    content: "" !important;
    position: absolute !important;
    top: 0; left: -140%;
    width: 55%; height: 100%;
    background: linear-gradient(115deg, transparent 0%, rgba(var(--rgb-lift),.40) 50%, transparent 100%) !important;
    transform: skewX(-18deg);
    pointer-events: none;
    opacity: 0;
}
.stButton > button[kind="primary"]:hover::after {
    opacity: 1;
    animation: btnSheen 1.05s ease-out !important;
}
@keyframes btnSheen { 0% { left: -140%; } 100% { left: 140%; } }
.stButton > button[kind="primary"]:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 14px 32px var(--ta420), inset 0 1px 0 rgba(var(--rgb-lift),.45) !important;
    filter: brightness(1.05) !important;
}
.stButton > button[kind="primary"]:active { transform: translateY(0) !important; }
.stButton > button[kind="primary"] *,
.stButton button[kind="primary"] p,
[data-testid="stButton"] button[kind="primary"] p,
[data-testid="stBaseButton-primary"] p,
[data-testid="stBaseButton-primary"] {
    color: var(--act-ink) !important;
    filter: none !important;
}
.stButton > button[kind="secondary"] {
    background: linear-gradient(180deg, rgba(var(--rgb-sup-a),.88), rgba(var(--rgb-sup-b),.82)) !important;
    border: 1px solid rgba(var(--rgb-n9),.16) !important;
    border-radius: 8px !important;
    color: var(--t-n10) !important;
    font-family: var(--font) !important;
    font-weight: 650 !important;
    transition: all .2s ease !important;
}
.stButton > button[kind="secondary"]:hover {
    border-color: var(--border-hi) !important;
    color: var(--t-n11) !important;
    box-shadow: none;
}
.stFormSubmitButton > button {
    background: linear-gradient(135deg, var(--f-acc) 0%, var(--f-acc) 55%, var(--f-acc) 100%) !important;
    border: 1px solid rgba(var(--rgb-acc-l),.45) !important;
    border-radius: 11px !important;
    color: var(--t-n1) !important;
    font-family: var(--mono) !important;
    font-weight: 700 !important;
    font-size: .8rem !important;
    letter-spacing: .10em !important;
    text-transform: uppercase !important;
    padding: 4px 0 !important;
    box-shadow: 0 8px 22px var(--ta280), inset 0 1px 0 rgba(var(--rgb-lift),.35) !important;
    transition: transform .18s ease, box-shadow .22s ease, filter .2s ease !important;
}
.stFormSubmitButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 14px 32px var(--ta420), inset 0 1px 0 rgba(var(--rgb-lift),.45) !important;
    filter: brightness(1.05) !important;
}
/* O rótulo do botão vive num <p> dentro de um <div>, e o Streamlit dá cor
   própria a ele. Fixar um token aqui quebra sempre que o FUNDO do botão muda
   com o tema — era o caso do AUTENTICAR: fundo preto no claro com texto
   `--t-n0`, também escuro. `inherit` faz o rótulo seguir a cor que o próprio
   botão já calculou, que é a única que conhece o fundo dele. */
.stFormSubmitButton > button *,
.stFormSubmitButton button p,
[data-testid="stFormSubmitButton"] button p,
[data-testid="stBaseButton-primary"] p {
    color: inherit !important;
    filter: none !important;
}

/* ── TABS ── */
.stTabs [data-baseweb="tab-list"] {
    background: rgba(var(--rgb-n1),.76) !important;
    gap: 6px !important;
    border: 1px solid var(--ta130) !important;
    border-radius: 10px !important;
    padding: 6px !important;
    box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.035), var(--shadow-soft) !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    border-radius: 8px !important;
    color: var(--txt-2) !important;
    font-family: var(--font) !important;
    font-weight: 700 !important;
    font-size: .85rem !important;
    border: 1px solid transparent !important;
    transition: all .15s ease !important;
}
.stTabs [data-baseweb="tab"]:hover {
    color: var(--txt-1) !important;
    background: var(--ta080) !important;
    border-color: var(--ta160) !important;
}
.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg, var(--ta170), rgba(var(--rgb-run-l),.07)) !important;
    color: var(--accent) !important;
    border-color: var(--ta350) !important;
    box-shadow: none;
}

/* ── DATAFRAME ── */
[data-testid="stDataFrame"] {
    border-radius: 10px !important;
    overflow: hidden !important;
    border: 1px solid var(--ta140) !important;
    box-shadow: var(--shadow-soft) !important;
    background: rgba(var(--rgb-n1),.84) !important;
}

/* ── METRICS ── */
[data-testid="stMetric"] {
    position: relative !important;
    overflow: hidden !important;
    background:
        linear-gradient(145deg, rgba(var(--rgb-sup-a),.92), rgba(var(--rgb-sup-b),.82)) !important;
    border: 1px solid var(--border-soft) !important;
    border-radius: var(--radius) !important;
    padding: 1.2rem !important;
    box-shadow: var(--shadow-soft) !important;
}
[data-testid="stMetric"]::before {
    content: "";
    position: absolute;
    left: 0; top: 0; bottom: 0;
    width: 3px;
    background: var(--accent);
    box-shadow: none;
}
[data-testid="stMetricLabel"] {
    font-size: .68rem !important;
    font-family: var(--mono) !important;
    text-transform: uppercase !important;
    letter-spacing: .1em !important;
    color: var(--t-n9) !important;
    font-weight: 700 !important;
}
[data-testid="stMetricValue"] {
    font-family: var(--mono) !important;
    font-size: 2rem !important;
    color: var(--txt-1) !important;
}

/* ── ALERTS ── */
[data-testid="stAlert"] {
    border-radius: 10px !important;
    border: 1px solid var(--ta140) !important;
    border-left-width: 4px !important;
    font-size: .875rem !important;
    background: rgba(var(--rgb-n1),.82) !important;
    box-shadow: var(--shadow-soft) !important;
}

/* ── FILE UPLOADER ── */
[data-testid="stFileUploader"] {
    border: 1px dashed var(--ta260) !important;
    border-radius: 10px !important;
    background:
        linear-gradient(135deg, var(--ta060), rgba(var(--rgb-run-l),.035)),
        rgba(var(--rgb-n1),.72) !important;
    transition: all .2s ease !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: var(--accent) !important;
    background: var(--ta030) !important;
}

/* ── DOWNLOAD BUTTON ── */
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(180deg, rgba(var(--rgb-sup-a),.88), rgba(var(--rgb-sup-b),.82)) !important;
    border: 1px solid rgba(var(--rgb-n9),.16) !important;
    border-radius: 8px !important;
    color: var(--t-n10) !important;
    font-family: var(--font) !important;
    font-weight: 650 !important;
    font-size: .8rem !important;
    transition: all .2s ease !important;
}
[data-testid="stDownloadButton"] > button:hover {
    border-color: var(--border-hi) !important;
    color: var(--accent) !important;
    box-shadow: none;
}

/* ── EXPANDERS / DIALOGS / STATUS ── */
[data-testid="stExpander"] {
    border: 1px solid var(--ta140) !important;
    border-radius: 10px !important;
    overflow: hidden !important;
    background: rgba(var(--rgb-n1),.78) !important;
    box-shadow: var(--shadow-soft) !important;
}
[data-testid="stExpander"] details {
    background: transparent !important;
}
[data-testid="stExpander"] summary {
    font-weight: 700 !important;
    color: var(--txt-1) !important;
}
/* Janela modal. A estrutura no 1.56 tem TRÊS níveis e confundi-los pinta a
   tela toda: [stDialog] é a camada fixa de fundo, o filho direto é um
   invólucro que cobre a viewport inteira, e só [role="dialog"] é o painel.
   Com o degradê do painel aplicado ao invólucro, o "véu" virava uma chapa
   clara sobre a página — parecia que o fundo tinha sido apagado, não
   escurecido. */
[data-testid="stDialog"] > div {
    /* o véu: mais forte no claro, onde um escurecimento leve não separa nada */
    background: var(--veu) !important;
}
[data-testid="stDialog"] [role="dialog"] {
    background:
        linear-gradient(145deg, rgba(var(--rgb-sup-a),.99), rgba(var(--rgb-sup-b),.97)) !important;
    border: 1px solid var(--ta220) !important;
    border-radius: 12px !important;
    box-shadow: 0 28px 80px var(--shc-1), 0 0 0 1px var(--ta060) !important;
}
[data-testid="stStatusWidget"] {
    border-radius: 10px !important;
}

/* ── SPINNER ── */
.stSpinner > div { border-top-color: var(--accent) !important; }

/* ── CUSTOM COMPONENTS ── */
.sidebar-header {
    position: relative;
    margin: 0 -16px 16px;
    padding: 28px 16px 18px;
    overflow: hidden;
    border-bottom: 1px solid var(--ta120);
    background: linear-gradient(180deg, rgba(var(--rgb-sup-a),.88), rgba(var(--rgb-sup-b),.18));
    isolation: isolate;
}
.sidebar-header::before {
    content: "";
    position: absolute;
    inset: 0;
    pointer-events: none;
    background: linear-gradient(90deg, transparent, var(--ta110), transparent) 0 0 / 100% 1px no-repeat;
}
.sidebar-header::after {
    content: "";
    position: absolute;
    inset: -40% -60%;
    pointer-events: none;
    z-index: 0;
    background: linear-gradient(180deg, transparent, var(--ta060), transparent);
    transform: translate3d(-2%, 0, 0);
    animation: none;
    mask-image: radial-gradient(circle at 50% 30%, rgba(0,0,0,.9), transparent 68%);
}
.sidebar-header:hover::after {
    opacity: .95;
}
.brand-lockup {
    position: relative;
    z-index: 1;
    display: grid;
    justify-items: center;
    gap: 12px;
    animation: none;
}
/* ── HEADER PARTICLES ── */
.header-particles {
    position: absolute; inset: 0;
    pointer-events: none; overflow: hidden; z-index: 0;
}
.header-particles::before {
    content: "";
    position: absolute;
    width: 2px; height: 2px; border-radius: 50%;
    background: var(--ta700);
    box-shadow: 22px 16px 0 var(--ta400), 44px  7px 0 rgba(var(--rgb-run-l),.32), 78px 22px 0 var(--ta280), 108px 11px 0 var(--ta440), 138px 20px 0 rgba(var(--rgb-run-l),.28), 168px  5px 0 var(--ta380), 198px 24px 0 var(--ta220), 34px 44px 0 var(--ta200), 68px 38px 0 rgba(var(--rgb-run-l),.24), 104px 48px 0 var(--ta280), 154px 36px 0 var(--ta200), 62px 58px 0 rgba(var(--rgb-run-l),.18);
    animation: none;
}
.header-particles::after {
    content: "";
    position: absolute;
    width: 1px; height: 1px; border-radius: 50%;
    background: rgba(var(--rgb-run-l),.6);
    box-shadow: 14px 32px 0 var(--ta280), 54px 14px 0 rgba(var(--rgb-run-l),.38), 88px 36px 0 var(--ta220), 124px  8px 0 var(--ta320), 162px 28px 0 rgba(var(--rgb-run-l),.28), 196px 18px 0 var(--ta360), 228px 42px 0 var(--ta180), 40px 54px 0 rgba(var(--rgb-run-l),.22);
    animation: none;
}
@keyframes particleFloat {
    0%,100% { transform: translateY(0)    translateX(0);    opacity: .65; }
    25%      { transform: translateY(-6px) translateX(3px);  opacity: .90; }
    50%      { transform: translateY(-3px) translateX(-2px); opacity: .55; }
    75%      { transform: translateY(-7px) translateX(4px);  opacity: .82; }
}

/* ── LOGO MARK — signal pulse + orbit ── */
.logo-ring {
    position: relative;
    width: clamp(54px, 8vw, 72px);
    height: clamp(54px, 8vw, 72px);
    display: flex; align-items: center; justify-content: center;
    flex-shrink: 0;
}
/* ping externo */
.logo-ring::before {
    content: "";
    position: absolute; inset: 0;
    border-radius: 14px;
    border: 1px solid var(--ta220);
    animation: none;
}
.logo-ring::after {
    content: "";
    position: absolute; inset: 6px;
    border-radius: 10px;
    border: 1px solid var(--ta100);
    animation: none;
}
@keyframes logoPing {
    0%   { opacity: .7; transform: scale(1);    }
    80%  { opacity: 0;  transform: scale(1.20); }
    100% { opacity: 0;  transform: scale(1.20); }
}

/* ── ORBIT — anel satélite azul (horário) ── */
.logo-orbit {
    position: absolute; inset: -10px;
    border-radius: 20px;
    pointer-events: none;
    animation: none;
    z-index: 2;
}
.logo-orbit::before {
    content: "";
    position: absolute; inset: 0;
    border-radius: 20px;
    border: 1.5px solid transparent;
    border-top-color: var(--ta620);
    border-right-color: var(--ta180);
}
.logo-orbit::after {
    content: "";
    position: absolute;
    top: -3px; left: 50%;
    width: 6px; height: 6px; margin-left: -3px;
    border-radius: 50%;
    background: var(--f-acc);
    box-shadow: none;
}
/* ── ORBIT 2 — anel satélite roxo (anti-horário) ── */
.logo-orbit-2 {
    position: absolute; inset: -5px;
    border-radius: 16px;
    pointer-events: none;
    animation: none;
    z-index: 2;
}
.logo-orbit-2::before {
    content: "";
    position: absolute; inset: 0;
    border-radius: 16px;
    border: 1px solid transparent;
    border-bottom-color: rgba(var(--rgb-run-l),.50);
    border-left-color: rgba(var(--rgb-run-l),.14);
}
.logo-orbit-2::after {
    content: "";
    position: absolute;
    bottom: -2px; left: 50%;
    width: 4px; height: 4px; margin-left: -2px;
    border-radius: 50%;
    background: var(--f-run-l);
    box-shadow: none;
}
@keyframes orbitSpin {
    from { transform: rotate(0deg); }
    to   { transform: rotate(360deg); }
}

.logo-mark {
    position: relative; z-index: 1;
    width: clamp(40px, 5.6vw, 56px);
    height: clamp(40px, 5.6vw, 56px);
    display: grid; place-items: center;
    background: linear-gradient(135deg, var(--ta140), var(--ta100));
    border-radius: 12px;
    border: 1px solid var(--ta250);
    box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.06);
    overflow: hidden;
    /* Camada GPU própria: confina o repaint do box-shadow pulsante a este layer,
       sem cascatear para a sidebar durante reruns. */
    transform: translateZ(0);
    will-change: transform;
    animation: none;
}
/* scan line que varre o logo de cima a baixo */
.logo-mark::before {
    content: "";
    position: absolute;
    left: 0; right: 0; height: 40%;
    background: linear-gradient(
        180deg,
        transparent 0%,
        var(--ta160) 40%,
        var(--ta300) 50%,
        var(--ta160) 60%,
        transparent 100%
    );
    animation: none;
    pointer-events: none; z-index: 3;
}
@keyframes logoScan {
    0%   { top: -40%; opacity: 0; }
    8%   { opacity: 1; }
    85%  { top: 100%;  opacity: .75; }
    100% { top: 100%;  opacity: 0; }
}
@keyframes markGlow {
    0%,100% {
        box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.06);
        border-color: var(--ta250);
        background: linear-gradient(135deg, var(--ta140), var(--ta100));
    }
    38% {
        box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.10);
        border-color: var(--ta480);
        background: linear-gradient(135deg, var(--ta220), var(--ta140));
    }
    68% {
        box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.10);
        border-color: rgba(var(--rgb-run-l),.40);
        background: linear-gradient(135deg, var(--ta160), rgba(var(--rgb-run-l),.12));
    }
}
.logo-mark svg { width: 60%; height: 60%; overflow: visible; position: relative; z-index: 1; }

/* ── CIRCUIT NODE ICON — pulso de rede tecnológico ── */
.circuit-node {
    animation: none;
}
@keyframes nodeGlow {
    0%,100% { opacity: .85; }
    50%      { opacity: 1; }
}
.circuit-pulse-ring {
    animation: none;
    transform-origin: center;
}
.circuit-pulse-ring-2 {
    animation: none;
    transform-origin: center;
}
@keyframes pulseRingExpand {
    0%   { opacity: .9;  transform: scale(.55); }
    70%  { opacity: 0;   transform: scale(1.0); }
    100% { opacity: 0;   transform: scale(1.0); }
}
.circuit-line-h {
    stroke-dasharray: 16;
    stroke-dashoffset: 16;
    animation: none;
}
.circuit-line-v {
    stroke-dasharray: 16;
    stroke-dashoffset: 16;
    animation: none;
}
.circuit-line-d {
    stroke-dasharray: 10;
    stroke-dashoffset: 10;
    animation: none;
}
@keyframes drawLine {
    0%,100% { stroke-dashoffset: 16; opacity: 0; }
    20%      { opacity: 1; }
    55%      { stroke-dashoffset: 0;  opacity: .9; }
    85%      { stroke-dashoffset: 0;  opacity: .3; }
}
.circuit-dot-1 { animation: none; }
.circuit-dot-2 { animation: none; }
.circuit-dot-3 { animation: none; }
.circuit-dot-4 { animation: none; }
@keyframes dotBlink {
    0%,100% { opacity: .18; r: 1.2; }
    45%      { opacity: 1;   r: 2.0; filter:,1)); }
    70%      { opacity: .7;  r: 1.6; }
}
.sidebar-title {
    color: var(--txt-1);
    font-family: var(--font);
    font-weight: 800;
    font-size: clamp(.86rem, 1.6vw, 1.02rem);
    letter-spacing: .14em;
    text-transform: uppercase;
    text-shadow: none;
    text-align: center;
    background: linear-gradient(90deg, var(--n11) 0%, var(--ta950) 35%, var(--n11) 70%, rgba(var(--rgb-run-l),.9) 100%);
    background-size: 220% 100%;
    -webkit-background-clip: text;
    background-clip: text;
    -webkit-text-fill-color: transparent;
    animation: none;
}
.sidebar-sub {
    color: var(--t-n8);
    font-size: clamp(.62rem, 1.3vw, .70rem);
    letter-spacing: .08em;
    margin-top: 4px;
    text-align: center;
    opacity: .92;
}
.status-strip {
    position: relative;
    z-index: 1;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    margin-top: 16px;
    padding: 8px 10px;
    border: 1px solid var(--ta220);
    box-shadow: var(--shadow-soft);
    border-radius: 10px;
    background: rgba(var(--rgb-n0),.45);
    overflow: hidden;
    color: var(--t-n9);
    font-size: .64rem;
    font-family: var(--mono);
    text-transform: uppercase;
    letter-spacing: .08em;
    transition: transform .18s ease, border-color .18s ease, box-shadow .18s ease;
}
.status-strip:hover {
    transform: translateY(-1px);
    border-color: var(--ta350);
    box-shadow: 0 18px 44px var(--shc-2);
}
.status-strip::after {
    content: "";
    position: absolute;
    inset: 0;
    pointer-events: none;
    background: linear-gradient(
        90deg,
        transparent 0%,
        var(--ta100) 35%,
        rgba(var(--rgb-run-l),.08) 50%,
        var(--ta100) 65%,
        transparent 100%
    );
    transform: translateX(-110%);
    opacity: .55;
    animation: none;
    mask-image: linear-gradient(180deg, transparent, rgba(0,0,0,.95), transparent);
}
.status-strip span {
    display: inline-flex;
    align-items: center;
    gap: 7px;
    min-width: 0;
}
.status-strip span::before {
    content: "";
    width: 7px;
    height: 7px;
    border-radius: 999px;
    background: var(--green);
    box-shadow: none;
    animation: none;
}

.status-strip strong {
    color: var(--c-acc-xl);
    font-weight: 700;
}
.status-strip em {
    color: var(--accent);
    font-style: normal;
    font-weight: 700;
    font-variant-numeric: tabular-nums;
    animation: none;
}
@keyframes uptimeFlicker {
    0%,91%,93%,95%,100% { opacity: 1; }
    92% { opacity: .50; }
    94% { opacity: .80; }
}

@keyframes statusShimmer {
    0% { transform: translateX(-115%); opacity: .0; }
    18% { opacity: .55; }
    45% { transform: translateX(115%); opacity: .15; }
    100% { transform: translateX(115%); opacity: 0; }
}

@keyframes onlinePulse {
    /* Só opacity (GPU). O filter:drop-shadow animado foi removido — repintava texto a cada frame. */
    0%, 100% { opacity: .92; }
    50% { opacity: .60; }
}

/* Navegacao principal — mais dinamica, guiada e alinhada ao pulso do topo */
[data-testid="stSidebar"] .stRadio > label {
    display: flex !important;
    align-items: center !important;
    gap: 7px !important;
    margin: 0 0 10px !important;
    color: var(--t-n8) !important;
    font-family: var(--mono) !important;
    font-size: .63rem !important;
    font-weight: 800 !important;
    letter-spacing: .12em !important;
    text-transform: uppercase !important;
}
[data-testid="stSidebar"] .stRadio > label::before {
    content: "";
    width: 6px;
    height: 6px;
    border-radius: 999px;
    background: var(--accent);
    box-shadow: none;
    animation: none;
}
[data-testid="stSidebar"] div[role="radiogroup"] {
    position: relative !important;
}
[data-testid="stSidebar"] div[role="radiogroup"]::before {
    content: "";
    position: absolute;
    left: 1px;
    top: 10px;
    bottom: 10px;
    width: 1px;
    pointer-events: none;
    background: linear-gradient(180deg, transparent, var(--ta280), rgba(var(--rgb-run-l),.18), transparent);
    box-shadow: none;
    opacity: .72;
}
[data-testid="stSidebar"] div[role="radiogroup"] > label {
    isolation: isolate !important;
    overflow: hidden !important;
    min-height: 46px !important;
    padding: 3px 14px 3px 15px !important;
    border-radius: 10px !important;
    transform-origin: left center !important;
}
[data-testid="stSidebar"] div[role="radiogroup"] > label [data-testid="stMarkdownContainer"] {
    position: relative !important;
    z-index: 2 !important;
    padding: 8px 30px 8px 8px !important;
}
[data-testid="stSidebar"] div[role="radiogroup"] > label::before {
    z-index: 3 !important;
}
[data-testid="stSidebar"] div[role="radiogroup"] > label::after {
    z-index: 4 !important;
}
[data-testid="stSidebar"] div[role="radiogroup"] > label > div:last-child::before {
    content: "";
    position: absolute;
    inset: 0;
    z-index: 1;
    pointer-events: none;
    background: linear-gradient(110deg, transparent 0%, var(--ta130) 42%, rgba(var(--rgb-run-l),.11) 50%, transparent 62%);
    opacity: 0;
    transform: translateX(-38%);
    transition: opacity .28s ease, transform .55s cubic-bezier(.16, 1, .3, 1);
}
[data-testid="stSidebar"] div[role="radiogroup"] > label:hover > div:last-child::before,
[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) > div:last-child::before {
    opacity: 1;
    transform: translateX(0);
}
[data-testid="stSidebar"] div[role="radiogroup"] > label p {
    position: relative !important;
    padding-right: 4px !important;
    font-size: .80rem !important;
    line-height: 1.12 !important;
}
[data-testid="stSidebar"] div[role="radiogroup"] > label p::after {
    content: "";
    position: absolute;
    left: 0;
    right: 100%;
    bottom: -7px;
    height: 1px;
    background: linear-gradient(90deg, var(--ta950), rgba(var(--rgb-run-l),.80), transparent);
    box-shadow: none;
    opacity: 0;
    transition: right .42s cubic-bezier(.16, 1, .3, 1), opacity .22s ease;
}
[data-testid="stSidebar"] div[role="radiogroup"] > label:hover p::after,
[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) p::after {
    right: 16%;
    opacity: .9;
}
[data-testid="stSidebar"] div[role="radiogroup"] > label:hover {
    transform: translateX(7px) scale(1.012) !important;
}
[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) {
    transform: translateX(2px) !important;
    animation: none;
}
[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked)::before {
    width: 4px !important;
    animation: none;
}
[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked)::after {
    animation: none;
}
@keyframes navGuidePulse {
    0%,100% { opacity: .45; transform: scale(.92); }
    50% { opacity: 1; transform: scale(1.18); }
}
@keyframes navRailGlow {
    0%,100% { box-shadow: none; }
    50% { box-shadow: none; }
}
@keyframes navDotOrbit {
    0%,100% { transform: translateY(-50%) scale(1.12); box-shadow: none; }
    50% { transform: translateY(-50%) scale(1.42); box-shadow: none; }
}
@keyframes navActiveBreath {
    0%,100% { box-shadow: 0 0 0 1px var(--ta160), 0 16px 38px var(--ta120); }
    50% { box-shadow: 0 0 0 1px var(--ta240), 0 18px 46px rgba(var(--rgb-run-l),.12); }
}

/* Subtle entrance for nav items */
[data-testid="stSidebar"] div[role="radiogroup"] > label {
    animation: navItemIn .42s cubic-bezier(.16, 1, .3, 1) both;
}
[data-testid="stSidebar"] div[role="radiogroup"] > label:nth-of-type(1) { animation-delay: .04s; }
[data-testid="stSidebar"] div[role="radiogroup"] > label:nth-of-type(2) { animation-delay: .09s; }
[data-testid="stSidebar"] div[role="radiogroup"] > label:nth-of-type(3) { animation-delay: .14s; }
[data-testid="stSidebar"] div[role="radiogroup"] > label:nth-of-type(4) { animation-delay: .19s; }
[data-testid="stSidebar"] div[role="radiogroup"] > label:nth-of-type(5) { animation-delay: .24s; }

@keyframes navItemIn {
    /* Sem filter:blur — animar filter força repaint na main thread a cada frame.
       translate/opacity são compositados na GPU (sem engasgo no rerun). */
    from { opacity: 0; transform: translateX(-10px) translateY(5px) scale(.985); }
    to { opacity: 1; transform: translateX(0) translateY(0) scale(1); }
}

@keyframes brandFloat {
    0%, 100% { transform: translateY(0)    scale(1); }
    50%       { transform: translateY(-3px) scale(1.009); }
}

@keyframes titleSheen {
    /* Glitch via opacity/transform (GPU). Removidos os filter:drop-shadow que repintavam. */
    0%   { background-position: 0% 50%;   opacity: 1; transform: none; }
    76%  { background-position: 100% 50%; opacity: 1; transform: none; }
    78%  { opacity: .70; transform: skewX(-.8deg); }
    79%  { opacity: .88; transform: skewX(.4deg); }
    80%  { opacity: 1; transform: none; }
    100% { background-position: 180% 50%; opacity: 1; }
}

@keyframes headerAtmosphere {
    0% { transform: translate3d(-2%, -1%, 0) scale(1); opacity: .75; }
    45% { transform: translate3d(1%, 1.5%, 0) scale(1.02); opacity: .92; }
    100% { transform: translate3d(-2%, -1%, 0) scale(1); opacity: .75; }
}
.nav-context {
    margin-top: 12px;
    padding: 10px 12px;
    border: 1px solid var(--ta140);
    border-radius: 10px;
    background:
        linear-gradient(135deg, var(--ta100), rgba(var(--rgb-run-l),.055)),
        rgba(var(--rgb-n1),.82);
    box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.035);
    /* Camada GPU própria — isola o repaint do pulse de box-shadow/borda. */
    transform: translateZ(0);
    will-change: transform;
    animation: none;
}
@keyframes navContextPulse {
    0%,100% { border-color: var(--ta140); box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.035); }
    50%      { border-color: var(--ta300); box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.035); }
}
.nav-context-label {
    color: var(--accent);
    font-size: .58rem;
    font-family: var(--mono);
    font-weight: 700;
    letter-spacing: .12em;
    text-transform: uppercase;
}
.nav-context-title {
    margin-top: 4px;
    color: var(--txt-1);
    font-size: .78rem;
    font-weight: 750;
    line-height: 1.18;
}
.nav-context-copy {
    margin-top: 3px;
    color: var(--t-n8);
    font-size: .66rem;
    line-height: 1.28;
    display: -webkit-box;
    -webkit-line-clamp: 2;
    -webkit-box-orient: vertical;
    overflow: hidden;
}
.nav-context {
    position: relative;
    overflow: hidden;
    padding: 10px 12px 11px;
    border-color: var(--ta200);
    box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.04), 0 12px 28px var(--shc-3);
}
.nav-context::before {
    content: "";
    position: absolute;
    inset: 0;
    pointer-events: none;
    background: linear-gradient(105deg, transparent, var(--ta100), rgba(var(--rgb-run-l),.08), transparent);
    transform: translateX(-64%);
    animation: none;
}
.nav-context::after {
    content: "";
    position: absolute;
    left: 14px;
    right: 14px;
    top: 0;
    height: 1px;
    background: linear-gradient(90deg, transparent, var(--ta820), rgba(var(--rgb-run-l),.52), transparent);
    box-shadow: none;
}
.nav-context-label,
.nav-context-title,
.nav-context-copy,
.nav-context-signal {
    position: relative;
    z-index: 1;
}
.nav-context-signal {
    display: flex;
    gap: 4px;
    margin-bottom: 6px;
}
.nav-context-signal span {
    width: 4px;
    height: 4px;
    border-radius: 999px;
    background: var(--ta880);
    box-shadow: none;
    animation: none;
}
.nav-context-signal span:nth-child(2) {
    background: rgba(var(--rgb-acc-l),.78);
    animation-delay: .18s;
}
.nav-context-signal span:nth-child(3) {
    background: rgba(var(--rgb-run-l),.82);
    animation-delay: .36s;
}
@keyframes navContextScan {
    0%,100% { transform: translateX(-64%); opacity: .10; }
    45% { transform: translateX(48%); opacity: .78; }
}
@keyframes signalStep {
    0%,100% { opacity: .32; transform: translateY(0) scale(.82); }
    50% { opacity: 1; transform: translateY(-1px) scale(1.18); }
}
.sidebar-footer {
    position: relative;
    width: 100%;
    margin: 18px 0 8px;
    padding: 12px 13px;
    font-size: .68rem;
    color: var(--t-n8);
    border: 1px solid var(--ta120);
    border-radius: 10px;
    background: rgba(var(--rgb-n1),.92);
    box-shadow: 0 12px 30px var(--shc-3), inset 0 1px 0 rgba(var(--rgb-lift),.03);
    animation: none;
}
@keyframes footerPulse {
    0%,100% { border-color: var(--ta120); box-shadow: 0 12px 30px var(--shc-3), inset 0 1px 0 rgba(var(--rgb-lift),.03); }
    50%      { border-color: var(--ta260); box-shadow: 0 12px 30px var(--shc-3), inset 0 1px 0 rgba(var(--rgb-lift),.03); }
}
.sidebar-footer a { color: var(--accent) !important; font-weight: 700; }
.sidebar-footer small {
    display: block;
    margin-top: 3px;
    color: var(--t-n6);
    font-size: .64rem;
}

/* Page hero */
.page-hero {
    position: relative;
    overflow: hidden;
    padding: 22px 26px 22px 28px;
    margin: 0 0 24px;
    border: 1px solid var(--ta090);
    border-radius: 10px;
    background: var(--n1);
    box-shadow: none;
}
.page-hero::before {
    content: "";
    position: absolute;
    left: 0; top: 0; bottom: 0;
    width: 2px;
    pointer-events: none;
    background: var(--f-acc);
    box-shadow: none;
}
.page-hero::after {
    position: absolute;
    right: 22px;
    top: 18px;
    color: rgba(var(--rgb-t-n8),.16);
    font-family: var(--mono);
    font-size: .72rem;
    font-weight: 700;
    letter-spacing: .18em;
}
.page-hero-eyebrow {
    position: relative;
    z-index: 1;
    display: inline-flex;
    align-items: center;
    gap: 8px;
    font-size: .58rem; font-weight: 600;
    font-family: var(--mono);
    letter-spacing: .14em; text-transform: uppercase;
    color: rgba(var(--rgb-c-acc-l),.75); margin-bottom: 9px;
}
.page-hero-eyebrow::before {
    content: "";
    width: 7px; height: 7px;
    border-radius: 999px;
    background: var(--green);
    box-shadow: none;
}
.page-hero-title {
    position: relative;
    z-index: 1;
    font-size: clamp(1.15rem, 1.6vw, 1.5rem); font-weight: 600;
    color: var(--t-n11); line-height: 1.2;
    margin: 0 0 5px;
    letter-spacing: -.005em;
}
.page-hero-desc {
    position: relative;
    z-index: 1;
    max-width: 680px;
    font-size: .9rem;
    line-height: 1.6;
    color: rgba(var(--rgb-t-n9),.82);
    margin: 0;
}

/* Section label */
.sec-label {
    display: flex; align-items: center; gap: 9px;
    font-size: .72rem; font-weight: 600;
    font-family: var(--mono);
    text-transform: uppercase; letter-spacing: .05em;
    color: var(--t-n9); padding-bottom: 12px;
    border-bottom: 1px solid var(--ta070);
    margin: 6px 0 18px;
}
.sec-dot {
    width: 6px; height: 6px; border-radius: 50%;
    background: var(--accent);
    box-shadow: none;
    flex-shrink: 0;
}

/* KPI cards */
.kpi-wrap { display: grid; grid-template-columns: repeat(4,1fr); gap: 14px; margin-bottom: 24px; }
.kpi-card {
    background:
        linear-gradient(145deg, rgba(var(--rgb-sup-a),.92), rgba(var(--rgb-sup-b),.80)) !important;
    border: 1px solid rgba(var(--rgb-n9),.14) !important;
    border-radius: var(--radius);
    padding: 18px 20px;
    position: relative; overflow: hidden;
    box-shadow: var(--shadow-soft);
    transition: border-color .2s, transform .2s, box-shadow .2s;
}
.kpi-card:hover {
    border-color: var(--border-hi) !important;
    transform: translateY(-2px);
    box-shadow: 0 18px 44px var(--shc-2), 0 0 0 1px var(--ta120);
}
.kpi-card::before {
    content: ''; position: absolute;
    top: 0; left: 0; bottom: 0; width: 3px;
    background: var(--c, var(--accent));
    box-shadow: none;
}
.kpi-card::after {
    content: "";
    position: absolute;
    inset: 0;
    background:
        linear-gradient(
            135deg,
            rgba(var(--rgb-lift),0.03) 0%,
            rgba(var(--rgb-lift),0.01) 30%,
            transparent 60%
        );
    pointer-events: none;
}
.kpi-card .kpi-ico { font-size: 1.22rem; margin-bottom: 10px; display: block; }
.kpi-card .kpi-val {
    font-family: var(--mono); font-size: 2.2rem;
    font-weight: 700; line-height: 1;
    color: var(--c, var(--txt-1));
    margin-bottom: 4px;
}
.kpi-card .kpi-lbl {
    font-size: .68rem; text-transform: uppercase;
    letter-spacing: .1em; color: var(--t-n9); font-weight: 700;
}

/* Status pills */
.pill {
    display: inline-flex; align-items: center; gap: 5px;
    padding: 4px 12px; border-radius: 999px;
    font-size: var(--text-xs); font-weight: 500;
    font-family: var(--font);
    letter-spacing: .01em; white-space: nowrap;
}
.pill-gray  { background: rgba(var(--rgb-n9),.14); color: var(--t-n10); }
.pill-amber { background: rgba(var(--rgb-run),.14);  color: var(--c-run); }
.pill-green { background: rgba(var(--rgb-ok),.14);  color: var(--c-ok); }
.pill-red   { background: rgba(var(--rgb-crit-l),.14); color: var(--c-crit-l);
    animation: none; }
@keyframes pillBreath {
    0%, 100% { box-shadow: 0 0 0 0 rgba(var(--rgb-crit-l),0); }
    50%      { box-shadow: 0 0 0 3px rgba(var(--rgb-crit-l),.10); }
}
.pill-blue  { background: var(--ta140);  color: var(--c-acc-l); }

/* Ticket confirm card */
.confirm-card {
    position: relative;
    overflow: hidden;
    background:
        linear-gradient(135deg, var(--ta100), rgba(var(--rgb-run-l),.045)),
        rgba(var(--rgb-n1),.82);
    border: 1px solid var(--ta220);
    border-radius: var(--radius); padding: 20px 24px;
    margin-top: 16px;
    box-shadow: var(--shadow-soft);
}
.confirm-card::before {
    content: "";
    position: absolute;
    left: 0; top: 0; bottom: 0;
    width: 3px;
    background: var(--green);
    box-shadow: none;
}
.confirm-card h4 { color: var(--accent) !important; font-size: .85rem !important;
    letter-spacing: .08em !important; text-transform: uppercase !important; margin: 0 0 8px !important; }
.confirm-card p { font-size: .82rem; color: var(--txt-2); margin: 3px 0; }

/* UX helpers */
.process-rail {
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 12px;
    margin: -8px 0 18px;
}
.process-step {
    position: relative;
    min-height: 86px;
    padding: 14px 14px 14px 16px;
    border: 1px solid var(--ta140);
    border-radius: 10px;
    background: linear-gradient(145deg, rgba(var(--rgb-sup-a),.80), rgba(var(--rgb-sup-b),.68));
    box-shadow: var(--shadow-soft);
}
.process-step::before {
    content: attr(data-step);
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 24px;
    height: 24px;
    margin-bottom: 9px;
    border-radius: 7px;
    background: var(--ta120);
    border: 1px solid var(--ta220);
    color: var(--accent);
    font-family: var(--mono);
    font-size: .68rem;
    font-weight: 800;
}
.process-step strong {
    display: block;
    color: var(--txt-1);
    font-size: .86rem;
    margin-bottom: 3px;
}
.process-step span {
    display: block;
    color: var(--t-n9) !important;
    font-size: .76rem;
    line-height: 1.35;
}

/* ── MINIMAL MODE (GLOBAL) ── */
/* Labels (ex.: "NOME COMPLETO", "SETOR") less loud */
[data-testid="stWidgetLabel"] p,
[data-testid="stWidgetLabel"] label,
.stTextInput label,
.stTextArea label,
.stSelectbox label,
.stSlider label,
.stSelectSlider label {
    font-size: .82rem !important;
    letter-spacing: .01em !important;
    text-transform: none !important;
    color: var(--c-acc-xl) !important;
    font-weight: 700 !important;
}

/* Helper/caption text less prominent */
[data-testid="stCaptionContainer"] p,
[data-testid="stCaptionContainer"],
.stCaption,
.st-emotion-cache-1v0mbdj, 
.st-emotion-cache-1wmy9hl {
    color: rgba(var(--rgb-t-n9),.73) !important;
    font-size: .72rem !important;
}

/* Placeholders more subtle */
input::placeholder,
textarea::placeholder {
    color: rgba(var(--rgb-t-n9),.65) !important;
    opacity: .75 !important;
}

/* File uploader helper line */
[data-testid="stFileUploader"] small,
[data-testid="stFileUploader"] [data-testid="stMarkdownContainer"] p {
    color: rgba(var(--rgb-t-n9),.72) !important;
    font-size: .72rem !important;
}

/* Dataframes: more dashboard, less Excel */
div[data-testid="stDataFrame"] {
    border-radius: 12px !important;
    overflow: hidden !important;
    border: 1px solid var(--ta140) !important;
    background: rgba(var(--rgb-n1),.55) !important;
    box-shadow: 0 18px 44px var(--shc-3) !important;
}
div[data-testid="stDataFrame"] [data-baseweb="data-table"] {
    border-radius: 12px !important;
}
div[data-testid="stDataFrame"] [role="columnheader"] {
    background: linear-gradient(180deg, rgba(var(--rgb-sup-a),.92), rgba(var(--rgb-sup-b),.70)) !important;
    border-bottom: 1px solid var(--ta140) !important;
}
div[data-testid="stDataFrame"] [role="columnheader"] * {
    font-family: var(--mono) !important;
    font-size: .62rem !important;
    letter-spacing: .14em !important;
    text-transform: uppercase !important;
    color: rgba(var(--rgb-t-n9),.87) !important;
}
div[data-testid="stDataFrame"] [role="row"] {
    border-bottom: 1px solid rgba(var(--rgb-n9),.08) !important;
}
div[data-testid="stDataFrame"] [role="gridcell"] {
    font-size: .78rem !important;
    color: rgba(var(--rgb-t-n11),.92) !important;
}
div[data-testid="stDataFrame"] [role="row"]:nth-child(even) [role="gridcell"] {
    background: rgba(var(--rgb-lift),.012) !important;
}
div[data-testid="stDataFrame"] [role="row"]:hover [role="gridcell"] {
    background: var(--ta050) !important;
}

/* Tabs: more product, less default */
div[data-testid="stTabs"] [role="tablist"] {
    gap: 8px !important;
    padding: 6px 6px 8px !important;
    border-radius: 12px !important;
    background: rgba(var(--rgb-n1),.55) !important;
    border: 1px solid var(--ta120) !important;
}
div[data-testid="stTabs"] button[role="tab"] {
    border-radius: 10px !important;
    padding: 10px 12px !important;
    background: rgba(var(--rgb-n2),.55) !important;
    border: 1px solid rgba(var(--rgb-n9),.12) !important;
    transition: transform .18s ease, border-color .18s ease, background .18s ease, box-shadow .18s ease !important;
}
div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
    border-color: var(--ta350) !important;
    background: linear-gradient(135deg, var(--ta100), rgba(var(--rgb-run-l),.055)) !important;
    box-shadow: none;
}
div[data-testid="stTabs"] button[role="tab"]:hover {
    transform: translateY(-1px) !important;
    border-color: var(--ta220) !important;
}
.ux-note {
    display: flex;
    align-items: flex-start;
    gap: 10px;
    padding: 10px 11px;
    border: 1px solid var(--ta160);
    border-radius: 10px;
    background:
        linear-gradient(135deg, var(--ta080), rgba(var(--rgb-run-l),.035)),
        rgba(var(--rgb-n1),.72);
    color: var(--t-n9);
    font-size: var(--text-sm);
    line-height: 1.45;
}
.ux-note::before {
    content: "";
    flex: 0 0 auto;
    width: 8px;
    height: 8px;
    margin-top: 5px;
    border-radius: 999px;
    background: var(--accent);
    box-shadow: none;
}
.form-action-note {
    min-height: 40px;
    display: flex;
    align-items: center;
    padding: 8px 10px;
    border: 1px solid rgba(var(--rgb-n9),.12);
    border-radius: 9px;
    background: rgba(var(--rgb-n1),.62);
    color: var(--t-n9);
    font-size: var(--text-sm);
    line-height: 1.35;
}
.empty-state {
    padding: 28px 22px;
    text-align: center;
    border: 1px dashed var(--ta160);
    border-radius: 18px;
    background:
        linear-gradient(135deg, var(--ta070), rgba(var(--rgb-run-l),.035)),
        rgba(var(--rgb-n1),.66);
}
.empty-state strong {
    display: block;
    color: var(--txt-1);
    font-size: var(--text-md);
    margin-bottom: 5px;
}
.empty-state span {
    color: var(--t-n9) !important;
    font-size: var(--text-sm);
}
.search-panel {
    margin-bottom: 18px;
    padding: 14px 16px;
    border: 1px solid var(--ta140);
    border-radius: 10px;
    background: rgba(var(--rgb-n1),.70);
}
.queue-card {
    position: relative;
    overflow: hidden;
    border: 1px solid rgba(var(--rgb-n9),.13);
    border-radius: 10px;
    padding: 10px 12px;
    background: linear-gradient(145deg, rgba(var(--rgb-sup-a),.82), rgba(var(--rgb-sup-b),.72));
    box-shadow: 0 14px 32px var(--shc-2);
}
.queue-card::before {
    content: "";
    position: absolute;
    left: 0; top: 0; bottom: 0;
    width: 3px;
    background: var(--queue-accent, var(--accent));
    box-shadow: none;
}
.queue-card + .queue-card { margin-top: 10px; }
.queue-meta {
    color: var(--t-n9);
    font-size: .74rem;
    font-family: var(--mono);
    letter-spacing: .04em;
    text-transform: uppercase;
}
.queue-critical { --queue-accent: var(--c-crit-l); box-shadow: -3px 0 20px rgba(var(--rgb-crit-l),.22); }
.queue-high { --queue-accent: var(--c-run); }
.queue-medium { --queue-accent: var(--c-acc); }
.queue-low { --queue-accent: var(--c-ok); }
.work-card-head {
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    gap: 12px;
    margin-bottom: 8px;
}
.work-card-id {
    color: var(--accent);
    font-family: var(--mono);
    font-size: .68rem;
    font-weight: 800;
    letter-spacing: .10em;
    text-transform: uppercase;
}
.work-card-title {
    color: var(--txt-1);
    font-size: .98rem;
    font-weight: 800;
    line-height: 1.25;
    margin-top: 2px;
}
.work-card-subtitle {
    color: var(--t-n9);
    font-size: .78rem;
    margin-top: 2px;
}
.work-card-desc {
    color: var(--t-n10);
    font-size: .84rem;
    line-height: 1.45;
    margin: 9px 0 12px;
}
/* linha de metadados compacta dentro do card (substitui a fileira de chips) */
.work-card-meta {
    display: flex; flex-wrap: wrap; gap: 6px 16px; align-items: center;
    margin-top: 10px; padding-top: 10px;
    border-top: 1px solid rgba(var(--rgb-n9),.09);
}
.work-card-meta span {
    font-size: .74rem; color: var(--t-n8); display: inline-flex; gap: 5px; align-items: center;
}
.work-card-meta strong { color: var(--t-n10); font-weight: 700; }
/* linha enxuta de arquivo (histórico de concluídos) */
.arc-row { display: flex; align-items: center; gap: 8px 14px; flex-wrap: wrap; }
.arc-row .arc-id {
    font-family: 'IBM Plex Mono', ui-monospace, monospace; font-size: .60rem; font-weight: 700;
    letter-spacing: .10em; text-transform: uppercase; color: var(--c-acc);
}
.arc-row .arc-user { font-weight: 700; color: var(--t-n11); font-size: .85rem; }
.arc-row .arc-meta { color: var(--t-n8); font-size: .76rem; }
.arc-row .arc-date { color: var(--t-n7); font-size: .72rem; margin-left: auto; white-space: nowrap; }
.work-badge-row {
    display: flex;
    flex-wrap: wrap;
    justify-content: flex-end;
    gap: 6px;
}
.meta-chip {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    min-height: 36px;
    width: 100%;
    padding: 8px 10px;
    border: 1px solid rgba(var(--rgb-n9),.12);
    border-radius: 9px;
    background: rgba(var(--rgb-n1),.64);
    color: var(--t-n9);
    font-size: .78rem;
    line-height: 1.25;
}
.meta-chip strong {
    color: var(--txt-1);
    font-weight: 750;
}
.history-panel-head {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    margin: 4px 0 16px;
    padding: 2px 0 10px;
    border-bottom: 1px solid rgba(var(--rgb-n9),.14);
}
.history-panel-title {
    display: flex;
    align-items: center;
    gap: 10px;
    color: var(--t-n11);
    font-size: 1.16rem;
    font-weight: 850;
    letter-spacing: 0;
}
.history-title-dot {
    width: 13px;
    height: 13px;
    border-radius: 999px;
    background: var(--f-ok);
    box-shadow: none;
}
.history-count {
    display: inline-flex;
    align-items: center;
    min-height: 30px;
    padding: 5px 10px;
    border: 1px solid var(--ta180);
    border-radius: 999px;
    background: var(--ta080);
    color: var(--c-acc-l);
    font-family: var(--mono);
    font-size: .72rem;
    font-weight: 800;
    text-transform: uppercase;
}
.history-record {
    display: grid;
    grid-template-columns: minmax(132px, .9fr) minmax(260px, 2.15fr) minmax(150px, .95fr);
    gap: 16px;
    align-items: center;
    min-height: 74px;
}
.history-protocol {
    color: var(--t-n11);
    font-family: var(--mono);
    font-size: .95rem;
    font-weight: 900;
    letter-spacing: .03em;
}
.history-priority {
    display: inline-flex;
    align-items: center;
    margin-top: 10px;
    min-height: 27px;
    padding: 5px 9px;
    border-radius: 999px;
    font-size: .74rem;
    font-weight: 850;
    border: 1px solid rgba(var(--rgb-n9),.16);
    white-space: nowrap;
}
.history-prio-critical {
    color: var(--c-crit-l);
    background: rgba(var(--rgb-crit-l),.14);
    border-color: rgba(var(--rgb-crit-l),.34);
}
.history-prio-medium {
    color: var(--c-run-l);
    background: rgba(var(--rgb-run),.13);
    border-color: rgba(var(--rgb-run),.30);
}
.history-prio-low {
    color: var(--c-ok-l);
    background: rgba(var(--rgb-ok),.12);
    border-color: rgba(var(--rgb-ok),.28);
}
.history-user {
    color: var(--t-n11);
    font-size: 1rem;
    font-weight: 850;
    line-height: 1.2;
}
.history-meta {
    margin-top: 5px;
    color: var(--t-n10);
    font-size: .80rem;
    font-weight: 650;
}
.history-desc {
    margin-top: 7px;
    color: var(--c-acc-xl);
    font-size: .84rem;
    line-height: 1.35;
}
.history-status {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    min-height: 30px;
    padding: 6px 10px;
    border-radius: 999px;
    background: rgba(var(--rgb-ok),.12);
    border: 1px solid rgba(var(--rgb-ok),.28);
    color: var(--c-ok-l);
    font-weight: 850;
    font-size: .82rem;
}
.history-date {
    margin-top: 8px;
    color: var(--t-n10);
    font-family: var(--mono);
    font-size: .78rem;
    font-weight: 700;
}
@media (max-width: 900px) {
    .history-panel-head,
    .history-record {
        display: flex;
        flex-direction: column;
        align-items: flex-start;
    }
    .history-record {
        gap: 10px;
        min-height: auto;
    }
}
.ops-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
    gap: 14px;
    margin: 14px 0 20px;
}
.ops-card {
    position: relative;
    overflow: hidden;
    min-height: 112px;
    padding: 16px 16px 14px;
    border: 1px solid var(--ta120);
    border-radius: 10px;
    background: linear-gradient(145deg, rgba(var(--rgb-sup-a),.94), rgba(var(--rgb-sup-b),.84));
    box-shadow: var(--shadow-soft), inset 0 1px 0 rgba(var(--rgb-lift),.035);
}
.ops-card::before {
    content: "";
    position: absolute;
    inset: 0 0 auto;
    height: 2px;
    background: var(--tone, var(--f-acc));
    box-shadow: none;
}
.ops-label {
    color: var(--t-n9);
    font-family: var(--mono);
    font-size: .66rem;
    font-weight: 800;
    letter-spacing: .1em;
    text-transform: uppercase;
}
.ops-value {
    margin-top: 8px;
    color: var(--t-n11);
    font-family: var(--mono);
    font-size: 1.65rem;
    font-weight: 800;
    line-height: 1;
}
.ops-copy {
    margin-top: 8px;
    color: var(--t-n8);
    font-size: .78rem;
    line-height: 1.35;
}
.insight-panel {
    border: 1px solid var(--ta160);
    border-radius: 10px;
    padding: 15px 16px;
    background:
        linear-gradient(135deg, var(--ta080), rgba(var(--rgb-run-l),.04)),
        rgba(var(--rgb-n1),.72);
    color: var(--t-n10);
    font-size: .86rem;
    line-height: 1.55;
}
.mini-kanban {
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 12px;
    margin: 8px 0 18px;
}
.kanban-lane {
    min-height: 92px;
    padding: 13px;
    border-radius: 10px;
    border: 1px solid rgba(var(--rgb-n9),.12);
    background: rgba(var(--rgb-n1),.62);
}
.kanban-lane strong {
    display: block;
    color: var(--t-n11);
    font-size: .88rem;
    margin-bottom: 7px;
}
.kanban-lane span {
    color: var(--t-n8) !important;
    font-size: .78rem;
}

.animate-fade-in {
    animation: fadeInUp .35s ease both;
}
@keyframes fadeInUp {
    from { opacity: 0; transform: translateY(6px); }
    to { opacity: 1; transform: translateY(0); }
}
/* ── UI SKELETONS (enterprise loading feel) ── */
.skeleton-wrap {
    display: grid;
    gap: 14px;
    padding: 10px 0 6px;
}
.skeleton-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 14px;
}
.skeleton-card {
    border-radius: var(--radius);
    border: 1px solid rgba(var(--rgb-n9),.14);
    background: rgba(var(--rgb-n1),.62);
    overflow: hidden;
    box-shadow: var(--shadow-soft);
    position: relative;
    min-height: 86px;
}
.skeleton-card::after {
    content: "";
    position: absolute;
    inset: 0;
    transform: translateX(-120%);
    background: linear-gradient(
        90deg,
        var(--ta000) 0%,
        var(--ta100) 38%,
        rgba(var(--rgb-run-l),.10) 55%,
        var(--ta000) 100%
    );
    animation: skeletonShimmer 1.15s ease-in-out infinite;
    pointer-events: none;
}
.skeleton-line {
    height: 12px;
    border-radius: 999px;
    background: rgba(var(--rgb-n9),.12);
    margin: 16px 14px;
}
.skeleton-line.sm { height: 9px; width: 72%; }
.skeleton-line.md { height: 10px; width: 86%; }
.skeleton-line.lg { height: 12px; width: 100%; }
.skeleton-label {
    color: rgba(var(--rgb-t-n9),.82);
    font-size: .68rem;
    font-family: var(--mono);
    letter-spacing: .12em;
    text-transform: uppercase;
    padding: 2px 4px 0;
}
@keyframes skeletonShimmer {
    0% { transform: translateX(-120%); opacity: .25; }
    40% { opacity: .55; }
    100% { transform: translateX(120%); opacity: .20; }
}

@media (prefers-reduced-motion: reduce) {
    .skeleton-card::after { animation: none !important; }
}

.skeleton-rail {
    display: grid;
    gap: 12px;
    margin-top: 2px;
}
.skeleton-queue {
    border-radius: 10px;
    border: 1px solid rgba(var(--rgb-n9),.13);
    background: linear-gradient(145deg, rgba(var(--rgb-sup-a),.74), rgba(var(--rgb-sup-b),.64));
    box-shadow: 0 14px 32px var(--shc-3);
    overflow: hidden;
    position: relative;
    padding: 10px 12px;
    min-height: 98px;
}
.skeleton-queue::before {
    content: "";
    position: absolute;
    inset: 0;
    background: transparent;
    pointer-events: none;
}
.skeleton-queue .skeleton-block {
    height: 14px;
    border-radius: 999px;
    background: rgba(var(--rgb-n9),.12);
    margin: 12px 14px;
}
.skeleton-queue .skeleton-block:nth-child(2) { width: 62%; }
.skeleton-queue .skeleton-block:nth-child(3) { width: 78%; }
.skeleton-queue .skeleton-block:nth-child(4) { width: 92%; }

@media (max-width: 900px) {
    .main .block-container { padding: 1.3rem 1rem 2rem; }
    .page-hero { padding: 10px 12px 12px; }
    .page-hero::after { display: none; }
    .process-rail { grid-template-columns: 1fr; }
    .kpi-wrap { grid-template-columns: repeat(2, minmax(0,1fr)); }
    .mini-kanban { grid-template-columns: 1fr; }
}
@media (max-width: 560px) {
    .kpi-wrap { grid-template-columns: 1fr; }
}
/* ═══════════════════════════════════════════════════════
       ── RESPONSIVIDADE PARA TELAS MENORES (MOBILE/TABLET) ──
       ═══════════════════════════════════════════════════════ */
    @media (max-width: 768px) {
        /* Sidebar becomes lighter + more touch friendly on mobile */
        [data-testid="stSidebar"] {
            box-shadow: inset -1px 0 0 rgba(var(--rgb-lift),.03), 12px 0 44px var(--shc-3) !important;
        }
        .sidebar-footer {
            position: relative !important;
            left: 0 !important;
            bottom: 0 !important;
            width: 100% !important;
            margin-top: 32px !important;
            box-shadow: none !important;
        }
        [data-testid="stSidebar"] > div {
            padding: 0 12px 32px !important;
        }
        .sidebar-header {
            padding: 16px 12px 12px !important;
            margin: 0 -12px 12px !important;
        }
        .logo-ring {
            width: 48px !important;
            height: 48px !important;
            font-size: 1.2rem !important;
            box-shadow: none;
        }
        .sidebar-title {
            font-size: 0.85rem !important;
        }
        .sidebar-sub {
            font-size: 0.62rem !important;
        }
        [data-testid="stSidebar"] div[role="radiogroup"] > label {
            min-height: 42px !important;
        }
        [data-testid="stSidebar"] div[role="radiogroup"] > label [data-testid="stMarkdownContainer"] {
            padding: 10px 28px 10px 12px !important;
        }
        [data-testid="stSidebar"] div[role="radiogroup"] > label p {
            font-size: 0.8rem !important;
        }
        .status-strip, .nav-context {
            margin-top: 10px !important;
            padding: 10px !important;
        }
    }

/* ══════════════════════════════════════════════════════════════════
   IDENTIDADE TLM — remove a cenografia sci-fi do tema anterior.
   Fundo chapado, sem grade, sem gradiente radial, sem brilho.
   ══════════════════════════════════════════════════════════════════ */
[data-testid="stAppViewContainer"] {
    background: var(--n0) !important;
    color: var(--t-n11) !important;
}
.stApp {
    background: var(--n0) !important;   /* tira a grade de 44px */
}
.stApp::before { display: none !important; }               /* véu superior */
[data-testid="stSidebar"] {
    background: var(--n1) !important;                      /* era radial ciano */
    border-right: 1px solid var(--n3) !important;
    box-shadow: none !important;
}
[data-testid="stSidebar"]::before { display: none !important; }  /* grade 28px */

/* Scrollbars: cinza discreto no lugar do gradiente ciano com brilho */
*, html, body { scrollbar-color: var(--n5) transparent !important; }
::-webkit-scrollbar { width: 8px !important; height: 8px !important; }
::-webkit-scrollbar-track { background: transparent !important; }
::-webkit-scrollbar-thumb {
    background: var(--n5) !important;
    border-radius: 99px !important;
    box-shadow: none !important;
}
::-webkit-scrollbar-thumb:hover { background: var(--n6) !important; }

/* Divisor: linha sólida em vez de faixa com brilho */
hr {
    background: var(--n3) !important;
    box-shadow: none !important;
}

/* Botão primário do Streamlit segue a marca: preto no claro, amarelo no escuro */
/* O texto de qualquer botão herda a cor do próprio botão — sem isto a regra
   global de `p`/`span` pinta o rótulo por cima e ele some no fundo escuro. */
button p, button span, button div, button [data-testid="stMarkdownContainer"] *,
[data-testid="stButton"] button [data-testid="stMarkdownContainer"] p,
[data-testid="stButton"] button [data-testid="stMarkdownContainer"] span,
[data-testid="stFormSubmitButton"] button [data-testid="stMarkdownContainer"] p,
[data-testid="stFormSubmitButton"] button [data-testid="stMarkdownContainer"] span {
    color: inherit !important;
}
.stButton > button[kind="primary"], .stFormSubmitButton > button { color: var(--act-ink) !important; }
.stButton > button[kind="primary"],
.stFormSubmitButton > button {
    background: var(--act) !important;
    border-color: var(--act) !important;
    color: var(--act-ink) !important;
    box-shadow: none !important;
}
.stButton > button[kind="primary"]:hover,
.stFormSubmitButton > button:hover {
    background: var(--act-hover) !important;
    border-color: var(--act-hover) !important;
}

/* Foco visível de teclado na cor da marca */
:focus-visible { outline: 2px solid var(--brand) !important; outline-offset: 2px !important; }

/* Widgets do Streamlit: as cores vinham fixas do config.toml. Aqui elas
   passam a seguir os tokens, para o tema claro tambem funcionar. */
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input,
[data-testid="stDateInput"] input,
[data-testid="stTextArea"] textarea,
[data-baseweb="select"] > div,
[data-baseweb="input"],
[data-baseweb="textarea"] {
    background: var(--n2) !important;
    border-color: var(--n4) !important;
    color: var(--t-n11) !important;
}
[data-testid="stTextInput"] input::placeholder,
[data-testid="stTextArea"] textarea::placeholder { color: var(--t-n7) !important; }
[data-baseweb="popover"] [role="listbox"],
[data-baseweb="menu"] {
    background: var(--n2) !important;
    border: 1px solid var(--n4) !important;
}
[data-baseweb="menu"] li:hover { background: var(--n3) !important; }
[data-testid="stExpander"] details,
[data-testid="stExpander"] summary {
    background: var(--n1) !important;
    border-color: var(--n3) !important;
}
[data-testid="stMetric"], [data-testid="stAlert"] {
    background: var(--n1) !important;
    border: 1px solid var(--n3) !important;
}
[data-testid="stDataFrame"] { border: 1px solid var(--n3) !important; }
[data-testid="stTabs"] button[role="tab"] { color: var(--t-n8) !important; }
[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
    color: var(--t-n11) !important;
    border-bottom-color: var(--f-acc) !important;
}

/* Envio de arquivo: o Streamlit escreve "Upload" e "200MB per file" em
   inglês e não expõe tradução. Trocamos o texto por CSS. O <span> das
   instruções não tem data-testid estável, então é identificado por
   exclusão: não é o ícone (tem data-testid) nem o wrapper do botão
   (contém um <button>). */
/* Zera o texto da zona de envio (some com "Drag and drop" / "200MB per file")
   e devolve tamanho apenas ao conteúdo do botão. Mais previsível que
   display:none, porque o rótulo do botão mora em spans sem data-testid. */
[data-testid="stFileUploaderDropzone"],
[data-testid="stFileUploaderDropzone"] *{ font-size:0 !important; }
[data-testid="stFileUploaderDropzone"] button,
[data-testid="stFileUploaderDropzone"] button *{ font-size:12.5px !important; }
[data-testid="stFileUploaderDropzone"] [data-testid="stIconMaterial"]{ font-size:18px !important; }
/* No Streamlit 1.56 os chips do arquivo enviado ficam DENTRO da zona, então o
   font-size:0 acima escondia o nome do anexo. Aqui ele volta. */
[data-testid="stFileUploaderDropzone"] [data-testid="stFileChips"],
[data-testid="stFileUploaderDropzone"] [data-testid="stFileChips"] *{
    font-size:11.5px !important;
}

[data-testid="stFileUploader"] [data-testid="stBaseButton-secondary"]
    [data-testid="stMarkdownContainer"] p{ font-size:0 !important; }
[data-testid="stFileUploader"] [data-testid="stBaseButton-secondary"]
    [data-testid="stMarkdownContainer"] p::after{
    content:"Escolher arquivo"; font-size:12.5px; font-weight:600;
}

/* Rótulo com ícone (ex.: Evidências) */
.lbl-anexo{display:flex;align-items:center;gap:8px;font-size:.88rem;color:var(--t-n10);margin:2px 0 6px}
.lbl-anexo i{font-style:normal;color:var(--t-n7)}
.ctx-tip{display:flex;align-items:flex-start}

/* Painéis em iframe (components.html) têm altura FIXA, e o Streamlit ainda
   trava o contêiner na mesma medida por uma classe gerada. Resultado: o
   dashboard era cortado quando os gráficos ficavam mais altos que o valor
   informado. Aqui o contêiner passa a acompanhar o iframe, que se mede e se
   ajusta sozinho por dentro. */
[data-testid="stElementContainer"]:has(> iframe[data-testid="stIFrame"]),
[data-testid="stElementContainer"]:has(> iframe) {
    height: auto !important;
    min-height: 0 !important;
}

/* ─────────────────────────────────────────────────────────────────
   Estilos dos assistentes (abrir chamado / solicitar compra).
   Antes eram injetados por st.markdown DENTRO do fragmento: a cada
   troca de urgência ou categoria o navegador recebia uma folha de
   estilo nova e refazia o cálculo da página inteira — era isso que
   parecia "refresh". Aqui entram uma vez por sessão.
   ───────────────────────────────────────────────────────────────── */
.ctx-panel{background:linear-gradient(160deg,rgba(var(--rgb-sup-a),.92),rgba(var(--rgb-sup-b),.85));
              border:1px solid var(--ta140);border-radius:14px;padding:20px 20px 18px;
              box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.05), 0 12px 28px var(--shc-2), 0 2px 6px var(--shc-2);position:relative;overflow:hidden}
            .ctx-panel::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;
              background:linear-gradient(90deg,var(--f-acc),var(--f-acc));box-shadow: none}
            .ctx-h{font-family:'IBM Plex Mono',monospace;font-size:.62rem;font-weight:800;letter-spacing:.16em;
              text-transform:uppercase;color:var(--c-acc);margin-bottom:16px}
            .ctx-step{display:flex;gap:12px;align-items:flex-start;margin-bottom:14px}
            .ctx-n{flex:0 0 24px;width:24px;height:24px;border-radius:50%;display:flex;align-items:center;
              justify-content:center;font-size:11px;font-weight:800;color:var(--c-acc);
              background:var(--ta100);border:1px solid var(--ta300)}
            .ctx-step strong{display:block;font-size:.84rem;color:var(--t-n11);font-weight:700;margin-bottom:2px}
            .ctx-step span{font-size:.74rem;color:var(--t-n8);line-height:1.45}
            .ctx-divider{height:1px;background:linear-gradient(90deg,transparent,var(--ta250),transparent);margin:4px 0 14px}
            .ctx-tip{font-size:.76rem;color:var(--c-acc-l);background:var(--ta060);
              border:1px solid var(--ta130);border-radius:9px;padding:10px 12px;line-height:1.5}
.cat3d:hover{
              transform:translateY(-3px)!important;
              border-color:var(--ta550)!important;
              box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.08), 0 14px 26px var(--shc-1);
            }

/* O indicador "Running" pisca a cada clique mesmo quando a resposta é
   instantânea, o que dá sensação de recarga. Ele só aparece se a operação
   passar de 1 segundo — aí a informação é útil de verdade. */
[data-testid="stStatusWidget"]{ animation: mostrarSeDemorar .01s linear 1s both; }
@keyframes mostrarSeDemorar{ from{opacity:0;visibility:hidden} to{opacity:1;visibility:visible} }

/* Cabeçalho embutido no cartão do formulário: sem borda, fundo ou sombra
   próprios — quem desenha a caixa agora é o container. */
.hz-embutido{background:none !important;border:0 !important;box-shadow:none !important;
  padding:0 0 14px !important;margin:0 0 4px !important;
  border-bottom:1px solid var(--n3) !important;border-radius:0 !important;}
.hz-embutido::before{display:none !important;}

/* ── BORDA DUPLA NOS CAMPOS ────────────────────────────────────────
   O Streamlit aninha <div data-baseweb="input"> > <div data-baseweb="base-input">.
   Várias regras do portal pintam borda, fundo e sombra nos DOIS, e o
   resultado é uma segunda linha desenhada por dentro da primeira. A moldura
   fica com o elemento de fora; o de dentro passa a ser só o invólucro do
   campo de texto. */
[data-baseweb="input"] [data-baseweb="base-input"],
[data-baseweb="textarea"] [data-baseweb="base-input"],
[data-baseweb="select"] [data-baseweb="base-input"] {
    border: 0 !important;
    background: transparent !important;
    box-shadow: none !important;
    border-radius: 0 !important;
    transform: none !important;
}

/* Faixa compacta quando não há chamados na fila */
.lf-vazio{display:flex;align-items:center;gap:10px;padding:9px 14px;margin-bottom:6px;
  border:1px solid var(--n3);border-radius:6px;background:var(--n1);}
.lf-vazio-txt{font-size:12.5px;color:var(--t-n7);}
.lf-vazio-tag{margin-left:auto;font-family:var(--mono);font-size:9.5px;letter-spacing:.12em;
  color:var(--t-n7);}

/* Marca da sidebar */
.tlm-brand { display:flex; align-items:center; gap:10px; padding:14px 4px 16px;
             border-bottom:1px solid var(--n3); margin-bottom:6px; }
.tlm-mark  { color:#fff !important; background:#0c0c0c;
             /* âmbar de preenchimento (#ffb700), não a variante de texto:
                --brand aponta para --c-acc, que no claro é marrom. Aqui o
                contorno fica sobre o quadrado preto do selo, então usa o
                âmbar da marca nos dois temas. */
             border:2px solid var(--f-acc); border-radius:4px;
             padding:5px 8px; font-weight:700; font-size:14px; line-height:1;
             letter-spacing:.02em; color:var(--t-n11); font-family:var(--font); }
.tlm-name  { font-size:13px; font-weight:600; color:var(--t-n11); line-height:1.25; }
.tlm-sub   { font-family:var(--mono); font-size:10px; letter-spacing:.09em;
             text-transform:uppercase; color:var(--t-n7); }

/* Alternador de tema */
[data-testid="stSidebar"] .stButton > button[kind="secondary"].tema-btn {
    width:100%; background:var(--n2) !important; border:1px solid var(--n3) !important;
    color:var(--t-n8) !important; font-size:11px !important;
}

/* ── CAMPOS DE FORMULÁRIO — UMA REGRA SÓ PARA O SITE INTEIRO ──────────
   Antes cada campo herdava duas bordas âmbar (uma no invólucro do BaseWeb,
   outra no contêiner do Streamlit) e um fundo tirado da PÁGINA. Dentro de um
   cartão branco isso vira um buraco bege com contorno amarelo. Aqui o campo
   ganha fundo próprio, UMA borda neutra, e o âmbar só no foco — que é quando
   ele significa alguma coisa.
   O prefixo `.stApp.stApp` é intencional: as regras antigas de campo usam
   seletores como `div[data-baseweb="select"] > div:first-child`, e sem repetir
   a classe a especificidade empata e a regra mais antiga vence. */
.stApp.stApp [data-testid="stTextInput"] > div,
.stApp.stApp [data-testid="stTextArea"] > div,
.stApp.stApp [data-testid="stNumberInput"] > div,
.stApp.stApp [data-testid="stDateInput"] > div,
.stApp.stApp [data-baseweb="input"],
.stApp.stApp [data-baseweb="textarea"],
.stApp.stApp [data-baseweb="select"] > div:first-child {
    background: var(--campo) !important;
    border: 1px solid var(--campo-bd) !important;
    border-radius: 8px !important;
    box-shadow: none !important;
}
/* invólucros e campos internos: transparentes e sem borda, senão volta a
   somar contorno sobre contorno e fundo sobre fundo */
.stApp.stApp [data-baseweb="input"] [data-baseweb="base-input"],
.stApp.stApp [data-baseweb="textarea"] [data-baseweb="base-input"],
.stApp.stApp [data-baseweb="select"] [data-baseweb="base-input"],
.stApp.stApp [data-testid="stTextInput"] input,
.stApp.stApp [data-testid="stNumberInput"] input,
.stApp.stApp [data-testid="stDateInput"] input,
.stApp.stApp [data-testid="stTextArea"] textarea {
    background: transparent !important;
    border: 0 !important;
    box-shadow: none !important;
}
/* ── LISTA DE CONCLUÍDOS ─────────────────────────────────────────────
   Uma linha por chamado, com botão próprio. Substituiu a grade porque no
   dataframe do Streamlit o clique numa célula NÃO seleciona a linha — só
   a caixinha da esquerda seleciona, e ninguém adivinha isso. */
.hcc-row{
    display:grid; grid-template-columns:3px 150px 1fr 78px 116px;
    align-items:center; gap:0 13px;
    padding:0 13px; height:58px; box-sizing:border-box; overflow:hidden;
    border:1px solid var(--n3); border-radius:8px; background:var(--n1);
    transition:border-color .15s ease, background .15s ease;
}
.hcc-row:hover{border-color:var(--ta380); background:var(--n2);}
.hcc-row.hcc-on{border-color:var(--ta550); background:var(--n2);}
.hcc-stripe{border-radius:2px; height:26px; width:3px;}
.hcc-crit {background:var(--f-crit-l);}
.hcc-alta {background:var(--f-acc);}
.hcc-media{background:var(--f-run);}
.hcc-baixa{background:rgba(var(--rgb-n8),.35);}
.hcc-quem{display:flex; flex-direction:column; gap:1px; min-width:0; overflow:hidden;}
.hcc-nome{font-size:.76rem; font-weight:700; color:var(--t-n11);
    white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}
.hcc-proto{font-family:var(--mono); font-size:.58rem; font-weight:700;
    letter-spacing:.02em; color:rgba(var(--rgb-c-acc),.75);
    white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}
.hcc-desc{font-size:.77rem; font-weight:600; color:var(--t-n10); line-height:1.4;
    display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;}
.hcc-sub{font-size:.66rem; color:rgba(var(--rgb-t-n8),.75); margin-top:1px;
    white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}
.hcc-chip{display:inline-flex; align-items:center; justify-content:center;
    padding:2px 8px; border-radius:5px; white-space:nowrap;
    font-size:.62rem; font-weight:800;}
.hcc-c-crit {background:rgba(var(--rgb-crit),.12); color:var(--c-crit-l);}
.hcc-c-alta {background:var(--ta160); color:var(--c-acc);}
.hcc-c-media{background:rgba(var(--rgb-run),.12); color:var(--c-run-l);}
.hcc-c-baixa{background:rgba(var(--rgb-n8),.10); color:var(--t-n8);}
.hcc-quando{font-family:var(--mono); font-size:.66rem; font-weight:700;
    color:rgba(var(--rgb-t-n9),.72); text-align:right; white-space:nowrap;}
.hcc-nada{padding:18px; text-align:center; border:1px dashed var(--n3);
    border-radius:8px; font-size:.78rem; color:rgba(var(--rgb-t-n8),.75);}
/* botão com a mesma altura da linha, para a lista ficar alinhada */
[data-testid="stExpander"] [data-testid="stHorizontalBlock"]:has(.hcc-row) [data-testid="stButton"] button{
    height:58px; font-size:.74rem; font-weight:700;
}

/* ── CHAMADO ABERTO NA BASE DE CONCLUÍDOS ────────────────────────────
   Cartão de leitura que aparece abaixo da grade quando uma linha é
   escolhida. Usa os mesmos tokens do resto do portal, então acompanha
   os dois temas sem regra própria de cor. */
.dch{
    margin:8px 0 12px; padding:15px 17px; border-radius:10px;
    background:var(--n1); border:1px solid var(--n3);
    border-left:3px solid var(--f-ok);
    box-shadow:0 1px 2px var(--shc-4), 0 10px 26px var(--shc-3);
}
.dch-topo{display:flex; align-items:center; gap:9px; flex-wrap:wrap; margin-bottom:13px;}
.dch-proto{
    font-family:var(--mono); font-size:12.5px; font-weight:700;
    letter-spacing:.02em; color:var(--t-n11); user-select:all;
}
.dch-chip{
    display:inline-flex; align-items:center; padding:2px 9px; border-radius:999px;
    font-size:10.5px; font-weight:800; letter-spacing:.04em; white-space:nowrap;
}
.dch-crit {background:rgba(var(--rgb-crit),.13); color:var(--c-crit-l);}
.dch-alta {background:var(--ta160); color:var(--c-acc);}
.dch-media{background:rgba(var(--rgb-run),.13); color:var(--c-run-l);}
.dch-baixa{background:rgba(var(--rgb-n8),.12); color:var(--t-n8);}
.dch-ok   {background:rgba(var(--rgb-ok),.13); color:var(--c-ok);}

.dch-grade{
    display:grid; grid-template-columns:repeat(auto-fit, minmax(148px, 1fr));
    gap:11px 16px; padding-bottom:13px; margin-bottom:12px;
    border-bottom:1px solid var(--n3);
}
.dch-campo{min-width:0;}
.dch-rot{
    font-family:var(--mono); font-size:9.5px; font-weight:800;
    letter-spacing:.13em; text-transform:uppercase;
    color:rgba(var(--rgb-t-n8),.85); margin-bottom:3px;
}
.dch-val{font-size:.82rem; font-weight:600; color:var(--t-n11);
    overflow-wrap:anywhere; line-height:1.35;}
.dch-bloco{margin-top:11px;}
.dch-corpo{
    font-size:.85rem; color:var(--t-n10); line-height:1.55;
    overflow-wrap:anywhere; white-space:pre-wrap; margin-top:4px;
}
.dch-resp .dch-corpo{
    padding:10px 12px; border-radius:8px;
    background:var(--n2); border-left:2px solid var(--f-ok);
}
.dch-vazio{color:rgba(var(--rgb-t-n8),.8); font-style:italic;}

/* ── LISTAS SUSPENSAS ────────────────────────────────────────────────
   O menu do BaseWeb é montado num portal FORA do .stApp — por isso nenhuma
   regra com prefixo de app o alcançava, e ele aparecia com o fundo da
   PÁGINA, sem borda, sem raio e sem sombra: um retângulo chapado que
   vazava por cima do cartão. Aqui ele vira um cartão de verdade, com os
   itens em pílula e o selecionado marcado por tinta + traço âmbar em vez
   de uma barra sólida. */
[data-testid="stSelectboxVirtualDropdown"],
[data-testid="stMultiSelectVirtualDropdown"],
[data-baseweb="popover"] ul[role="listbox"],
[data-baseweb="menu"] {
    background: var(--n1) !important;
    border: 1px solid var(--n3) !important;
    border-radius: 10px !important;
    /* sem padding próprio: a lista é VIRTUALIZADA — o Streamlit calcula a
       posição de cada item em pixels. Qualquer padding/altura que eu
       imponha aqui desalinha as linhas e abre buracos entre elas. */
    padding: 0 !important;
    box-shadow: 0 2px 6px var(--shc-4), 0 16px 40px var(--shc-2) !important;
    scrollbar-width: thin;
}
/* o invólucro do portal já não precisa de sombra própria */
[data-baseweb="popover"] > div { box-shadow: none !important; background: transparent !important; }

li[role="option"] {
    /* Só cor e realce. NADA de `position` aqui: o virtualizador posiciona
       cada linha por estilo inline, e forçar `position:relative` tira os
       itens do lugar — vira uma lista com o dobro do espaçamento. Pelo
       mesmo motivo o marcador do item ativo é um inset shadow, não um
       ::before (que exigiria posicionamento). */
    color: var(--t-n10) !important;
    transition: background .12s ease, color .12s ease !important;
}
li[role="option"]:hover { background: var(--n2) !important; color: var(--t-n11) !important; }
li[role="option"][aria-selected="true"] {
    background: var(--ta140) !important;
    color: var(--t-n11) !important;
    font-weight: 700 !important;
    box-shadow: inset 3px 0 0 var(--f-acc) !important;
}

/* seta do seletor fechado */
.stApp.stApp [data-baseweb="select"] svg { fill: var(--t-n8) !important; }
.stApp.stApp [data-baseweb="select"]:hover svg { fill: var(--c-acc) !important; }
/* altura uniforme: sem isto um seletor tinha 38px e o vizinho 46px */
.stApp.stApp [data-baseweb="select"] > div:first-child { min-height: 42px !important; }

/* fichas do multiselect */
.stApp.stApp [data-baseweb="tag"] {
    background: var(--ta140) !important;
    border: 1px solid var(--ta300) !important;
    border-radius: 6px !important;
    color: var(--c-acc) !important;
}
.stApp.stApp [data-baseweb="tag"] svg { fill: var(--c-acc) !important; }

.stApp.stApp [data-testid="stTextInput"] > div:focus-within,
.stApp.stApp [data-testid="stTextArea"] > div:focus-within,
.stApp.stApp [data-testid="stNumberInput"] > div:focus-within,
.stApp.stApp [data-testid="stDateInput"] > div:focus-within,
.stApp.stApp [data-baseweb="input"]:focus-within,
.stApp.stApp [data-baseweb="textarea"]:focus-within,
.stApp.stApp [data-baseweb="select"] > div:first-child:focus-within {
    border-color: var(--ta600) !important;
    box-shadow: 0 0 0 3px var(--ta120) !important;
}

</style>
"""

st.markdown(
    """
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link rel="stylesheet"
          href="https://fonts.googleapis.com/css2?family=Archivo:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap">
    """,
    unsafe_allow_html=True,
)
def _injetar_css_global_uma_vez():
    _css_html = _minify_css(_get_global_css(_tema_atual()))
    st.markdown(_css_html, unsafe_allow_html=True)

_injetar_css_global_uma_vez()

st.markdown("""
<style>
* { scrollbar-width: thin !important; scrollbar-color: var(--ta550) transparent !important; }
</style>
<script>
(function injectScrollbarCSS() {
    // Guard de instância única: o Streamlit re-renderiza este bloco a cada
    // rerun. Sem o guard, cada clique criava um novo MutationObserver +
    // setInterval que se acumulavam pela sessão, deixando o app cada vez
    // mais lento. Agora instala UMA vez só.
    if (window.__sdScrollbarReady) return;
    window.__sdScrollbarReady = true;

    const css = `
        * { scrollbar-width: thin; scrollbar-color: var(--ta550) transparent; }
        ::-webkit-scrollbar { width: 5px !important; height: 5px !important; }
        ::-webkit-scrollbar-track { background: rgba(var(--rgb-n1),.40) !important; border-radius: 99px !important; }
        ::-webkit-scrollbar-thumb {
            background: linear-gradient(180deg, var(--f-acc) 0%, var(--ta600) 100%) !important;
            border-radius: 99px !important;
            box-shadow: none;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: var(--f-acc-l) !important;
            box-shadow: none;
        }
        ::-webkit-scrollbar-corner { background: transparent !important; }
    `;
    function applyToDoc(doc) {
        try {
            if (!doc.getElementById('_sd_scrollbar_fix')) {
                const s = doc.createElement('style');
                s.id = '_sd_scrollbar_fix';
                s.textContent = css;
                (doc.head || doc.documentElement).appendChild(s);
            }
        } catch(e) {}
    }
    function scanIframes() {
        document.querySelectorAll('iframe').forEach(function(f) {
            try { applyToDoc(f.contentDocument || f.contentWindow.document); } catch(e) {}
        });
    }
    // Debounce: o observer dispara a cada mutação do DOM (constantes no
    // Streamlit). Sem debounce, varria todos os iframes centenas de vezes
    // por segundo. Agora agrupa via requestAnimationFrame.
    let _scheduled = false;
    function scheduleScan() {
        if (_scheduled) return;
        _scheduled = true;
        requestAnimationFrame(function() { _scheduled = false; scanIframes(); });
    }
    applyToDoc(document);
    scanIframes();
    try {
        new MutationObserver(scheduleScan).observe(document.body, {childList:true, subtree:true});
    } catch(e) {}
})();
</script>
""", unsafe_allow_html=True)

st.markdown("""
<style>
html.sd-anim-paused *,
html.sd-anim-paused *::before,
html.sd-anim-paused *::after { animation-play-state: paused !important; }
@media (prefers-reduced-motion: reduce) {
    *, *::before, *::after {
        animation-duration: .001ms !important;
        animation-iteration-count: 1 !important;
        transition-duration: .001ms !important;
    }
}
</style>
<script>
(function pauseAnimsWhenHidden(){
    if (window.__sdVisReady) return;
    window.__sdVisReady = true;
    function sync(){
        var h = document.documentElement;
        if (document.hidden) h.classList.add('sd-anim-paused');
        else h.classList.remove('sd-anim-paused');
    }
    document.addEventListener('visibilitychange', sync);
    sync();
})();
</script>
""", unsafe_allow_html=True)

st.markdown("""
<style>
.wz{display:flex;align-items:center;padding:6px 0 20px;user-select:none}
.wz-node{display:flex;flex-direction:column;align-items:center;gap:5px;flex:0 0 auto}
.wz-circle{width:30px;height:30px;border-radius:50%;display:flex;align-items:center;
  justify-content:center;font-size:12px;font-weight:800;transition:all .25s}
.wz-c-done{background:var(--f-ok);color:var(--t-n11);box-shadow: none}
.wz-c-act{background:var(--n2);border:2px solid var(--c-acc);color:var(--c-acc);box-shadow: none}
.wz-c-idle{background:var(--n1);border:2px solid var(--n3);color:var(--t-n3)}
.wz-lbl{font-size:10px;font-weight:700;letter-spacing:.02em;white-space:nowrap}
.wz-lbl-done{color:var(--c-ok)}
.wz-lbl-act{color:var(--t-n10)}
.wz-lbl-idle{color:var(--t-n3)}
.wz-line{flex:1;height:2px;margin:0 8px;margin-bottom:15px;border-radius:2px}
.wz-line-done{background:var(--f-ok)}
.wz-line-idle{background:var(--n3)}
.hz{background:linear-gradient(135deg,var(--sup-a),var(--sup-b));border:1px solid var(--n3);
  border-radius:12px;padding:18px 22px;position:relative;overflow:hidden;margin-bottom:14px;
  box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.05), 0 10px 22px var(--shc-2), 0 2px 5px var(--shc-2)}
.hz::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;
  background:linear-gradient(90deg,var(--f-acc),var(--f-run-l));box-shadow: none}
.hz-eye{font-size:9px;font-weight:800;letter-spacing:.15em;text-transform:uppercase;color:var(--t-n5);margin-bottom:5px}
.hz-title{font-size:1.1rem;font-weight:800;color:var(--t-n11);margin-bottom:3px}
.hz-desc{font-size:.78rem;color:var(--t-n6);line-height:1.5}
.hz-crumb{font-size:10px;color:var(--t-n5);margin-bottom:5px}
.hz-crumb b{color:var(--t-n7)}
.prio-chip{display:inline-flex;align-items:center;gap:7px;padding:6px 14px;
  border-radius:8px;font-size:11px;font-weight:700;margin:6px 0 2px}
.prio-dot{width:7px;height:7px;border-radius:50%;flex-shrink:0}
.char-c{font-size:9px;text-align:right;margin-top:3px}
.ok-card{background:var(--n1);border:1px solid var(--n3);border-left:3px solid var(--f-ok);
  border-radius:6px;padding:14px 16px;margin:2px 0 14px;}
.ok-topo{display:flex;align-items:center;gap:9px;flex-wrap:wrap}
.ok-titulo{font-size:14px;font-weight:600;color:var(--t-n11)}
.ok-proto{font-family:var(--mono);font-size:11.5px;color:var(--t-n9);background:var(--n2);
  border:1px solid var(--n3);border-radius:4px;padding:2px 8px;letter-spacing:.04em;user-select:all}
.ok-sub{font-size:12.5px;color:var(--t-n7);margin-top:7px;line-height:1.5}
.ok-fluxo{display:flex;align-items:center;gap:9px;margin-top:12px;flex-wrap:wrap}
.ok-etapa{display:flex;align-items:center;gap:6px;font-size:11px;color:var(--t-n7)}
.ok-etapa::before{content:"";width:6px;height:6px;border-radius:50%;background:var(--n4)}
.ok-etapa.feita{color:var(--c-ok)}
.ok-etapa.feita::before{background:var(--f-ok)}
.ok-sep{width:16px;height:1px;background:var(--n4)}
.wz-c-act{background:var(--n1);border:2px solid var(--c-acc);color:var(--c-acc);box-shadow: none}
.hzp{background:linear-gradient(135deg,var(--sup2-a),var(--sup2-b));border:1px solid var(--n3);
  border-radius:12px;padding:18px 22px;position:relative;overflow:hidden;margin-bottom:14px;
  box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.05), 0 10px 22px var(--shc-2), 0 2px 5px var(--shc-2)}
.hzp::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;
  background:linear-gradient(90deg,var(--f-acc),var(--f-acc-d));box-shadow: none}
.hzp-eye{font-size:9px;font-weight:800;letter-spacing:.15em;text-transform:uppercase;color:var(--t-n3);margin-bottom:5px}
.hzp-title{font-size:1.1rem;font-weight:800;color:var(--t-n11);margin-bottom:3px}
.hzp-desc{font-size:.78rem;color:var(--t-n5);line-height:1.5}
.hzp-crumb{font-size:10px;color:var(--t-n4);margin-bottom:5px}
.hzp-crumb b{color:var(--t-n7)}
.prio-chip2{display:inline-flex;align-items:center;gap:7px;padding:6px 14px;
  border-radius:8px;font-size:11px;font-weight:700;margin:6px 0 2px}
.prio-dot2{width:7px;height:7px;border-radius:50%;flex-shrink:0}
.ctx-panel{background:linear-gradient(160deg,rgba(var(--rgb-sup-a),.92),rgba(var(--rgb-sup-b),.85));
          border:1px solid var(--ta140);border-radius:14px;padding:20px 20px 18px;
          box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.05), 0 12px 28px var(--shc-2), 0 2px 6px var(--shc-2);position:relative;overflow:hidden}
.ctx-panel::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;
          background:linear-gradient(90deg,var(--f-acc),var(--f-acc));box-shadow: none}
.ctx-h{font-family:'IBM Plex Mono',monospace;font-size:.62rem;font-weight:800;letter-spacing:.16em;
          text-transform:uppercase;color:var(--c-acc);margin-bottom:16px}
.ctx-step{display:flex;gap:12px;align-items:flex-start;margin-bottom:14px}
.ctx-n{flex:0 0 24px;width:24px;height:24px;border-radius:50%;display:flex;align-items:center;
          justify-content:center;font-size:11px;font-weight:800;color:var(--c-acc);
          background:var(--ta100);border:1px solid var(--ta300)}
.ctx-step strong{display:block;font-size:.84rem;color:var(--t-n11);font-weight:700;margin-bottom:2px}
.ctx-step span{font-size:.74rem;color:var(--t-n8);line-height:1.45}
.ctx-divider{height:1px;background:linear-gradient(90deg,transparent,var(--ta250),transparent);margin:4px 0 14px}
.ctx-tip{font-size:.76rem;color:var(--c-acc-l);background:var(--ta060);
          border:1px solid var(--ta130);border-radius:9px;padding:10px 12px;line-height:1.5}
.cat3d:hover{
          transform:translateY(-3px)!important;
          border-color:var(--ta550)!important;
          box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.08), 0 14px 26px var(--shc-1);
        }
</style>
""", unsafe_allow_html=True)

ANEXO_DIR = os.path.join(_DADOS, "anexos")
os.makedirs(ANEXO_DIR, exist_ok=True)

DATABASE = os.path.join(_DADOS, "chamados.csv")
DATABASE_COMPRAS = os.path.join(_DADOS, "solicitacoes_compras.csv")
BACKUP_DIR = os.path.join(_DADOS, "backups")

COLUNAS = [
    "Data", "Protocolo", "Usuário", "Setor", "Categoria", "Impacto", "Prioridade",
    "Descrição", "Status", "Resposta do Técnico", "Técnico Responsável",
    "Atualizado em", "Anexo", "Anexo Resposta"
]
COLUNAS_COMPRAS = [
    "Data", "Protocolo", "Usuário", "Setor", "Categoria", "Item", "Quantidade",
    "Justificativa", "Urgência", "Status", "Observação TI", "Responsável",
    "Atualizado em"
]

DATABASE_SQLITE = os.path.join(_DADOS, "servicedesk.db")
_DB_ANTIGO = os.path.join(_RAIZ, "servicedesk.db")
if not os.path.exists(DATABASE_SQLITE) and os.path.exists(_DB_ANTIGO):
    DATABASE_SQLITE = _DB_ANTIGO
TABLE_CHAMADOS  = "chamados"
TABLE_COMPRAS   = "compras"


def _conectar():
    """Conexão SQLite centralizada.

    - busy_timeout: em vez de falhar com 'database is locked' quando dois
      acessos coincidem, espera até 30s pela liberação.

    Journaling no modo padrão (DELETE): um único arquivo .db entre transações.
    NÃO usamos WAL — nesta pasta sincronizada o WAL deixava arquivos -wal/-shm
    que a sincronização corrompia, deixando o banco 'malformed'.

    Uma conexão por operação (não global): o sqlite3 deste Python tem
    threadsafety=1, e o custo de abrir conexão é desprezível.
    """
    import sqlite3 as _sq3
    _conn = _sq3.connect(DATABASE_SQLITE, timeout=30)
    _conn.execute("PRAGMA busy_timeout=30000")
    return _conn


def obter_token_admin():
    try:
        token_secret = st.secrets.get("PORTAL_SUPORTE_ADMIN_TOKEN", "")
        if token_secret:
            return str(token_secret)
    except Exception:
        pass
    return os.getenv("PORTAL_SUPORTE_ADMIN_TOKEN", "admin123")


def validar_token_admin(token_digitado):
    return hmac.compare_digest(str(token_digitado or ""), obter_token_admin())


_TTL_SESSAO_TI = 3600


def sessao_ti_ativa() -> bool:
    """True se o analista está autenticado e a sessão não expirou."""
    if not st.session_state.get("gestor_autenticado"):
        return False
    if (time.time() - st.session_state.get("gestor_auth_ts", 0.0)) > _TTL_SESSAO_TI:
        st.session_state["gestor_autenticado"] = False
        st.session_state["gestor_auth_ts"] = 0.0
        return False
    return True


def abrir_sessao_ti(token: str) -> bool:
    """Valida o token e abre a sessão. Devolve False se o token não confere."""
    if not validar_token_admin(token):
        return False
    st.session_state["gestor_autenticado"] = True
    st.session_state["gestor_auth_ts"] = time.time()
    st.session_state["gestor_tentativas"] = 0
    st.session_state["gestor_lockout_ate"] = 0.0
    return True


def token_padrao_ativo():
    return obter_token_admin() == "admin123"


def garantir_colunas(df, colunas):
    df_ok = df.copy()
    for coluna in colunas:
        if coluna not in df_ok.columns:
            df_ok[coluna] = ""
    extras = [coluna for coluna in df_ok.columns if coluna not in colunas]
    return df_ok[colunas + extras]


_SAN_CTRL_RE  = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_SAN_QUOTE_RE = re.compile(r'"{3,}')

def _sanitizar_celula(v):
    if v is None:
        return ""
    txt = str(v)
    txt = txt.replace("\r\n", "\n").replace("\r", "\n")
    txt = _SAN_CTRL_RE.sub("", txt)
    txt = _SAN_QUOTE_RE.sub("", txt)
    txt = re.sub(r"\n{3,}", "\n\n", txt)
    return txt.strip()


def salvar_dataframe_seguro(df, caminho, colunas):
    """Persiste o DataFrame no SQLite — fonte única de verdade.

    (O snapshot automático em CSV foi removido: o banco SQLite é o único
    armazenamento agora. Isso evita o crescimento descontrolado de backups/.)
    """
    df_ok = garantir_colunas(df, colunas)

    for _col in df_ok.columns:
        df_ok[_col] = df_ok[_col].map(_sanitizar_celula)

    tabela = TABLE_CHAMADOS if caminho == DATABASE else TABLE_COMPRAS
    _conn  = _conectar()
    df_ok.to_sql(tabela, _conn, if_exists="replace", index=False)
    if tabela == TABLE_CHAMADOS:
        _conn.execute(f'CREATE INDEX IF NOT EXISTS idx_ch_status ON "{TABLE_CHAMADOS}" (Status)')
        _conn.execute(f'CREATE INDEX IF NOT EXISTS idx_ch_prio   ON "{TABLE_CHAMADOS}" (Prioridade)')
    else:
        _conn.execute(f'CREATE INDEX IF NOT EXISTS idx_cp_status ON "{TABLE_COMPRAS}" (Status)')
        _conn.execute(f'CREATE INDEX IF NOT EXISTS idx_cp_urg    ON "{TABLE_COMPRAS}" (Urgência)')
    _conn.commit()
    _conn.close()
    carregar_dados_cacheados.clear()


def gerar_protocolo(prefixo):
    return f"{prefixo}-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:4].upper()}"

def _ui_skeleton_html(context_label: str = "Atualizando...") -> str:
    return f"""
    <div class="skeleton-wrap" aria-busy="true">
        <div class="skeleton-label">{html.escape(context_label)}</div>
        <div class="skeleton-grid">
            {"".join([
                '<div class="skeleton-card">'
                '<div class="skeleton-line lg"></div>'
                '<div class="skeleton-line md"></div>'
                '<div class="skeleton-line sm"></div>'
                '</div>'
                for _ in range(4)
            ])}
        </div>
        <div class="skeleton-rail">
            {"".join([
                '<div class="skeleton-queue">'
                '<div class="skeleton-block"></div>'
                '<div class="skeleton-block"></div>'
                '<div class="skeleton-block"></div>'
                '<div class="skeleton-block"></div>'
                '</div>'
                for _ in range(2)
            ])}
        </div>
    </div>
    """

def trigger_ui_skeleton(context_label: str = "Atualizando...", duration_s: float = 0.25) -> None:
    st.session_state["ui_skeleton_until"] = time.time() + float(duration_s)
    st.session_state["ui_skeleton_label"] = context_label

def maybe_show_ui_skeleton() -> None:
    st.session_state.pop("ui_skeleton_until", None)
    st.session_state.pop("ui_skeleton_label", None)


APP_HOST = "192.168.1.222"
APP_PORT = 8501

WHATSAPP_DESTINO = "554830293440"


def _url_pagina(page: str) -> str:
    """Retorna a URL completa da página para ser linkada no WhatsApp."""
    return f"http://{APP_HOST}:{APP_PORT}/?page={page}"


def _wa_config():
    """Lê o número de destino e a apikey do CallMeBot de secrets.toml,
    com fallback para a constante padrão. Nunca lança exceção."""
    destino = WHATSAPP_DESTINO
    apikey = ""
    try:
        destino = str(st.secrets.get("WHATSAPP_DESTINO", destino)) or destino
        apikey = str(st.secrets.get("CALLMEBOT_APIKEY", ""))
    except Exception:
        pass
    return destino.strip(), apikey.strip()


def notificar_whatsapp(titulo: str, corpo: str, url: str | None = None):
    """
    Envia uma notificação para o WhatsApp da TI.

    Estratégia (em ordem):
      1) Se CALLMEBOT_APIKEY estiver definido em secrets.toml -> envia via
         CallMeBot por HTTP. Rápido, não abre navegador e não trava o portal.
      2) Caso contrário -> usa pywhatkit (abre o WhatsApp Web no navegador
         padrão; exige sessão logada). Requisito: pip install pywhatkit

    O envio roda numa thread separada para nunca segurar a interface do
    Streamlit. Falhas são silenciosas: a notificação é "best effort" e
    jamais derruba o app.
    """
    destino, apikey = _wa_config()

    linhas = [f"*{titulo}*", ""]
    linhas += corpo.strip().splitlines()
    if url:
        linhas += ["", url]
    mensagem = "\n".join(linhas)

    def _enviar():
        if apikey:
            try:
                import urllib.parse
                import urllib.request
                phone = destino if destino.startswith("+") else "+" + destino
                qs = urllib.parse.urlencode({
                    "phone": phone,
                    "text": mensagem,
                    "apikey": apikey,
                })
                req = urllib.request.Request(
                    "https://api.callmebot.com/whatsapp.php?" + qs,
                    headers={"User-Agent": "PortalSuporteTI/1.0"},
                )
                urllib.request.urlopen(req, timeout=20).read()
                return
            except Exception:
                pass

        try:
            import pywhatkit as pwk
            try:
                pwk.sendwhatmsg_instantly(
                    destino, mensagem,
                    wait_time=15, tab_close=True, close_time=3,
                )
            except AttributeError:
                import datetime as _dt
                agora = _dt.datetime.now()
                pwk.sendwhatmsg(
                    destino, mensagem, agora.hour, (agora.minute + 1) % 60,
                    wait_time=15, tab_close=True, close_time=3,
                )
        except Exception:
            pass

    try:
        import threading
        threading.Thread(target=_enviar, daemon=True).start()
    except Exception:
        _enviar()


def notificar_windows(titulo, mensagem="Nova atividade registrada"):
    import subprocess
    t = titulo.replace('"', "'").replace('<', '').replace('>', '').replace('&', 'e')
    m = mensagem.replace('"', "'").replace('<', '').replace('>', '').replace('&', 'e').replace('\n', ' | ')

    ps_script = f"""
Add-Type -AssemblyName System.Windows.Forms
$n = New-Object System.Windows.Forms.NotifyIcon
$n.Icon = [System.Drawing.SystemIcons]::Information
$n.Visible = $true
$n.BalloonTipTitle = "{t}"
$n.BalloonTipText  = "{m}"
$n.BalloonTipIcon  = [System.Windows.Forms.ToolTipIcon]::Warning
$n.ShowBalloonTip(7000)
Start-Sleep -Milliseconds 7500
$n.Dispose()
"""
    try:
        flags = 0x08000000
        subprocess.Popen(
            ["powershell", "-WindowStyle", "Hidden", "-NonInteractive", "-Command", ps_script],
            creationflags=flags,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
    except Exception:
        pass


def _backup_diario_db(manter: int = 14):
    """Cópia de segurança diária do banco SQLite em backups/.

    Substitui o antigo snapshot em CSV. Usa a API de backup do SQLite, que
    gera uma cópia consistente mesmo com o WAL ativo. Roda no máximo uma vez
    por dia (se a cópia do dia já existir, não faz nada) e mantém apenas as
    'manter' cópias mais recentes.
    """
    try:
        if not os.path.exists(DATABASE_SQLITE):
            return
        os.makedirs(BACKUP_DIR, exist_ok=True)
        destino = os.path.join(BACKUP_DIR, f"servicedesk_{datetime.now():%Y%m%d}.db")
        if os.path.exists(destino):
            return
        import sqlite3 as _sq3
        _src = _sq3.connect(DATABASE_SQLITE)
        _dst = _sq3.connect(destino)
        with _dst:
            _src.backup(_dst)
        _dst.close()
        _src.close()
        copias = sorted(
            f for f in os.listdir(BACKUP_DIR)
            if f.startswith("servicedesk_") and f.endswith(".db")
        )
        for antigo in copias[:-manter]:
            try:
                os.remove(os.path.join(BACKUP_DIR, antigo))
            except OSError:
                pass
    except Exception:
        pass


@st.cache_resource(show_spinner=False)
def _inicializar_banco():
    """
    Cria as tabelas SQLite com índices e migra CSVs existentes.
    Executa uma única vez por sessão do servidor (cache_resource).
    Retorna True se houve migração de dados legados.
    """
    _conn = _conectar()

    try:
        _conn.execute("PRAGMA journal_mode=DELETE")
    except Exception:
        pass

    _conn.execute(f"""
        CREATE TABLE IF NOT EXISTS "{TABLE_CHAMADOS}" (
            Data                   TEXT,
            Protocolo              TEXT,
            Usuário                TEXT,
            Setor                  TEXT,
            Categoria              TEXT,
            Impacto                TEXT,
            Prioridade             TEXT,
            Descrição              TEXT,
            Status                 TEXT,
            "Resposta do Técnico"  TEXT,
            "Técnico Responsável"  TEXT,
            "Atualizado em"        TEXT,
            Anexo                  TEXT
        )
    """)

    _conn.execute(f"""
        CREATE TABLE IF NOT EXISTS "{TABLE_COMPRAS}" (
            Data               TEXT,
            Protocolo          TEXT,
            Usuário            TEXT,
            Setor              TEXT,
            Categoria          TEXT,
            Item               TEXT,
            Quantidade         TEXT,
            Justificativa      TEXT,
            Urgência           TEXT,
            Status             TEXT,
            "Observação TI"    TEXT,
            Responsável        TEXT,
            "Atualizado em"    TEXT
        )
    """)

    _conn.execute(f'CREATE INDEX IF NOT EXISTS idx_ch_status ON "{TABLE_CHAMADOS}" (Status)')
    _conn.execute(f'CREATE INDEX IF NOT EXISTS idx_ch_prio   ON "{TABLE_CHAMADOS}" (Prioridade)')
    _conn.execute(f'CREATE INDEX IF NOT EXISTS idx_cp_status ON "{TABLE_COMPRAS}"  (Status)')
    _conn.execute(f'CREATE INDEX IF NOT EXISTS idx_cp_urg    ON "{TABLE_COMPRAS}"  (Urgência)')
    _conn.commit()
    _conn.close()


def carregar_dados(caminho, colunas):
    """Carrega dados do SQLite via pandas (substituiu leitura direta de CSV)."""
    tabela = TABLE_CHAMADOS if caminho == DATABASE else TABLE_COMPRAS
    try:
        _conn = _conectar()
        _df   = pd.read_sql(f'SELECT * FROM "{tabela}"', _conn)
        _conn.close()
        if _df.empty:
            return pd.DataFrame(columns=colunas)
        _df = garantir_colunas(_df, colunas)
        if "Prioridade" in _df.columns:
            _df["Prioridade"] = _df["Prioridade"].replace({"Urgente": "Crítica"})
        return _df
    except Exception:
        return pd.DataFrame(columns=colunas)


_inicializar_banco()

@st.cache_resource(show_spinner=False)
def _backup_diario_db_once(_dia: str):
    """Garante que o backup diário rode no máximo 1x por dia por sessão do
    servidor, evitando checagens de disco a cada rerun (cada clique reexecuta
    o script inteiro no Streamlit). A chave _dia invalida o cache à meia-noite."""
    _backup_diario_db()

_backup_diario_db_once(datetime.now().strftime("%Y%m%d"))

df_base    = carregar_dados_cacheados(DATABASE, COLUNAS)
df_compras = carregar_dados_cacheados(DATABASE_COMPRAS, COLUNAS_COMPRAS)


def aplicar_badges_chamados(df):
    df_visual = df.copy()
    if not df_visual.empty:
        status_map = {"Aberto": "⚪ Aberto", "Em Atendimento": "🟡 Em Atendimento", "Concluído": "🟢 Concluído"}
        prio_map = {"Baixa": "🔵 Baixa", "Média": "🟡 Média", "Alta": "🔴 Alta", "Urgente": "🚨 Urgente"}
        df_visual["Status"] = df_visual["Status"].map(status_map).fillna(df_visual["Status"])
        df_visual["Prioridade"] = df_visual["Prioridade"].map(prio_map).fillna(df_visual["Prioridade"])
    return df_visual


def aplicar_badges_compras(df):
    df_visual = df.copy()
    if not df_visual.empty:
        status_map = {
            "Pendente": "⏳ Pendente", "Em Análise": "🔍 Em Análise",
            "Aprovado": "✅ Aprovado", "Recusado": "❌ Recusado", "Concluído": "🟢 Concluído"
        }
        urg_map = {
            "Sem pressa": "🌱 Sem pressa", "Moderada": "🟡 Moderada",
            "Alta": "🔴 Alta", "Crítica": "🚨 Crítica"
        }
        df_visual["Status"] = df_visual["Status"].map(status_map).fillna(df_visual["Status"])
        df_visual["Urgência"] = df_visual["Urgência"].map(urg_map).fillna(df_visual["Urgência"])
    return df_visual


def normalizar_texto(texto):
    return str(texto).encode('ascii', 'ignore').decode('ascii')


def escapar_html(valor, padrao="—"):
    if pd.isna(valor):
        return padrao
    texto = str(valor).strip()
    return html.escape(texto if texto else padrao)


def formatar_id(indice):
    try:
        return f"{int(indice):03d}"
    except (TypeError, ValueError):
        return str(indice)


def horas_decorridas(data_texto):
    try:
        abertura = datetime.strptime(str(data_texto), "%d/%m/%Y %H:%M")
        delta = datetime.now() - abertura
        return max(delta.total_seconds() / 3600, 0)
    except Exception:
        return None


def tempo_decorrido(data_texto):
    horas = horas_decorridas(data_texto)
    if horas is None:
        return "sem data"
    minutos = int(horas * 60)
    if minutos < 1:
        return "agora"
    if minutos < 60:
        return f"{minutos} min"
    if horas < 24:
        return f"{int(horas)} h"
    return f"{int(horas // 24)} d"


def peso_prioridade(valor):
    pesos = {
        "Urgente": 5, "Crítica": 5, "Alta": 4, "Média": 3, "Moderada": 3,
        "Baixa": 2, "Sem pressa": 1
    }
    return pesos.get(str(valor), 0)


def classe_fila(valor):
    valor = str(valor)
    if valor in ["Urgente", "Crítica", "Alta"]:
        return "queue-critical"
    if valor in ["Média", "Moderada"]:
        return "queue-high"
    if valor == "Baixa":
        return "queue-medium"
    return "queue-low"


def prioridade_por_impacto(prioridade, impacto):
    if impacto == "Empresa parada" and prioridade in ["Baixa", "Média"]:
        return "Alta"
    if impacto == "Setor impactado" and prioridade == "Baixa":
        return "Média"
    return prioridade


def extrair_categoria(row):
    categoria = str(row.get("Categoria", "")).strip()
    if categoria:
        return categoria
    descricao = str(row.get("Descrição", "")).strip()
    if descricao.startswith("[") and "]" in descricao:
        return descricao[1:descricao.find("]")]
    return "Não classificado"


def primeiro_valor(row, *colunas, padrao="—"):
    for coluna in colunas:
        valor = row.get(coluna, "")
        if not pd.isna(valor) and str(valor).strip():
            return str(valor).strip()
    return padrao


_EXT_IMAGEM = (".png", ".jpg", ".jpeg")

def zona_colar(altura: int = 78):
    """Área que aceita Ctrl+V de imagem e joga o resultado no anexo acima.

    O st.file_uploader do Streamlit só abre o seletor de arquivos — não tem
    evento de colar. Aqui um iframe escuta o `paste` no documento PAI, monta
    um File a partir do que veio da área de transferência e injeta no
    <input type=file> do uploader mais próximo ACIMA desta zona, disparando
    o mesmo evento 'change' de quando se escolhe um arquivo. Ou seja: o
    upload segue o caminho normal do Streamlit, sem gambiarra de base64.
    """
    components.html(
        _resolver_vars("""
<!DOCTYPE html><html><head><meta charset="utf-8"><style>
  html,body{margin:0;padding:0;background:transparent;
    font-family:'Archivo',system-ui,'Segoe UI',Roboto,sans-serif}
  .zc{display:flex;align-items:center;gap:9px;padding:11px 13px;cursor:pointer;
    border:1px dashed var(--n4);border-radius:6px;background:var(--n2);
    color:var(--t-n8);font-size:12.5px;transition:border-color .15s,color .15s}
  .zc:hover,.zc.foco{border-color:var(--f-acc);color:var(--t-n10)}
  .zc.ok{border-color:var(--f-ok);color:var(--c-ok);border-style:solid}
  .zc.err{border-color:var(--f-crit-l);color:var(--c-crit-l)}
  .zc svg{flex:none}
  kbd{font-family:'IBM Plex Mono',monospace;font-size:11px;background:var(--n3);
    border:1px solid var(--n5);border-radius:4px;padding:1px 5px;color:var(--t-n10)}
</style></head><body>
<div class="zc" id="z" tabindex="0">
  <svg viewBox="0 0 20 20" width="16" height="16" fill="none" stroke="currentColor"
       stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round">
    <rect x="4.5" y="3.5" width="11" height="13" rx="1.6"/>
    <path d="M7.8 3.5V2.8a1 1 0 0 1 1-1h2.4a1 1 0 0 1 1 1v.7z"/>
    <path d="M10 8.4v4.2M7.9 10.5l2.1-2.1 2.1 2.1"/>
  </svg>
  <span id="txt">Tirou print? Clique aqui e cole com <kbd>Ctrl</kbd>+<kbd>V</kbd></span>
</div>
<script>
(function(){
  var z = document.getElementById('z'), txt = document.getElementById('txt');
  var pai = window.parent.document;

  function aviso(msg, classe){
    txt.innerHTML = msg;
    z.className = 'zc ' + (classe || '');
    if (classe !== 'ok') setTimeout(reset, 3500);
  }
  function reset(){
    z.className = 'zc';
    txt.innerHTML = 'Tirou print? Clique aqui e cole com <kbd>Ctrl</kbd>+<kbd>V</kbd>';
  }

  // O <input type=file> do uploader imediatamente ACIMA desta zona.
  function alvo(){
    var eu = window.frameElement;
    if (!eu) return null;
    var todos = pai.querySelectorAll('input[type=file]');
    var anterior = null;
    for (var i = 0; i < todos.length; i++){
      var pos = eu.compareDocumentPosition(todos[i]);
      if (pos & Node.DOCUMENT_POSITION_PRECEDING) anterior = todos[i];
    }
    return anterior;
  }

  function colar(ev){
    var dados = ev.clipboardData || window.clipboardData;
    if (!dados) return;
    var itens = dados.items || [];
    for (var i = 0; i < itens.length; i++){
      if (itens[i].kind !== 'file') continue;
      if (String(itens[i].type).indexOf('image/') !== 0) continue;
      var blob = itens[i].getAsFile();
      if (!blob) continue;
      var ext = (blob.type.split('/')[1] || 'png').replace('jpeg', 'jpg');
      var arquivo = new File([blob], 'colado_' + Date.now() + '.' + ext, {type: blob.type});
      var inp = alvo();
      if (!inp){ aviso('Não encontrei o campo de anexo.', 'err'); return; }
      var dt = new DataTransfer();
      dt.items.add(arquivo);
      inp.files = dt.files;
      inp.dispatchEvent(new Event('change', {bubbles: true}));
      ev.preventDefault();
      aviso('Imagem colada e anexada.', 'ok');
      return;
    }
    aviso('A área de transferência não tem imagem.', 'err');
  }

  // Escuta nos dois documentos: o foco pode estar na página ou nesta zona.
  pai.addEventListener('paste', colar, true);
  document.addEventListener('paste', colar, true);
  z.addEventListener('click', function(){ z.classList.add('foco'); z.focus(); });
})();
</script></body></html>
        """),
        height=altura,
    )


def _detalhe_chamado_concluido(row):
    """Cartão de leitura do chamado escolhido na base de concluídos.

    A grade só cabe sete colunas truncadas; a descrição do usuário e o
    parecer do técnico — que é o que interessa quando se consulta um
    chamado fechado — não apareciam em lugar nenhum sem abrir a janela
    de edição. Aqui o chamado abre embaixo da tabela, em modo leitura.
    """
    def _txt(campo, padrao="—"):
        v = str(row.get(campo, "") or "").strip()
        return escapar_html(v, padrao) if v else padrao

    _prio = str(row.get("Prioridade", "")).strip().lower()
    _cls_prio = ("crit" if ("cr" in _prio or "urg" in _prio)
                 else "alta" if "alt" in _prio
                 else "media" if _prio.startswith("m") else "baixa")

    _abriu = str(row.get("Data", "") or "").strip()
    _fechou = str(row.get("Atualizado em", "") or "").strip() or _abriu
    _dur = "—"
    try:
        _a = datetime.strptime(_abriu, "%d/%m/%Y %H:%M")
        _f = datetime.strptime(_fechou, "%d/%m/%Y %H:%M")
        _m = max(0, int((_f - _a).total_seconds() // 60))
        _dur = (f"{_m} min" if _m < 60
                else f"{_m // 60}h{_m % 60:02d}" if _m < 1440
                else f"{_m // 1440}d {(_m % 1440) // 60}h")
    except Exception:
        pass

    _campos = [
        ("Solicitante", _txt("Usuário")), ("Setor", _txt("Setor")),
        ("Categoria", escapar_html(extrair_categoria(row), "—")),
        ("Impacto", _txt("Impacto")),
        ("Técnico", _txt("Técnico Responsável", "Não atribuído")),
        ("Aberto em", escapar_html(_abriu, "—")),
        ("Concluído em", escapar_html(_fechou, "—")),
        ("Tempo até fechar", _dur),
    ]
    _meta = "".join(
        f'<div class="dch-campo"><div class="dch-rot">{r}</div>'
        f'<div class="dch-val">{v}</div></div>' for r, v in _campos
    )

    _resp = _txt("Resposta do Técnico", "")
    _bloco_resp = (
        f'<div class="dch-bloco dch-resp"><div class="dch-rot">Parecer técnico</div>'
        f'<div class="dch-corpo">{_resp}</div></div>'
        if _resp else
        '<div class="dch-bloco dch-resp"><div class="dch-rot">Parecer técnico</div>'
        '<div class="dch-corpo dch-vazio">Fechado sem parecer registrado.</div></div>'
    )

    st.markdown(
        f'<div class="dch">'
        f'<div class="dch-topo">'
        f'<span class="dch-proto">{_txt("Protocolo")}</span>'
        f'<span class="dch-chip dch-{_cls_prio}">{_txt("Prioridade")}</span>'
        f'<span class="dch-chip dch-ok">Conclu&iacute;do</span>'
        f'</div>'
        f'<div class="dch-grade">{_meta}</div>'
        f'<div class="dch-bloco"><div class="dch-rot">Descri&ccedil;&atilde;o do problema</div>'
        f'<div class="dch-corpo">{_txt("Descrição")}</div></div>'
        f'{_bloco_resp}'
        f'</div>',
        unsafe_allow_html=True,
    )


def renderizar_anexo(nome_arquivo, key=None):
    """Mostra o anexo de um chamado (imagem inline ou botão de download do PDF).

    Antes disso o anexo só aparecia como texto (nome do arquivo) — salvo em
    disco em anexos/, mas sem nenhum jeito de abrir/baixar pela tela do
    Gestor. Isso deixava a triagem sem acesso à evidência anexada pelo
    colaborador.
    """
    nome = str(nome_arquivo or "").strip()
    if not nome or nome.lower() in ("nenhum", "nan", "none", ""):
        st.caption("Sem anexo.")
        return

    caminho = os.path.join(ANEXO_DIR, nome)
    if not os.path.isfile(caminho):
        st.warning(f"Anexo registrado (**{nome}**) mas o arquivo não foi encontrado em `anexos/`.")
        return

    with open(caminho, "rb") as _f:
        dados = _f.read()

    _ext = os.path.splitext(nome)[1].lower()
    if _ext in _EXT_IMAGEM:
        st.image(dados, caption=nome, width=320)
    st.download_button(
        "Abrir/baixar anexo" if _ext not in _EXT_IMAGEM else "Baixar em tamanho original",
        data=dados,
        file_name=nome,
        key=key or f"dl_anexo_{nome}",
        use_container_width=False,
    )


def ordenar_por_prioridade_e_data(df, coluna_prioridade):
    if df.empty:
        return df
    df_ord = df.copy()
    df_ord["_peso"] = df_ord[coluna_prioridade].apply(peso_prioridade)
    df_ord["_data_ord"] = pd.to_datetime(df_ord["Data"], format="%d/%m/%Y %H:%M", errors="coerce")
    return df_ord.sort_values(["_peso", "_data_ord"], ascending=[False, True]).drop(columns=["_peso", "_data_ord"])


def gerar_pdf(df_geral):
    try:
        from fpdf import FPDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, "Relatorio de Chamados - TI", 0, 1, 'C')
        pdf.ln(10)
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(35, 10, "Data", 1); pdf.cell(30, 10, "Setor", 1)
        pdf.cell(30, 10, "Status", 1); pdf.cell(90, 10, "Descricao", 1, 1)
        pdf.set_font("Arial", size=9)
        for _, row in df_geral.iterrows():
            pdf.cell(35, 10, normalizar_texto(row.get('Data', 'N/A')), 1)
            pdf.cell(30, 10, normalizar_texto(row.get('Setor', 'N/A')), 1)
            pdf.cell(30, 10, normalizar_texto(row.get('Status', 'N/A')), 1)
            desc = normalizar_texto(row.get('Descrição', 'N/A'))[:45]
            pdf.cell(90, 10, desc, 1, 1)
        return pdf.output(dest='S').encode('latin-1')
    except Exception as e:
        return f"Erro ao gerar PDF: {e}".encode('latin-1')


def gerar_pdf_compras(df):
    try:
        from fpdf import FPDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, "Relatorio de Solicitacoes de Compras - TI", 0, 1, 'C')
        pdf.ln(10)
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(30, 10, "Data", 1); pdf.cell(28, 10, "Usuario", 1)
        pdf.cell(25, 10, "Setor", 1); pdf.cell(25, 10, "Categoria", 1)
        pdf.cell(50, 10, "Item", 1); pdf.cell(15, 10, "Qtd", 1); pdf.cell(22, 10, "Status", 1, 1)
        pdf.set_font("Arial", size=8)
        for _, row in df.iterrows():
            pdf.cell(30, 10, normalizar_texto(row.get('Data', 'N/A')), 1)
            pdf.cell(28, 10, normalizar_texto(row.get('Usuário', 'N/A'))[:18], 1)
            pdf.cell(25, 10, normalizar_texto(row.get('Setor', 'N/A')), 1)
            pdf.cell(25, 10, normalizar_texto(row.get('Categoria', 'N/A')), 1)
            item = normalizar_texto(row.get('Item', 'N/A'))[:30]
            pdf.cell(50, 10, item, 1)
            pdf.cell(15, 10, normalizar_texto(row.get('Quantidade', 'N/A')), 1)
            pdf.cell(22, 10, normalizar_texto(row.get('Status', 'N/A')), 1, 1)
        return pdf.output(dest='S').encode('latin-1')
    except Exception as e:
        return f"Erro ao gerar PDF: {e}".encode('latin-1')


with st.sidebar:
    st.markdown("""
        <div class="tlm-brand">
            <div class="tlm-mark">TLM</div>
            <div>
                <div class="tlm-name">IT Service Desk</div>
                <div class="tlm-sub">Suporte TI &middot; v3.8</div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    components.html(f"""
    <!DOCTYPE html><html><head><meta charset="utf-8">
    <style>
      {_tokens_vars(_tema_atual())}
      html,body{{margin:0;padding:0;background:transparent;}}
      .strip{{
        display:flex;align-items:center;gap:8px;
        margin-top:2px;padding:8px 10px;
        border:1px solid var(--ta220);
        border-radius:10px;background:rgba(var(--rgb-n0),.45);
        color:var(--t-n9);font-size:11px;
        font-family:'IBM Plex Mono',ui-monospace,SFMono-Regular,monospace;
        text-transform:uppercase;letter-spacing:.08em;
        box-shadow: none;overflow:hidden;
      }}
      .strip .dot{{width:7px;height:7px;border-radius:999px;background:var(--f-ok);
        box-shadow: none;display:inline-block;
        animation: none;}}
      .strip strong{{color:var(--c-acc-xl);font-weight:700;}}
      .strip .up{{color:var(--c-acc);font-weight:700;font-variant-numeric:tabular-nums;margin-left:auto;}}
      @keyframes onp{{0%,100%{{opacity:1}}50%{{opacity:.45}}}}
    </style></head>
    <body>
      <div class="strip">
        <span class="dot"></span><strong>Portal ativo</strong>
        <span class="up" id="uptime-display">00h 00m 00s</span>
      </div>
      <script>
        var serverStart = {server_start} * 1000;
        function tick() {{
            var dist = new Date().getTime() - serverStart;
            var h = Math.floor(dist / 3600000);
            var m = Math.floor((dist % 3600000) / 60000);
            var s = Math.floor((dist % 60000) / 1000);
            var fmt = (n) => (n < 10 ? "0" + n : "" + n);
            var el = document.getElementById("uptime-display");
            if (el) el.textContent = fmt(h) + "h " + fmt(m) + "m " + fmt(s) + "s";
        }}
        tick();
        setInterval(tick, 1000);
      </script>
    </body></html>
    """, height=48)
    
    
    paginas_nav = {
        "registrar_chamado": {
            "label": "Registrar Chamado",
            "titulo": "Entrada de Chamados",
            "descricao": "Abertura guiada com prioridade, categoria e evidências."
        },
        "solicitar_compra": {
            "label": "Solicitar Compra TI",
            "titulo": "Compra de TI",
            "descricao": "Solicitações padronizadas para equipamentos e serviços."
        },
        "historico": {
            "label": "Meu Histórico",
            "titulo": "Histórico Operacional",
            "descricao": "Acompanhe chamados e requisições por status."
        },
        "gestor": {
            "label": "Painel do Gestor (TI)",
            "titulo": "Console Gerencial",
            "descricao": "Fila ativa, prioridades e decisões da equipe de TI."
        },
        "relatorios": {
            "label": "Relatórios & Exportação",
            "titulo": "Central de Dados",
            "descricao": "KPIs, backups e exportações para análise."
        },
    }
    lista_paginas = list(paginas_nav.keys())
    rotas_legadas = {dados["label"]: chave for chave, dados in paginas_nav.items()}

    APELIDOS = {
        "chamado":    "registrar_chamado",
        "compras":    "solicitar_compra",
        "historico":  "historico",
        "painel":     "gestor",
        "relatorios": "relatorios",
    }
    APELIDO_DE = {v: k for k, v in APELIDOS.items()}

    def _pagina_da_url():
        bruto = st.query_params.get("p") or st.query_params.get("page") or ""
        bruto = str(bruto).strip()
        if bruto in APELIDOS:        return APELIDOS[bruto]
        if bruto in lista_paginas:   return bruto
        if bruto in rotas_legadas:   return rotas_legadas[bruto]
        return "registrar_chamado"

    if "pagina_atual" not in st.session_state:
        st.session_state.pagina_atual = _pagina_da_url()
    elif st.session_state.pagina_atual not in lista_paginas:
        st.session_state.pagina_atual = rotas_legadas.get(st.session_state.pagina_atual, "registrar_chamado")

    def atualizar_url():
        st.query_params.clear()
        st.query_params["p"] = APELIDO_DE.get(st.session_state.pagina_atual,
                                              st.session_state.pagina_atual)

    pagina = st.radio(
        "NAVEGAÇÃO PRINCIPAL",
        lista_paginas,
        key="pagina_atual",
        on_change=atualizar_url,
        format_func=lambda chave: paginas_nav[chave]["label"],
        label_visibility="visible"
    )

    nav_ativa = paginas_nav[pagina]
    st.markdown(f"""
        <div class="nav-context">
            <div class="nav-context-signal" aria-hidden="true"><span></span><span></span><span></span></div>
            <div class="nav-context-label">Seção atual</div>
            <div class="nav-context-title">{nav_ativa["titulo"]}</div>
            <div class="nav-context-copy">{nav_ativa["descricao"]}</div>
        </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
        <div class="sidebar-footer">
            Desenvolvido por <a href="https://www.linkedin.com/in/hiago-cavalheiro-42887b387/" target="_blank" style="text-decoration: none;">Hiago Cavalheiro</a>
            <small>Analista TI Pleno • Service Operations</small>
        </div>
    """, unsafe_allow_html=True)

maybe_show_ui_skeleton()

CATEGORIAS_CHAMADO = {
    "Hardware (PC/Note)": [
        "computador", "pc", "notebook", "monitor", "tela",
        "hardware", "liga", "desliga", "travando", "lento", "memoria", "hd",
        "ssd", "fonte", "cabo", "gabinete", "placa", "bateria", "carregador"
    ],
    "Periféricos": [
        "periferico", "perifericos", "teclado", "mouse", "headset", "fone",
        "fones", "webcam", "camera", "microfone", "mic", "hub", "usb",
        "adaptador", "dock", "mousepad", "scanner de mao", "leitor"
    ],
    "Internet/Rede": [
        "internet", "rede", "wifi", "wi-fi", "conexao", "conectividade",
        "cabo de rede", "sem rede", "sem internet", "ping", "vpn",
        "roteador", "switch", "dns", "ip", "ethernet", "lento", "lenta",
        "oscilando", "caindo"
    ],
    "E-mail/Login": [
        "email", "e-mail", "outlook", "login", "senha", "usuario", "conta",
        "autenticacao", "bloqueado", "mfa", "codigo", "entrar",
        "credencial", "office", "microsoft", "gmail", "godaddy"
    ],
    "Impressora": [
        "impressora", "imprimir", "impressao", "scanner", "digitalizar",
        "toner", "papel", "fila", "spooler", "driver", "xerox"
    ],
    "Sistema/Software": [
        "sistema", "software", "programa", "aplicativo", "app", "erro",
        "bug", "tela branca", "sistema lento", "instalar", "atualizar",
        "planilha", "navegador", "chrome", "windows"
    ],
    "Telefonia": [
        "telefonia", "telefone", "ramal", "celular", "voip", "chip",
        "sim", "simcard", "ligacao", "chamada", "discar", "fixo",
        "operadora", "linha telefonica", "sinal"
    ],
    "Acesso/Permissão": [
        "permissao", "permissoes", "pasta", "pastas", "compartilhada",
        "compartilhamento", "diretorio", "autorizacao", "autorizar",
        "liberar acesso", "grupo", "active directory", "nivel de acesso",
        "permitir", "negado"
    ],
    "Outros": [
        "outros", "duvida", "solicitacao", "orientacao", "ajuda", "suporte"
    ]
}

TECNICOS_RESPONSAVEIS = [
    "Não atribuído",
    "Hiago Cavalheiro"
]

def normalizar_busca(texto):
    texto = str(texto or "").lower()
    texto = unicodedata.normalize("NFD", texto)
    return "".join(c for c in texto if unicodedata.category(c) != "Mn")

def texto_para_widget(valor, padrao=""):
    if pd.isna(valor):
        return padrao

    texto = str(valor).strip()
    if texto.lower() in ["aguardando", "nan", "none", "<na>"]:
        return padrao

    return texto

def detectar_tipo_chamado(descricao, categoria_padrao="Outros"):
    texto = normalizar_busca(descricao)
    palavras_texto = Counter(re.findall(r"[a-z0-9]+", texto))
    pontuacao = Counter()

    for categoria, palavras_chave in CATEGORIAS_CHAMADO.items():
        for termo in palavras_chave:
            termo_norm = normalizar_busca(termo)
            partes = re.findall(r"[a-z0-9]+", termo_norm)

            if len(partes) == 1:
                quantidade = palavras_texto.get(partes[0], 0)
                if quantidade:
                    pontuacao[categoria] += quantidade
            elif termo_norm in texto:
                pontuacao[categoria] += len(partes) + 1

    if pontuacao:
        categoria_vencedora, score = pontuacao.most_common(1)[0]
        if score > 0:
            return categoria_vencedora

    categoria_padrao = str(categoria_padrao or "").strip()
    _pad = re.sub(r"^[^\w(]+", "", str(categoria_padrao)).strip()
    return _pad if _pad in CATEGORIAS_CHAMADO else "Outros"

def montar_resposta_resolucao(row, tipo_chamado):
    descricao = str(row.get("Descrição", "")).strip()
    tipo_limpo = normalizar_busca(tipo_chamado)

    respostas_por_tipo = {
        "impressora": {
            "diagnostico": "Falha relacionada à impressora, fila de impressão, conexão ou comunicação com o equipamento.",
            "acoes": [
                "Verificada a conexão da impressora e o status do equipamento.",
                "Conferida a fila de impressão e removidos possíveis trabalhos travados.",
                "Reiniciado o serviço de impressão quando necessário.",
                "Validado o driver e enviada uma página de teste."
            ],
            "validacao": "Impressão de teste realizada ou orientada para confirmar o funcionamento."
        },
        "internet": {
            "diagnostico": "Instabilidade ou falha de conectividade na rede/internet.",
            "acoes": [
                "Verificado cabo, Wi-Fi ou ponto de rede utilizado pelo usuário.",
                "Conferida a conectividade local e o acesso à internet.",
                "Aplicada renovação de conexão/reinicialização do adaptador quando necessário.",
                "Orientado novo teste de acesso aos sistemas e sites utilizados."
            ],
            "validacao": "Conexão testada com navegação e acesso aos recursos de rede."
        },
        "rede": {
            "diagnostico": "Instabilidade ou falha de conectividade na rede/internet.",
            "acoes": [
                "Verificado cabo, Wi-Fi ou ponto de rede utilizado pelo usuário.",
                "Conferida a conectividade local e o acesso à internet.",
                "Aplicada renovação de conexão/reinicialização do adaptador quando necessário.",
                "Orientado novo teste de acesso aos sistemas e sites utilizados."
            ],
            "validacao": "Conexão testada com navegação e acesso aos recursos de rede."
        },
        "mail": {
            "diagnostico": "Falha relacionada a e-mail, login, senha ou autenticação.",
            "acoes": [
                "Conferido o usuário informado e o cenário de acesso.",
                "Orientada nova tentativa de login com credenciais corretas.",
                "Verificado bloqueio, senha expirada ou necessidade de redefinição.",
                "Validado acesso ao e-mail/sistema após o procedimento."
            ],
            "validacao": "Acesso validado com login ou envio/recebimento de mensagem de teste."
        },
        "login": {
            "diagnostico": "Falha relacionada a e-mail, login, senha ou autenticação.",
            "acoes": [
                "Conferido o usuário informado e o cenário de acesso.",
                "Orientada nova tentativa de login com credenciais corretas.",
                "Verificado bloqueio, senha expirada ou necessidade de redefinição.",
                "Validado acesso ao e-mail/sistema após o procedimento."
            ],
            "validacao": "Acesso validado com login ou envio/recebimento de mensagem de teste."
        },
        "hardware": {
            "diagnostico": "Falha ou instabilidade relacionada ao computador, notebook ou periférico.",
            "acoes": [
                "Realizada análise inicial do equipamento informado.",
                "Verificados cabos, energia, periféricos e comportamento do sistema.",
                "Aplicados testes básicos de funcionamento e reinicialização controlada.",
                "Orientado acompanhamento caso a falha volte a ocorrer."
            ],
            "validacao": "Equipamento validado com teste de uso após os procedimentos."
        },
        "software": {
            "diagnostico": "Falha relacionada a sistema, software, aplicativo ou configuração.",
            "acoes": [
                "Identificado o sistema ou programa afetado.",
                "Verificada mensagem de erro e comportamento informado pelo usuário.",
                "Aplicados ajustes, reinicialização ou orientação de uso conforme o caso.",
                "Realizado teste de abertura e funcionamento do sistema."
            ],
            "validacao": "Sistema testado novamente para confirmar funcionamento."
        },
        "sistema": {
            "diagnostico": "Falha relacionada a sistema, software, aplicativo ou configuração.",
            "acoes": [
                "Identificado o sistema ou programa afetado.",
                "Verificada mensagem de erro e comportamento informado pelo usuário.",
                "Aplicados ajustes, reinicialização ou orientação de uso conforme o caso.",
                "Realizado teste de abertura e funcionamento do sistema."
            ],
            "validacao": "Sistema testado novamente para confirmar funcionamento."
        },
        "outros": {
            "diagnostico": "Solicitação classificada como atendimento geral de suporte de TI.",
            "acoes": [
                "Analisada a descrição enviada pelo usuário.",
                "Realizada triagem inicial para identificar o impacto e o melhor encaminhamento.",
                "Aplicada orientação técnica conforme o cenário informado.",
                "Registrada a tratativa para acompanhamento do atendimento."
            ],
            "validacao": "Solicitação validada com base no retorno do usuário ou na orientação aplicada."
        }
    }

    modelo = respostas_por_tipo["software"]
    for chave, resposta in respostas_por_tipo.items():
        if chave in tipo_limpo:
            modelo = resposta
            break

    acoes = "\n".join(f"- {acao}" for acao in modelo["acoes"])
    return f"""Chamado analisado e classificado como: {tipo_chamado}.

Descrição informada pelo usuário:
{descricao}

Diagnóstico:
{modelo["diagnostico"]}

Ações realizadas/recomendadas:
{acoes}

Validação:
{modelo["validacao"]}

Orientação ao usuário:
Caso o problema volte a ocorrer, registrar um novo chamado ou reabrir este atendimento informando o horário, mensagem de erro exibida e um print da tela, se possível.

Status sugerido: Concluído após validação com o usuário."""

@st.dialog("Despachar Ticket")
def modal_responder_chamado(idx, row):
    categoria_atual = extrair_categoria(row)
    tipo_detectado = detectar_tipo_chamado(row.get("Descrição", ""), categoria_atual)
    status_opcoes = ["Aberto", "Em Atendimento", "Concluído"]
    status_atual = str(row.get("Status", "Aberto")).strip()
    if status_atual not in status_opcoes:
        status_atual = "Aberto"

    st.markdown(f"### Chamado {primeiro_valor(row, 'Protocolo', padrao=f'#{formatar_id(idx)}')}")
    st.write(f"**Solicitante:** {row['Usuário']} | **Setor:** {row['Setor']}")
    st.write(f"**Data de Abertura:** {row['Data']}")
    st.write(f"**Tipo detectado:** {tipo_detectado} | **Impacto:** {primeiro_valor(row, 'Impacto')}")
    st.markdown(f"> **Descrição:** {row['Descrição']}")
    renderizar_anexo(row.get('Anexo'), key=f"dl_anexo_ch_{idx}")
    st.markdown("---")

    ccc1, ccc2 = st.columns(2)
    with ccc1:
        novo_status = st.selectbox(
            "Alterar Status do Ticket:",
            status_opcoes,
            index=status_opcoes.index(status_atual)
        )
    with ccc2:
        categorias_opcoes = list(CATEGORIAS_CHAMADO.keys())
        tipo_chamado = st.selectbox(
            "Tipo do chamado:",
            categorias_opcoes,
            index=categorias_opcoes.index(tipo_detectado),
            help="Detectado automaticamente pela descrição. Ajuste manualmente se necessário."
        )

    tecnico_atual = primeiro_valor(row, "Técnico Responsável", padrao="Não atribuído").strip()
    tecnicos_opcoes = TECNICOS_RESPONSAVEIS.copy()
    if tecnico_atual and tecnico_atual not in tecnicos_opcoes:
        tecnicos_opcoes.append(tecnico_atual)

    tecnico_resp = st.selectbox(
        "Técnico responsável:",
        tecnicos_opcoes,
        index=tecnicos_opcoes.index(tecnico_atual) if tecnico_atual in tecnicos_opcoes else 0
    )

    nota_key = f"nota_tecnica_chamado_{idx}"
    resposta_atual = texto_para_widget(row.get("Resposta do Técnico", ""))
    nota_sessao = st.session_state.get(nota_key, "")
    if (
        nota_key not in st.session_state
        or not isinstance(nota_sessao, str)
        or nota_sessao.strip().lower() in ["aguardando", "nan", "none", "<na>"]
    ):
        st.session_state[nota_key] = resposta_atual

    resposta_resolucao = montar_resposta_resolucao(row, tipo_chamado)
    if st.button("Gerar resposta técnica automática", use_container_width=True, key=f"btn_gerar_resposta_{idx}"):
        st.session_state[nota_key] = resposta_resolucao

    resp_tec = st.text_area(
        "Notas Técnicas de Resolução:",
        key=nota_key,
        height=220,
        placeholder="Insira o parecer técnico ou gere uma resposta técnica automática para este chamado..."
    )

    st.markdown("**Imagem da resposta** *(opcional)*")
    _anexo_resp_atual = str(row.get("Anexo Resposta", "") or "").strip()
    if _anexo_resp_atual and _anexo_resp_atual.lower() not in ("nenhum", "nan", "none"):
        st.caption("Já anexada nesta resposta:")
        renderizar_anexo(_anexo_resp_atual, key=f"dl_anexo_resp_{idx}")
    img_resposta = st.file_uploader(
        "", type=["png", "jpg", "jpeg"], key=f"img_resp_{idx}",
        label_visibility="collapsed",
    )
    zona_colar()

    prio_colors = {
        "Baixa": "var(--c-acc)",
        "Média": "var(--c-run)",
        "Alta": "var(--c-crit-l)",
        "Crítica": "var(--c-crit)",
        "Urgente": "var(--c-crit)"
    }
    st.markdown(
        f"<div style='text-align:right; padding-top:10px;'>Prioridade: "
        f"<strong style='color:{prio_colors.get(row['Prioridade'], '#fff')}'>{row['Prioridade']}</strong></div>",
        unsafe_allow_html=True
    )

    if st.button("Gravar Alterações e Fechar", use_container_width=True, type="primary"):
        df_base.at[idx, "Status"] = novo_status
        df_base.at[idx, "Categoria"] = tipo_chamado
        df_base.at[idx, "Resposta do Técnico"] = resp_tec.strip() or "Aguardando"
        df_base.at[idx, "Técnico Responsável"] = tecnico_resp
        _nome_img = _anexo_resp_atual
        if img_resposta is not None:
            _ts_img = datetime.now().strftime("%Y%m%d_%H%M%S")
            _nome_img = f"{_ts_img}_resposta_{img_resposta.name}"
            os.makedirs(ANEXO_DIR, exist_ok=True)
            with open(os.path.join(ANEXO_DIR, _nome_img), "wb") as _f_img:
                _f_img.write(img_resposta.getbuffer())
        df_base.at[idx, "Anexo Resposta"] = _nome_img
        df_base.at[idx, "Atualizado em"] = datetime.now().strftime("%d/%m/%Y %H:%M")
        salvar_dataframe_seguro(df_base, DATABASE, COLUNAS)
        st.session_state.pop(nota_key, None)
        st.success("Chamado atualizado com sucesso!")
        st.rerun(scope="app")


if pagina == "registrar_chamado":

    @st.fragment
    def _render_registrar():
        global df_base

        if "ch_step" not in st.session_state: st.session_state.ch_step = 1
        wz = st.session_state.ch_step

        CATS = [
            ("💻", "Hardware",       "PC, notebook, monitor"),
            ("🖱️", "Periféricos",    "Teclado, mouse, headset"),
            ("🌐", "Internet/Rede",  "Sem acesso, lentidão, VPN"),
            ("📧", "E-mail/Login",   "Godaddy, Gmail, Outlook, Login"),
            ("🖨️", "Impressora",     "Impressão, scanner, toner"),
            ("📑", "Sistema/SW",     "Erros em software, travamentos"),
            ("📱", "Telefonia",      "Ramal, celular, VoIP"),
            ("🔐", "Acesso/Permissão","Pastas, senhas, permissões"),
            ("❓", "Outro",          "Não se encaixa acima"),
        ]
        PRIOS = [
            ("Baixa",   "var(--c-acc)", "var(--ta120)",  72, "🔵", "Posso aguardar"),
            ("Média",   "var(--c-run)", "rgba(var(--rgb-run),.12)",  24, "🟡", "Afeta meu trabalho"),
            ("Alta",    "var(--c-run)", "rgba(var(--rgb-run),.12)",   4, "🟠", "Estou parado"),
            ("Crítica", "var(--c-crit)", "rgba(var(--rgb-crit),.12)",    2, "🔴", "Empresa parada"),
        ]

        st.markdown("""
<style>
    label,.stRadio label{outline:none!important;box-shadow: none!important}
    .wz{display:flex;align-items:center;padding:6px 0 20px;user-select:none}
    .wz-node{display:flex;flex-direction:column;align-items:center;gap:5px;flex:0 0 auto}
    .wz-circle{width:30px;height:30px;border-radius:50%;display:flex;align-items:center;
      justify-content:center;font-size:12px;font-weight:800;transition:all .25s}
    .wz-c-done{background:var(--f-ok);color:var(--t-n11);box-shadow: none}
    .wz-c-act{background:var(--n2);border:2px solid var(--c-acc);color:var(--c-acc);box-shadow: none}
    .wz-c-idle{background:var(--n1);border:2px solid var(--n3);color:var(--t-n3)}
    .wz-lbl{font-size:10px;font-weight:700;letter-spacing:.02em;white-space:nowrap}
    .wz-lbl-done{color:var(--c-ok)}.wz-lbl-act{color:var(--t-n10)}.wz-lbl-idle{color:var(--t-n3)}
    .wz-line{flex:1;height:2px;margin:0 8px;margin-bottom:15px;border-radius:2px}
    .wz-line-done{background:var(--f-ok)}.wz-line-idle{background:var(--n3)}
    .hz{background:linear-gradient(135deg,var(--sup-a),var(--sup-b));border:1px solid var(--n3);
      border-radius:12px;padding:18px 22px;position:relative;overflow:hidden;margin-bottom:14px;
      box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.05), 0 10px 22px var(--shc-2), 0 2px 5px var(--shc-2)}
    .hz::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;
      background:linear-gradient(90deg,var(--f-acc),var(--f-run-l));box-shadow: none}
    .hz-eye{font-size:9px;font-weight:800;letter-spacing:.15em;text-transform:uppercase;color:var(--t-n5);margin-bottom:5px}
    .hz-title{font-size:1.1rem;font-weight:800;color:var(--t-n11);margin-bottom:3px}
    .hz-desc{font-size:.78rem;color:var(--t-n6);line-height:1.5}
    .hz-crumb{font-size:10px;color:var(--t-n5);margin-bottom:5px}
    .hz-crumb b{color:var(--t-n7)}
    .prio-chip{display:inline-flex;align-items:center;gap:7px;padding:6px 14px;
      border-radius:8px;font-size:11px;font-weight:700;margin:6px 0 2px}
    .prio-dot{width:7px;height:7px;border-radius:50%;flex-shrink:0}
    .char-c{font-size:9px;text-align:right;margin-top:3px}
    .ok-card{background:var(--n1);border:1px solid var(--n3);border-left:3px solid var(--f-ok);
      border-radius:6px;padding:14px 16px;margin:2px 0 14px;}
    .ok-topo{display:flex;align-items:center;gap:9px;flex-wrap:wrap}
    .ok-titulo{font-size:14px;font-weight:600;color:var(--t-n11)}
    .ok-proto{font-family:var(--mono);font-size:11.5px;color:var(--t-n9);background:var(--n2);
      border:1px solid var(--n3);border-radius:4px;padding:2px 8px;letter-spacing:.04em;user-select:all}
    .ok-sub{font-size:12.5px;color:var(--t-n7);margin-top:7px;line-height:1.5}
    .ok-fluxo{display:flex;align-items:center;gap:9px;margin-top:12px;flex-wrap:wrap}
    .ok-etapa{display:flex;align-items:center;gap:6px;font-size:11px;color:var(--t-n7)}
    .ok-etapa::before{content:"";width:6px;height:6px;border-radius:50%;background:var(--n4)}
    .ok-etapa.feita{color:var(--c-ok)}
    .ok-etapa.feita::before{background:var(--f-ok)}
    .ok-sep{width:16px;height:1px;background:var(--n4)}
    /* CAT CARD — tudo no botão, sem markdown separado */
    div[data-testid="stButton"] button.cat-btn {
      background:var(--n2); border:1.5px solid var(--n3); border-radius:11px;
      padding:18px 8px 14px; width:100%; height:auto; min-height:110px;
      display:flex; flex-direction:column; align-items:center; justify-content:center;
      gap:5px; cursor:pointer; transition:all .2s; white-space:normal;
    }

    /* ══ INPUT FIELD OVERHAUL — seletores data-baseweb diretos ══ */
    .stTextInput label p,
    .stSelectbox label p,
    .stNumberInput label p,
    .stTextArea label p {
      font-family:var(--font)!important;
      font-size:.82rem!important;font-weight:700!important;
      letter-spacing:.01em!important;text-transform:none!important;
      color:var(--c-acc-xl)!important;
    }
    div[data-baseweb="base-input"],div[data-baseweb="input"]{
      background:linear-gradient(160deg,var(--n1) 0%,var(--n1) 100%)!important;
      border:1px solid var(--ta220)!important;
      border-radius:10px!important;
      box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.04), inset 0 -1px 0 var(--shc-2), 0 4px 18px var(--shc-2)!important;
      transition:border-color .22s ease,box-shadow .22s ease,transform .18s ease!important;
    }
    div[data-baseweb="base-input"]:focus-within,div[data-baseweb="input"]:focus-within{
      border-color:var(--ta720)!important;
      box-shadow: 0 0 0 3px var(--ta130), inset 0 1px 0 var(--ta100)!important;
      transform:translateY(-2px)!important;
    }
    div[data-baseweb="base-input"]:hover:not(:focus-within),div[data-baseweb="input"]:hover:not(:focus-within){
      border-color:var(--ta360)!important;transform:translateY(-1px)!important;
    }
    div[data-baseweb="base-input"] input,div[data-baseweb="input"] input{
      background:transparent!important;color:var(--c-acc-l)!important;
      font-family:'Archivo',sans-serif!important;font-size:.90rem!important;
      font-weight:500!important;caret-color:var(--c-acc)!important;padding:0 14px!important;
    }
    div[data-baseweb="base-input"] input::placeholder,div[data-baseweb="input"] input::placeholder{
      color:rgba(var(--rgb-t-n9),.68)!important;font-style:normal!important;font-size:.85rem!important;
    }
    div[data-baseweb="select"]>div:first-child{
      background:linear-gradient(160deg,var(--n1) 0%,var(--n1) 100%)!important;
      border:1px solid var(--ta220)!important;border-radius:10px!important;
      box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.04), 0 4px 18px var(--shc-2)!important;
      transition:border-color .22s ease,box-shadow .22s ease,transform .18s ease!important;
    }
    div[data-baseweb="select"]>div:first-child:hover{
      border-color:var(--ta380)!important;transform:translateY(-1px)!important;
    }
    div[data-baseweb="select"]:focus-within>div:first-child{
      border-color:var(--ta720)!important;
      box-shadow: 0 0 0 3px var(--ta130), inset 0 1px 0 var(--ta100)!important;
      transform:translateY(-2px)!important;
    }
    div[data-baseweb="select"] svg{
      fill:var(--ta520)!important;transition:transform .22s ease,fill .2s ease!important;
    }
    div[data-baseweb="select"]:focus-within svg{fill:var(--f-acc)!important;transform:rotate(180deg)!important;}
    ul[data-baseweb="menu"],div[data-baseweb="menu"]{
      background:linear-gradient(160deg,var(--n1),var(--n1))!important;
      border:1px solid var(--ta260)!important;border-radius:12px!important;
      box-shadow: 0 0 0 1px var(--ta060), 0 24px 60px var(--shc-1);
      padding:5px!important;
    }
    li[role="option"]{
      border-radius:8px!important;color:var(--t-n9)!important;
      font-family:'Archivo',sans-serif!important;font-size:.87rem!important;font-weight:500!important;
      transition:background .14s ease,color .14s ease,padding-left .14s ease!important;padding:9px 14px!important;
    }
    li[role="option"]:hover{
      background:linear-gradient(90deg,var(--ta160),var(--ta060))!important;
      color:var(--t-n11)!important;padding-left:18px!important;
    }
    li[role="option"][aria-selected="true"]{
      color:var(--c-acc-l)!important;font-weight:700!important;border-left:2px solid var(--ta750)!important;
      background:linear-gradient(90deg,var(--ta180),var(--ta060))!important;
    }
    </style>
    """, unsafe_allow_html=True)

        st.markdown("""
<style>
    /* ── LIVE FEED PANEL ── */
    .lf-wrap {
        position: relative;
        margin-bottom: 18px;
        border: 1px solid var(--ta140);
        border-radius: 12px;
        background: linear-gradient(145deg, rgba(var(--rgb-sup2-a),.96), rgba(var(--rgb-sup2-b),.92));
        overflow: hidden;
        box-shadow: 0 0 0 1px var(--ta040), 0 14px 36px var(--shc-2);
    }
    /* top scanner line */
    .lf-wrap::before {
        content:"";
        position:absolute; top:0; left:0; right:0; height:1px;
        background: linear-gradient(90deg, transparent, var(--ta550), rgba(var(--rgb-run-l),.35), transparent);
        animation: none;
    }
    @keyframes lfScanline {
        0%,100%{opacity:.55} 50%{opacity:1}
    }

    .lf-head {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 10px 16px 9px;
        border-bottom: 1px solid var(--ta080);
        background: rgba(var(--rgb-n0),.55);
        
    }
    .lf-head-left  { display:flex; align-items:center; gap:10px; }
    .lf-head-right { display:flex; align-items:center; gap:8px; }

    .lf-live-dot {
        position: relative;
        width: 7px; height: 7px; border-radius: 50%;
        background: var(--f-ok);
        box-shadow: none;
        flex-shrink: 0;
    }
    .lf-live-dot::after {
        content:"";
        position:absolute; inset:-4px; border-radius:50%;
        background: rgba(var(--rgb-ok),.25);
        animation: none;
    }
    @keyframes lfPulse {
        0%{transform:scale(.6);opacity:1} 100%{transform:scale(2.2);opacity:0}
    }

    .lf-label {
        font-family:'IBM Plex Mono',monospace;
        font-size:.60rem; font-weight:800; letter-spacing:.18em;
        text-transform:uppercase; color:rgba(var(--rgb-t-n9),.76);
    }
    .lf-count-badge {
        display:inline-flex; align-items:center;
        padding:2px 9px; border-radius:999px;
        background:var(--ta080);
        border:1px solid var(--ta200);
        font-family:'IBM Plex Mono',monospace;
        font-size:.62rem; font-weight:800; color:var(--c-acc);
        letter-spacing:.06em;
    }
    .lf-live-tag {
        font-family:'IBM Plex Mono',monospace;
        font-size:.55rem; font-weight:800; letter-spacing:.16em;
        color:rgba(var(--rgb-c-ok),.65); text-transform:uppercase;
    }

    /* ── ROWS ── */
    .lf-body {
        max-height: 236px;
        overflow-y: auto;
        overflow-x: hidden;
        scrollbar-width: thin;
        scrollbar-color: var(--ta250) transparent;
    }
    .lf-body::-webkit-scrollbar { width:3px; }
    .lf-body::-webkit-scrollbar-thumb {
        background:var(--ta280);
        border-radius:99px;
    }

    .lf-row {
        display: grid;
        grid-template-columns: 4px 112px 1fr 80px 74px 76px;
        align-items: center;
        gap: 0 14px;
        padding: 8px 16px 8px 12px;
        border-bottom: 1px solid rgba(var(--rgb-n9),.06);
        transition: background .15s ease;
        min-height: 44px;
    }
    .lf-row:last-child { border-bottom: none; }
    .lf-row:hover { background: var(--ta040); }

    /* priority stripe */
    .lf-stripe { border-radius: 2px; height: 100%; min-height: 28px; width: 4px; }
    .lf-stripe-critica { background:var(--f-crit-l); box-shadow: none; }
    .lf-stripe-alta    { background:var(--f-run); box-shadow: none; }
    .lf-stripe-media   { background:var(--f-run); box-shadow: none; }
    .lf-stripe-baixa   { background:var(--f-acc); box-shadow: none; }

    .lf-user {
        display:flex; flex-direction:column; gap:1px;
        overflow:hidden; min-width:0;
    }
    .lf-user-name {
        font-size:.74rem; font-weight:700; color:var(--t-n11);
        white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
    }
    .lf-protocol {
        font-family:'IBM Plex Mono',monospace;
        font-size:.58rem; font-weight:700; letter-spacing:.03em;
        color:rgba(var(--rgb-c-acc),.65); white-space:nowrap; overflow:hidden;
        text-overflow:ellipsis;
    }
    .lf-desc {
        font-size:.76rem; font-weight:600; color:var(--t-n10);
        white-space:normal; overflow-wrap:anywhere; line-height:1.45;
    }
    .lf-meta {
        font-size:.68rem; color:rgba(var(--rgb-t-n8),.79);
        white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
    }
    .lf-time {
        font-family:'IBM Plex Mono',monospace;
        font-size:.64rem; color:rgba(var(--rgb-t-n9),.72);
        text-align:right; white-space:nowrap;
    }
    .lf-status {
        display:inline-flex; align-items:center; gap:4px;
        padding:3px 8px; border-radius:999px;
        font-size:.60rem; font-weight:800; letter-spacing:.04em;
        white-space:nowrap;
    }
    .lf-status-aberto {
        background:rgba(var(--rgb-n9),.10); color:var(--t-n9);
        border:1px solid rgba(var(--rgb-n9),.20);
    }
    .lf-status-em {
        background:rgba(var(--rgb-run-l),.10); color:var(--c-run-l);
        border:1px solid rgba(var(--rgb-run-l),.25);
    }

    /* ── EMPTY STATE ── */
    .lf-empty {
        display:flex; align-items:center; justify-content:center; gap:10px;
        padding:18px;
        color:rgba(var(--rgb-t-n8),.64);
        font-family:'IBM Plex Mono',monospace;
        font-size:.68rem; letter-spacing:.06em;
    }
    .lf-empty::before {
        content:""; width:6px; height:6px; border-radius:50%;
        background:rgba(var(--rgb-n8),.30);
        box-shadow: none;
    }

    /* ── FOOTER BAR ── */
    .lf-foot {
        display:flex; align-items:center; justify-content:space-between;
        padding:6px 16px;
        border-top:1px solid var(--ta060);
        background:rgba(var(--rgb-n0),.35);
    }
    .lf-foot-txt {
        font-family:'IBM Plex Mono',monospace;
        font-size:.55rem; color:rgba(var(--rgb-c-acc),.30);
        letter-spacing:.08em; text-transform:uppercase;
    }
    .lf-critical-alert {
        display:inline-flex; align-items:center; gap:5px;
        font-family:'IBM Plex Mono',monospace;
        font-size:.56rem; font-weight:800; color:var(--c-crit-l);
        letter-spacing:.08em; animation: none;
    }
    @keyframes lfAlertBlink {
        0%,100%{opacity:.9} 50%{opacity:.4}
    }

    /* ══ FILA DE TRABALHO DA TI ══════════════════════════════════════
       Lista acionável: cada chamado é um cartão com o próprio botão de
       despacho, em vez de um bloco só de leitura + seletor separado. */
    .tif-bar {
        display:flex; align-items:center; gap:10px; flex-wrap:wrap;
        padding:9px 14px; margin:2px 0 10px;
        border:1px solid var(--ta160); border-radius:9px;
        background:linear-gradient(145deg, rgba(var(--rgb-sup2-a),.96), rgba(var(--rgb-sup2-b),.90));
    }
    .tif-bar-tit {
        font-family:'IBM Plex Mono',monospace;
        font-size:.60rem; font-weight:800; letter-spacing:.16em;
        text-transform:uppercase; color:rgba(var(--rgb-t-n9),.83);
    }
    .tif-badge {
        display:inline-flex; align-items:center; padding:2px 9px; border-radius:999px;
        font-family:'IBM Plex Mono',monospace; font-size:.62rem; font-weight:800;
        background:var(--ta080); border:1px solid var(--ta200);
        color:var(--c-acc);
    }
    .tif-badge-crit {
        background:rgba(var(--rgb-crit),.10); border-color:rgba(var(--rgb-crit),.28);
        color:var(--c-crit-l);
    }
    .tif-sessao {
        margin-left:auto; display:inline-flex; align-items:center; gap:6px;
        font-family:'IBM Plex Mono',monospace; font-size:.55rem; font-weight:800;
        letter-spacing:.10em; text-transform:uppercase;
        color:rgba(var(--rgb-c-ok),.75);
    }

    /* ── cartão de um chamado ── */
    .tif-row {
        display:grid;
        grid-template-columns: 3px 132px 1fr 86px 58px 86px;
        align-items:center; gap:0 13px;
        padding:0 13px; height:60px; box-sizing:border-box; overflow:hidden;
        border:1px solid var(--n3); border-radius:8px;
        background:var(--n1);
        transition: border-color .15s ease, background .15s ease;
    }
    .tif-row:hover { border-color:var(--ta380); background:var(--n2); }
    .tif-row-crit { border-left:3px solid var(--f-crit-l); }

    .tif-stripe { border-radius:2px; height:26px; width:3px; }
    .tif-p-critica { background:var(--f-crit-l); }
    .tif-p-alta    { background:var(--f-acc); }
    .tif-p-media   { background:var(--f-run); }
    .tif-p-baixa   { background:rgba(var(--rgb-n8),.35); }

    .tif-quem  { display:flex; flex-direction:column; gap:1px; min-width:0; overflow:hidden; }
    .tif-nome  {
        font-size:.76rem; font-weight:700; color:var(--t-n11);
        white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
    }
    .tif-proto {
        font-family:'IBM Plex Mono',monospace; font-size:.58rem; font-weight:700;
        letter-spacing:.03em; color:rgba(var(--rgb-c-acc),.70);
        white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
    }
    .tif-desc {
        font-size:.77rem; font-weight:600; color:var(--t-n10);
        line-height:1.42; overflow-wrap:anywhere;
        display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;
    }
    .tif-sub  { font-size:.67rem; color:rgba(var(--rgb-t-n8),.82); margin-top:1px; }
    .tif-prio {
        display:inline-flex; align-items:center; justify-content:center;
        padding:2px 8px; border-radius:5px; white-space:nowrap;
        font-size:.62rem; font-weight:800; letter-spacing:.02em;
    }
    .tif-prio-critica { background:rgba(var(--rgb-crit),.12); color:var(--c-crit-l); }
    .tif-prio-alta    { background:var(--ta140);  color:var(--c-acc); }
    .tif-prio-media   { background:rgba(var(--rgb-run),.12);  color:var(--c-run-l); }
    .tif-prio-baixa   { background:rgba(var(--rgb-n8),.10);   color:var(--t-n8); }

    .tif-tempo {
        font-family:'IBM Plex Mono',monospace; font-size:.68rem; font-weight:700;
        color:rgba(var(--rgb-t-n9),.77); text-align:right; white-space:nowrap;
    }
    .tif-tempo-velho { color:var(--c-crit-l); }
    .tif-st {
        display:inline-flex; align-items:center; gap:5px;
        padding:3px 8px; border-radius:999px; white-space:nowrap;
        font-size:.59rem; font-weight:800; letter-spacing:.04em;
    }
    .tif-st-aberto { background:rgba(var(--rgb-n9),.10); color:var(--t-n9);
                     border:1px solid rgba(var(--rgb-n9),.20); }
    .tif-st-em     { background:rgba(var(--rgb-run-l),.10); color:var(--c-run-l);
                     border:1px solid rgba(var(--rgb-run-l),.25); }

    .tif-nada {
        padding:18px; text-align:center; border:1px dashed var(--n3); border-radius:8px;
        font-size:.76rem; color:rgba(var(--rgb-t-n8),.79);
    }

    /* botões e filtros dentro da área da TI */
    /* botão com a mesma altura do cartão, para a lista ficar alinhada */
    .st-key-tif_area [data-testid="stButton"] button { height:60px; font-size:.74rem; font-weight:700; }
    .st-key-tif_area [data-testid="stVerticalBlock"] { gap:.4rem; }
    .st-key-tif_area [data-testid="stHorizontalBlock"] { align-items:stretch; }
    </style>
    """, unsafe_allow_html=True)

        _df_ativos = carregar_dados_cacheados(DATABASE, COLUNAS)
        _df_ativos = _df_ativos[_df_ativos["Status"].isin(["Aberto", "Em Atendimento"])].copy()

        _PRIO_PESO = {"Crítica": 5, "Urgente": 5, "Alta": 4, "Média": 3, "Baixa": 1}
        if not _df_ativos.empty:
            _df_ativos["_pw"] = _df_ativos["Prioridade"].map(_PRIO_PESO).fillna(0)
            _df_ativos["_dt"] = pd.to_datetime(_df_ativos["Data"], format="%d/%m/%Y %H:%M", errors="coerce")
            _df_ativos = _df_ativos.sort_values(["_pw", "_dt"], ascending=[False, True])

        _total_ativos = len(_df_ativos)
        _criticos     = int((_df_ativos["Prioridade"].isin(["Crítica", "Urgente", "Alta"])).sum()) if not _df_ativos.empty else 0

        if not sessao_ti_ativa():
            with st.expander("Acesso da equipe de TI", expanded=False):
                st.caption(
                    "Área restrita: exibe a fila de chamados e permite despachar. "
                    "Use o mesmo token do Painel do Gestor."
                )
                with st.form("form_acesso_ti_feed", clear_on_submit=True):
                    _tk = st.text_input(
                        "Token de acesso", type="password",
                        placeholder="token de acesso da TI",
                        label_visibility="collapsed", key="pwd_ti_feed",
                    )
                    if st.form_submit_button("Entrar", use_container_width=True):
                        if abrir_sessao_ti(_tk):
                            st.rerun(scope="app")
                        else:
                            st.error("Token inválido.")
        else:
            def _tif_prio_cls(prio):
                p = str(prio).lower()
                if "cr" in p or "urg" in p: return "critica"
                if "alt" in p:              return "alta"
                if "m\u00e9d" in p or "med" in p: return "media"
                return "baixa"

            def _tif_idade(data_txt):
                """(rótulo, minutos) desde a abertura."""
                try:
                    dt = datetime.strptime(str(data_txt), "%d/%m/%Y %H:%M")
                    m = int((datetime.now() - dt).total_seconds() / 60)
                except Exception:
                    return "—", 0
                if m < 60:   return f"{m}m", m
                if m < 1440: return f"{m//60}h", m
                return f"{m//1440}d", m

            _c_tit, _c_sair = st.columns([5, 1], vertical_alignment="center")
            with _c_tit:
                _crit_badge = (
                    f'<span class="tif-badge tif-badge-crit">{_criticos} priorit&aacute;rio'
                    f'{"s" if _criticos > 1 else ""}</span>' if _criticos else ""
                )
                st.markdown(
                    '<div class="tif-bar">'
                    '<span class="lf-live-dot"></span>'
                    '<span class="tif-bar-tit">Fila de trabalho &middot; TI</span>'
                    f'<span class="tif-badge">{_total_ativos} ativo{"s" if _total_ativos != 1 else ""}</span>'
                    f'{_crit_badge}'
                    '<span class="tif-sessao">&#9679; sess&atilde;o da TI ativa</span>'
                    '</div>',
                    unsafe_allow_html=True,
                )
            with _c_sair:
                if st.button("Sair da TI", key="btn_sair_ti_feed",
                             use_container_width=True,
                             help="Encerra a sessão da TI neste navegador"):
                    st.session_state["gestor_autenticado"] = False
                    st.session_state["gestor_auth_ts"] = 0.0
                    st.rerun(scope="app")

            with st.container(key="tif_area"):

                @st.fragment
                def _fila_ti():
                    _f1, _f2, _f3 = st.columns([5, 3, 2], vertical_alignment="bottom")
                    with _f1:
                        _busca = st.text_input(
                            "Buscar", key="tif_busca", label_visibility="collapsed",
                            placeholder="Buscar por protocolo, pessoa, setor ou descrição…",
                        )
                    with _f2:
                        _fstatus = st.segmented_control(
                            "Status", ["Todos", "Aberto", "Em atendimento"],
                            default="Todos", key="tif_status",
                            label_visibility="collapsed",
                        ) or "Todos"
                    with _f3:
                        _so_crit = st.toggle("Só prioritários", key="tif_crit")

                    _df = _df_ativos
                    if _fstatus == "Aberto":
                        _df = _df[_df["Status"] == "Aberto"]
                    elif _fstatus == "Em atendimento":
                        _df = _df[_df["Status"] == "Em Atendimento"]
                    if _so_crit:
                        _df = _df[_df["Prioridade"].isin(["Crítica", "Urgente", "Alta"])]
                    if _busca:
                        _t = normalizar_busca(_busca)
                        _cols = ["Protocolo", "Usuário", "Setor", "Descrição", "Categoria"]
                        _mask = False
                        for _c in _cols:
                            if _c in _df.columns:
                                _m = _df[_c].astype(str).map(lambda v: _t in normalizar_busca(v))
                                _mask = _m if _mask is False else (_mask | _m)
                        if _mask is not False:
                            _df = _df[_mask]

                    _visiveis = st.session_state.get("tif_qtd", 8)
                    _n = len(_df)

                    if _n == 0:
                        _msg = ("Nenhum chamado ativo no momento." if _total_ativos == 0
                                else "Nenhum chamado corresponde ao filtro.")
                        st.markdown(f'<div class="tif-nada">{_msg}</div>',
                                    unsafe_allow_html=True)
                        return

                    for _i, _r in _df.head(_visiveis).iterrows():
                        _pcls   = _tif_prio_cls(_r.get("Prioridade", ""))
                        _proto  = html.escape(str(_r.get("Protocolo", "—"))[-16:])
                        _nome   = html.escape(str(_r.get("Usuário", "—")).strip()[:24] or "—")
                        _cat    = html.escape(str(_r.get("Categoria", "—")))
                        _setor  = html.escape(str(_r.get("Setor", "—")))
                        _desc   = html.escape(str(_r.get("Descrição", "")).strip() or "—")
                        _prio   = html.escape(str(_r.get("Prioridade", "—")))
                        _status = str(_r.get("Status", "—"))
                        _lbl, _min = _tif_idade(_r.get("Data", ""))
                        _velho  = " tif-tempo-velho" if _min >= 1440 else ""
                        _stcls  = "tif-st-em" if _status == "Em Atendimento" else "tif-st-aberto"
                        _stlbl  = "EM ATEND." if _status == "Em Atendimento" else "ABERTO"
                        _crit   = " tif-row-crit" if _pcls == "critica" else ""

                        _ci, _cb = st.columns([9, 1.5], vertical_alignment="center")
                        with _ci:
                            st.markdown(
                                f'<div class="tif-row{_crit}">'
                                f'<div class="tif-stripe tif-p-{_pcls}"></div>'
                                f'<div class="tif-quem"><div class="tif-nome">{_nome}</div>'
                                f'<div class="tif-proto">{_proto}</div></div>'
                                f'<div><div class="tif-desc">{_desc}</div>'
                                f'<div class="tif-sub">{_cat} &middot; {_setor}</div></div>'
                                f'<div><span class="tif-prio tif-prio-{_pcls}">{_prio}</span></div>'
                                f'<div class="tif-tempo{_velho}" title="tempo desde a abertura">{_lbl}</div>'
                                f'<div><span class="tif-st {_stcls}">{_stlbl}</span></div>'
                                f'</div>',
                                unsafe_allow_html=True,
                            )
                        with _cb:
                            if st.button(
                                "Despachar", key=f"tif_desp_{_i}",
                                use_container_width=True,
                                type="primary" if _pcls == "critica" else "secondary",
                                help=f"Responder e concluir {_proto}",
                            ):
                                modal_responder_chamado(_i, _df.loc[_i])

                    if _n > _visiveis:
                        if st.button(f"Mostrar mais ({_n - _visiveis} restantes)",
                                     key="tif_mais", use_container_width=True):
                            st.session_state["tif_qtd"] = _visiveis + 8
                            st.rerun(scope="fragment")
                    elif _visiveis > 8:
                        if st.button("Recolher lista", key="tif_menos",
                                     use_container_width=True):
                            st.session_state["tif_qtd"] = 8
                            st.rerun(scope="fragment")

                _fila_ti()

        if "ultimo_chamado_registrado" in st.session_state:
            proto_ok = st.session_state.pop("ultimo_chamado_registrado")
            st.markdown(f"""
        <div class="ok-card">
          <div class="ok-topo">
            {_ico("ok", 17, "var(--c-ok)", 1.7)}
            <span class="ok-titulo">Chamado aberto</span>
            <span class="ok-proto">{proto_ok}</span>
          </div>
          <div class="ok-sub">A equipe de TI já foi notificada. Acompanhe em "Meu Histórico".</div>
          <div class="ok-fluxo">
            <span class="ok-etapa feita">Na fila</span><span class="ok-sep"></span>
            <span class="ok-etapa">Triagem</span><span class="ok-sep"></span>
            <span class="ok-etapa">Técnico</span>
          </div>
        </div>
        """, unsafe_allow_html=True)

        def _c(n): return "wz-c-done" if wz>n else ("wz-c-act" if wz==n else "wz-c-idle")
        def _l(n): return "wz-lbl-done" if wz>n else ("wz-lbl-act" if wz==n else "wz-lbl-idle")
        def _i(n): return "✓" if wz>n else str(n)
        def _ln(n): return "wz-line-done" if wz>n else "wz-line-idle"
        st.markdown(f"""
    <div class="wz">
      <div class="wz-node"><div class="wz-circle {_c(1)}">{_i(1)}</div><div class="wz-lbl {_l(1)}">Identificação</div></div>
      <div class="wz-line {_ln(1)}"></div>
      <div class="wz-node"><div class="wz-circle {_c(2)}">{_i(2)}</div><div class="wz-lbl {_l(2)}">Categoria</div></div>
      <div class="wz-line {_ln(2)}"></div>
      <div class="wz-node"><div class="wz-circle {_c(3)}">{_i(3)}</div><div class="wz-lbl {_l(3)}">Detalhes</div></div>
    </div>
    """, unsafe_allow_html=True)

        if wz == 1:
            col_form, col_ctx = st.columns([1.5, 1], gap="large")
            with col_form:
                with st.container(border=True, key="cartao_01"):
                    st.markdown("""<div class="hz hz-embutido">
          <div class="hz-eye">Suporte de TI</div>
          <div class="hz-title">Abertura de Chamado</div>
          <div class="hz-desc">Comece pela sua identificação — leva menos de 1 minuto.</div>
        </div>""", unsafe_allow_html=True)
                    nome_tmp = st.text_input("Nome completo",
                        value=st.session_state.get("ch_nome",""), placeholder="Ex: Hiago Cavalheiro")
                    SETORES = ["GERENCIA","SDR","TI","FINANCEIRO","OPERACIONAL","VENDAS","PÓS-VENDAS","RH","SUPERVISÃO","MARKETING","OUTRO"]
                    idx_s = SETORES.index(st.session_state.get("ch_setor","OPERACIONAL")) if st.session_state.get("ch_setor","OPERACIONAL") in SETORES else 4
                    setor_tmp = st.selectbox("Setor", SETORES, index=idx_s)
                if st.button("Continuar →", use_container_width=True, type="primary", key="ch_n1"):
                    if nome_tmp.strip():
                        st.session_state.ch_nome = nome_tmp.strip()
                        st.session_state.ch_setor = setor_tmp
                        st.session_state.ch_step = 2
                        st.rerun(scope="fragment")
                    else:
                        st.warning("Informe seu nome para continuar.")
            with col_ctx:
                st.markdown(f"""<div class="ctx-panel">
          <div class="ctx-h">Como funciona</div>
          <div class="ctx-step"><span class="ctx-n">1</span><div><strong>Identifique-se</strong><span>Seu nome e setor para localizarmos você.</span></div></div>
          <div class="ctx-step"><span class="ctx-n">2</span><div><strong>Categoria</strong><span>Escolha o tipo de problema em cards.</span></div></div>
          <div class="ctx-step"><span class="ctx-n">3</span><div><strong>Detalhes &amp; urgência</strong><span>Descreva o problema e defina a prioridade.</span></div></div>
          <div class="ctx-divider"></div>
          <div class="ctx-tip">{_ico("info", 15, extra="margin-right:7px")}Quanto mais detalhes na descrição, mais rápido o time resolve.</div>
        </div>""", unsafe_allow_html=True)

        elif wz == 2:
            nome_s2 = st.session_state.get("ch_nome","")
            st.markdown(f"""<div class="hz">
          <div class="hz-eye">Suporte de TI · Categoria</div>
          <div class="hz-title">Olá, {nome_s2}! Qual o tipo de problema?</div>
          <div class="hz-desc">Clique no card que melhor descreve o incidente.</div>
        </div>""", unsafe_allow_html=True)


            cat_sel = st.session_state.get("ch_cat_idx", None)

            def _sel_cat_ch(_idx, _emoji_cat, _label):
                st.session_state.ch_cat_idx = _idx
                st.session_state.ch_cat = f"{_emoji_cat} {_label}"

            for row in [CATS[i:i+3] for i in range(0, len(CATS), 3)]:
                cols = st.columns(3, gap="small")
                for col, (ico, label, hint) in zip(cols, row):
                    i_abs = CATS.index((ico, label, hint))
                    with col:
                        is_sel = cat_sel == i_abs
                        bc  = "var(--c-acc)" if is_sel else "#2f2f2b"
                        bg  = ("linear-gradient(160deg,var(--n4) 0%,var(--n2) 100%)" if is_sel
                               else "linear-gradient(160deg,var(--sup-a) 0%,var(--sup-b) 100%)")
                        _sh3d = ("inset 0 1px 0 rgba(var(--rgb-lift),.07),"
                                 "0 1px 2px var(--shc-4),"
                                 "0 8px 18px var(--shc-3)")
                        glw = (f"box-shadow:{_sh3d},0 0 22px var(--ta300);transform:translateY(-2px);"
                               if is_sel else f"box-shadow:{_sh3d};")
                        ico_col   = "var(--f-acc-l)" if is_sel else "var(--t-n9)"
                        lbl_col   = "var(--c-acc-l)" if is_sel else "var(--t-n9)"
                        hint_col  = "#ffcc4d99" if is_sel else "var(--t-n7)"
                        chk = f'<div style="position:absolute;top:8px;right:9px;width:16px;height:16px;border-radius:50%;background:var(--f-acc);color:var(--t-n11);font-size:9px;font-weight:900;display:flex;align-items:center;justify-content:center;line-height:1">✓</div>' if is_sel else ""
                        st.markdown(f"""
<div class="cat3d" style="background:{bg};border:1.5px solid {bc};border-radius:12px;
padding:20px 10px 8px;text-align:center;position:relative;{glw}
transition:transform .18s ease,box-shadow .18s ease;margin-bottom:4px">
{chk}
<div style="margin-bottom:9px;line-height:1;display:flex;justify-content:center">{_ico(_ICONE_CATEGORIA.get(label, "duvida"), 26, ico_col, 1.5)}</div>
<div style="font-size:12px;font-weight:800;color:{lbl_col};margin-bottom:4px;letter-spacing:.02em">{label}</div>
<div style="font-size:10px;color:{hint_col};line-height:1.4">{hint}</div>
</div>""", unsafe_allow_html=True)
                        st.button(
                            "Selecionado" if is_sel else "Selecionar",
                            key=f"cat_{i_abs}", use_container_width=True,
                            type="primary" if is_sel else "secondary",
                            on_click=_sel_cat_ch, args=(i_abs, ico, label),
                        )

            st.write("")
            cols_nav = st.columns([1, 3, 1])
            with cols_nav[0]:
                if st.button("← Voltar", key="ch_b2"): st.session_state.ch_step = 1; st.rerun(scope="fragment")
            if st.button("Continuar →", use_container_width=True, type="primary", key="ch_n2", disabled=(cat_sel is None)):
                st.session_state.ch_step = 3; st.rerun(scope="fragment")
            if cat_sel is None:
                st.caption("Selecione uma categoria para continuar.")

        elif wz == 3:
            cat_ch   = st.session_state.get("ch_cat","—")
            cat_ch_txt = re.sub(r"^[^\w(]+", "", str(cat_ch)).strip() or cat_ch
            nome_ch  = st.session_state.get("ch_nome","—")
            setor_ch = st.session_state.get("ch_setor","—")
            st.markdown(f"""<div class="hz">
          <div class="hz-crumb"><b>{nome_ch}</b> &nbsp;·&nbsp; {setor_ch} &nbsp;·&nbsp; {cat_ch_txt}</div>
          <div class="hz-title">Quase lá! Defina urgência e descreva o problema.</div>
        </div>""", unsafe_allow_html=True)
            with st.container(border=True, key="cartao_02"):
                st.markdown("**Urgência**")
                prio_labels = [p[0] for p in PRIOS]
                prio_prev   = st.session_state.get("ch_prio","Média")
                if prio_prev not in prio_labels: prio_prev = "Média"
                prio_sel = st.radio("Urgência", [f"{p[4]} {p[0]}" for p in PRIOS],
                    index=prio_labels.index(prio_prev), horizontal=True, label_visibility="collapsed", key="ch_pr")
                prio_name = prio_sel.split(" ",1)[1]
                st.session_state.ch_prio = prio_name
                pd_ = next(p for p in PRIOS if p[0]==prio_name)
                st.markdown(
                    f'<div class="prio-chip" style="background:{pd_[2]};border:1px solid {pd_[1]}44">'
                    f'<div class="prio-dot" style="background:{pd_[1]}"></div>'
                    f'<span style="color:{pd_[1]}">Prioridade: <strong>{prio_name}</strong></span></div>',
                    unsafe_allow_html=True)
                st.markdown("---")
                desc = st.text_area("Descrição do problema",
                    value=st.session_state.get("ch_desc",""),
                    placeholder="Descreva o que acontece, quando começou e qualquer mensagem de erro.\nEx: 'O Outlook não abre desde as 8h — erro 0x8004010F'",
                    height=120, key="ch_da")
                nc = len(desc)
                st.markdown(
                    f'<div class="char-c" style="color:{"var(--c-ok)" if nc>=30 else "var(--c-run)" if nc>0 else "var(--n5)"}">'
                    f'{nc} caracteres {"✓" if nc>=30 else "— recomendado 30+"}</div>',
                    unsafe_allow_html=True)


                st.markdown("---")
                st.markdown(f'<div class="lbl-anexo">{_ico("clipe", 15)}<span><b>Evidências</b> <i>(opcional)</i> — prints agilizam o diagnóstico</span></div>', unsafe_allow_html=True)
                arquivo_anexado = st.file_uploader("", type=["png","jpg","jpeg","pdf"],
                    label_visibility="collapsed", key="ch_f")
                zona_colar()
                st.caption("Até 200 MB por arquivo · PNG, JPG ou PDF")
            st.markdown(
                f'<div style="background:var(--n1);border:1px solid var(--n3);border-radius:9px;'
                f'padding:10px 15px;margin:8px 0 14px;font-size:11px;color:var(--t-n6)">'
                f'<strong style="color:var(--t-n10)">{nome_ch}</strong> &nbsp;·&nbsp; {setor_ch}'
                f' &nbsp;·&nbsp; {cat_ch_txt} &nbsp;·&nbsp; '
                f'<strong style="color:{pd_[1]}">{prio_name}</strong></div>',
                unsafe_allow_html=True)
            cols3 = st.columns([1, 2, 1])
            with cols3[0]:
                if st.button("← Voltar", key="ch_b3"):
                    st.session_state.ch_desc = desc
                    st.session_state.ch_step = 2; st.rerun(scope="fragment")
            enviar = st.button("Enviar Chamado →", use_container_width=True, type="primary", key="ch_env")
            if enviar:
                if not desc.strip() or len(desc.strip()) < 5:
                    st.warning("Descreva o problema antes de enviar.")
                else:
                    with st.status("Abrindo chamado...", expanded=True) as status:
                        nome_arquivo_salvo = "Nenhum"
                        if arquivo_anexado:
                            status.write("Enviando anexo…")
                            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                            nome_arquivo_salvo = f"{ts}_{arquivo_anexado.name}"
                            with open(os.path.join(ANEXO_DIR, nome_arquivo_salvo), "wb") as f:
                                f.write(arquivo_anexado.getbuffer())
                        status.write("Registrando o chamado…")
                        protocolo = gerar_protocolo("CH")
                        novo_chamado = {
                            "Data": datetime.now().strftime("%d/%m/%Y %H:%M"),
                            "Protocolo": protocolo, "Usuário": nome_ch, "Setor": setor_ch,
                            "Categoria": cat_ch, "Impacto": "", "Prioridade": prio_name,
                            "Descrição": f"[{cat_ch}] {desc.strip()}", "Status": "Aberto",
                            "Resposta do Técnico": "Aguardando", "Técnico Responsável": "Não atribuído",
                            "Atualizado em": datetime.now().strftime("%d/%m/%Y %H:%M"),
                            "Anexo": nome_arquivo_salvo
                        }
                        df_base = pd.concat([df_base, pd.DataFrame([novo_chamado])], ignore_index=True)
                        salvar_dataframe_seguro(df_base, DATABASE, COLUNAS)
                        status.write("Notificando a equipe de TI…")
                        _corpo = f"Protocolo: {protocolo}\nUsuário: {nome_ch}\nSetor: {setor_ch}\nPrioridade: {prio_name}"
                        notificar_windows("🚨 NOVO CHAMADO TI!", _corpo)
                        notificar_whatsapp("🚨 NOVO CHAMADO TI!", _corpo, url=_url_pagina("gestor"))
                        status.update(label=f"✅ {protocolo} aberto!", state="complete", expanded=False)
                    for k in ["ch_step","ch_nome","ch_setor","ch_cat","ch_cat_idx","ch_prio","ch_desc"]:
                        st.session_state.pop(k, None)
                    st.session_state.ultimo_chamado_registrado = protocolo
                    st.rerun(scope="fragment")


    _render_registrar()
elif pagina == "solicitar_compra":

    @st.fragment
    def _render_solicitar():

        if "cp_step" not in st.session_state: st.session_state.cp_step = 1
        wz_c = st.session_state.cp_step

        CATS_C = [
            ("💻", "Hardware",     "Notebooks, PCs, monitores"),
            ("🖱️", "Periféricos",  "Teclados, mouses, headsets"),
            ("🔌", "Infra/Rede",   "Switches, cabos, roteadores"),
            ("📑", "Licença/SW",   "Office, antivírus, licenças"),
            ("🗳️", "Suprimentos",  "Toner, papel, baterias"),
            ("📱", "Mobilidade",   "Celulares, tablets, chips"),
            ("🔋", "Energia",      "No-break, estabilizador"),
            ("❓", "Outro",        "Outro item de TI"),
        ]
        ITEMS_C = {
            "Hardware":    ["Notebook corporativo","Desktop / PC","Monitor","Memória RAM","SSD / HD","Outro"],
            "Periféricos": ["Teclado USB","Mouse Sem Fio","Headset c/ Microfone","Webcam HD","Hub USB-C","Outro"],
            "Infra/Rede":  ["Switch","Roteador / Access Point","Cabo de Rede (Patch)","Rack / Patch Panel","Outro"],
            "Licença/SW":  ["Microsoft Office / 365","Antivírus","Adobe Creative Cloud","Outro"],
            "Suprimentos": ["Toner / Cartucho","Papel A4","Pilha / Bateria","Outro"],
            "Mobilidade":  ["Celular corporativo","Tablet","Chip / Linha de dados","Capa / Película","Outro"],
            "Energia":     ["No-break / UPS","Estabilizador","Bateria de No-break","Filtro de linha","Outro"],
            "Outro":       ["Outro item de TI"],
        }
        URG_DATA = {
            "Sem pressa": ("var(--c-acc)","var(--ta120)",120),
            "Moderada":   ("var(--c-acc)","var(--ta120)", 48),
            "Alta":       ("var(--c-run)","rgba(var(--rgb-run),.12)",  12),
            "Crítica":    ("var(--c-crit)","rgba(var(--rgb-crit),.12)",    4),
        }

        st.markdown("""
<style>
    label,.stRadio label{outline:none!important;box-shadow: none!important}
    .wz{display:flex;align-items:center;padding:6px 0 20px;user-select:none}
    .wz-node{display:flex;flex-direction:column;align-items:center;gap:5px;flex:0 0 auto}
    .wz-circle{width:30px;height:30px;border-radius:50%;display:flex;align-items:center;
      justify-content:center;font-size:12px;font-weight:800;transition:all .25s}
    .wz-c-done{background:var(--f-ok);color:var(--t-n11);box-shadow: none}
    .wz-c-act{background:var(--n1);border:2px solid var(--c-acc);color:var(--c-acc);box-shadow: none}
    .wz-c-idle{background:var(--n1);border:2px solid var(--n3);color:var(--t-n3)}
    .wz-lbl{font-size:10px;font-weight:700;letter-spacing:.02em;white-space:nowrap}
    .wz-lbl-done{color:var(--c-ok)}.wz-lbl-act{color:var(--t-n10)}.wz-lbl-idle{color:var(--t-n3)}
    .wz-line{flex:1;height:2px;margin:0 8px;margin-bottom:15px;border-radius:2px}
    .wz-line-done{background:var(--f-ok)}.wz-line-idle{background:var(--n3)}
    .hzp{background:linear-gradient(135deg,var(--sup2-a),var(--sup2-b));border:1px solid var(--n3);
      border-radius:12px;padding:18px 22px;position:relative;overflow:hidden;margin-bottom:14px;
      box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.05), 0 10px 22px var(--shc-2), 0 2px 5px var(--shc-2)}
    .hzp::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;
      background:linear-gradient(90deg,var(--f-acc),var(--f-acc-d));box-shadow: none}
    .hzp-eye{font-size:9px;font-weight:800;letter-spacing:.15em;text-transform:uppercase;color:var(--t-n3);margin-bottom:5px}
    .hzp-title{font-size:1.1rem;font-weight:800;color:var(--t-n11);margin-bottom:3px}
    .hzp-desc{font-size:.78rem;color:var(--t-n5);line-height:1.5}
    .hzp-crumb{font-size:10px;color:var(--t-n4);margin-bottom:5px}
    .hzp-crumb b{color:var(--t-n7)}
    .prio-chip2{display:inline-flex;align-items:center;gap:7px;padding:6px 14px;
      border-radius:8px;font-size:11px;font-weight:700;margin:6px 0 2px}
    .prio-dot2{width:7px;height:7px;border-radius:50%;flex-shrink:0}
        /* ══ INPUT FIELD OVERHAUL — violet theme ══ */
    .stTextInput label p,
    .stSelectbox label p,
    .stNumberInput label p,
    .stTextArea label p {
      font-family:var(--font)!important;
      font-size:.82rem!important;font-weight:700!important;
      letter-spacing:.01em!important;text-transform:none!important;
      color:var(--c-acc-xl)!important;
    }
    div[data-baseweb="base-input"],div[data-baseweb="input"]{
      background:linear-gradient(160deg,var(--sup2-a) 0%,var(--sup2-b) 100%)!important;
      border:1px solid var(--ta220)!important;
      border-radius:10px!important;
      box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.04), inset 0 -1px 0 var(--shc-2), 0 4px 18px var(--shc-2)!important;
      transition:border-color .22s ease,box-shadow .22s ease,transform .18s ease!important;
    }
    div[data-baseweb="base-input"]:focus-within,div[data-baseweb="input"]:focus-within{
      border-color:var(--ta720)!important;
      box-shadow: 0 0 0 3px var(--ta130), inset 0 1px 0 var(--ta100)!important;
      transform:translateY(-2px)!important;
    }
    div[data-baseweb="base-input"]:hover:not(:focus-within),div[data-baseweb="input"]:hover:not(:focus-within){
      border-color:var(--ta360)!important;transform:translateY(-1px)!important;
    }
    div[data-baseweb="base-input"] input,div[data-baseweb="input"] input{
      background:transparent!important;color:var(--t-n11)!important;
      font-family:'Archivo',sans-serif!important;font-size:.90rem!important;
      font-weight:500!important;caret-color:var(--c-acc)!important;padding:0 14px!important;
    }
    div[data-baseweb="base-input"] input::placeholder,div[data-baseweb="input"] input::placeholder{
      color:rgba(var(--rgb-t-n9),.68)!important;font-style:normal!important;font-size:.85rem!important;
    }
    div[data-baseweb="select"]>div:first-child{
      background:linear-gradient(160deg,var(--sup2-a) 0%,var(--sup2-b) 100%)!important;
      border:1px solid var(--ta220)!important;border-radius:10px!important;
      box-shadow: inset 0 1px 0 rgba(var(--rgb-lift),.04), 0 4px 18px var(--shc-2)!important;
      transition:border-color .22s ease,box-shadow .22s ease,transform .18s ease!important;
    }
    div[data-baseweb="select"]>div:first-child:hover{
      border-color:var(--ta380)!important;transform:translateY(-1px)!important;
    }
    div[data-baseweb="select"]:focus-within>div:first-child{
      border-color:var(--ta720)!important;
      box-shadow: 0 0 0 3px var(--ta130), inset 0 1px 0 var(--ta100)!important;
      transform:translateY(-2px)!important;
    }
    div[data-baseweb="select"] svg{
      fill:var(--ta520)!important;transition:transform .22s ease,fill .2s ease!important;
    }
    div[data-baseweb="select"]:focus-within svg{fill:var(--f-acc)!important;transform:rotate(180deg)!important;}
    ul[data-baseweb="menu"],div[data-baseweb="menu"]{
      background:linear-gradient(160deg,var(--sup2-a),var(--sup2-b))!important;
      border:1px solid var(--ta260)!important;border-radius:12px!important;
      box-shadow: 0 0 0 1px var(--ta060), 0 24px 60px var(--shc-1);
      padding:5px!important;
    }
    li[role="option"]{
      border-radius:8px!important;color:var(--c-acc-l)!important;
      font-family:'Archivo',sans-serif!important;font-size:.87rem!important;font-weight:500!important;
      transition:background .14s ease,color .14s ease,padding-left .14s ease!important;padding:9px 14px!important;
    }
    li[role="option"]:hover{
      background:linear-gradient(90deg,var(--ta160),var(--ta060))!important;
      color:var(--t-n11)!important;padding-left:18px!important;
    }
    li[role="option"][aria-selected="true"]{
      color:var(--c-acc-l)!important;font-weight:700!important;border-left:2px solid var(--ta750)!important;
      background:linear-gradient(90deg,var(--ta180),var(--ta060))!important;
    }
        /* O CTA desta página usava fundo âmbar com rótulo `--act-ink`, que no
           tema claro é BRANCO: âmbar + branco dá 1.1:1, ilegível. Sem esta
           sobrescrita ele herda o botão primário único do portal — preto no
           claro, âmbar no escuro, sempre com o texto contrastando. */
</style>
    """, unsafe_allow_html=True)

        if "ultima_compra_registrada" in st.session_state:
            proto_c = st.session_state.pop("ultima_compra_registrada")
            st.markdown(f"""<div class="ok-card">
          <div class="ok-topo">
            {_ico("ok", 17, "var(--c-ok)", 1.7)}
            <span class="ok-titulo">Solicitação registrada</span>
            <span class="ok-proto">{proto_c}</span>
          </div>
          <div class="ok-sub">Enviada para análise da TI. Acompanhe em "Meu Histórico".</div>
        </div>""", unsafe_allow_html=True)

        def _cc(n): return "wz-c-done" if wz_c>n else ("wz-c-act" if wz_c==n else "wz-c-idle")
        def _lc(n): return "wz-lbl-done" if wz_c>n else ("wz-lbl-act" if wz_c==n else "wz-lbl-idle")
        def _ic(n): return "✓" if wz_c>n else str(n)
        def _lnc(n): return "wz-line-done" if wz_c>n else "wz-line-idle"
        st.markdown(f"""
    <div class="wz">
      <div class="wz-node"><div class="wz-circle {_cc(1)}">{_ic(1)}</div><div class="wz-lbl {_lc(1)}">Solicitante</div></div>
      <div class="wz-line {_lnc(1)}"></div>
      <div class="wz-node"><div class="wz-circle {_cc(2)}">{_ic(2)}</div><div class="wz-lbl {_lc(2)}">O que precisa?</div></div>
      <div class="wz-line {_lnc(2)}"></div>
      <div class="wz-node"><div class="wz-circle {_cc(3)}">{_ic(3)}</div><div class="wz-lbl {_lc(3)}">Justificativa</div></div>
    </div>
    """, unsafe_allow_html=True)

        if wz_c == 1:
            col_form, col_ctx = st.columns([1.5, 1], gap="large")
            with col_form:
                st.markdown("""<div class="hzp">
          <div class="hzp-eye">IT Procurement · Requisição</div>
          <div class="hzp-title">Solicitação de Compras</div>
          <div class="hzp-desc">Identifique-se e indique a urgência do item.</div>
        </div>""", unsafe_allow_html=True)
                with st.container(border=True, key="cartao_03"):
                    nome_ct = st.text_input("Nome completo",
                        value=st.session_state.get("cp_nome",""), placeholder="Quem receberá o item?")
                    cF1, cF2 = st.columns(2)
                    with cF1:
                        SETORES_C = ["GERENCIA","SDR","TI","FINANCEIRO","OPERACIONAL","VENDAS","RH","DIRETORIA","SUPERVISÃO","MARKETING","OUTRO"]
                        idx_sc = SETORES_C.index(st.session_state.get("cp_setor","OPERACIONAL")) if st.session_state.get("cp_setor","OPERACIONAL") in SETORES_C else 4
                        setor_ct = st.selectbox("Setor", SETORES_C, index=idx_sc)
                    with cF2:
                        urg_t = st.selectbox("Urgência", list(URG_DATA.keys()),
                            index=list(URG_DATA.keys()).index(st.session_state.get("cp_urg","Moderada")) if st.session_state.get("cp_urg","Moderada") in URG_DATA else 1)
                    uc,ubg,uh = URG_DATA[urg_t]
                    st.markdown(
                        f'<div class="prio-chip2" style="background:{ubg};border:1px solid {uc}44">'
                        f'<div class="prio-dot2" style="background:{uc}"></div>'
                        f'<span style="color:{uc}">Urgência: <strong>{urg_t}</strong></span></div>',
                        unsafe_allow_html=True)
                if st.button("Continuar →", use_container_width=True, type="primary", key="cp_n1"):
                    if nome_ct.strip():
                        st.session_state.cp_nome  = nome_ct.strip()
                        st.session_state.cp_setor = setor_ct
                        st.session_state.cp_urg   = urg_t
                        st.session_state.cp_step  = 2
                        st.rerun(scope="fragment")
                    else:
                        st.warning("Informe seu nome para continuar.")
            with col_ctx:
                st.markdown(f"""<div class="ctx-panel">
          <div class="ctx-h">Como funciona</div>
          <div class="ctx-step"><span class="ctx-n">1</span><div><strong>Solicitante</strong><span>Quem vai receber o item e a urgência.</span></div></div>
          <div class="ctx-step"><span class="ctx-n">2</span><div><strong>O que precisa</strong><span>Categoria e item desejado.</span></div></div>
          <div class="ctx-step"><span class="ctx-n">3</span><div><strong>Justificativa</strong><span>Motivo da compra e link do fornecedor.</span></div></div>
          <div class="ctx-divider"></div>
          <div class="ctx-tip">{_ico("info", 15, extra="margin-right:7px")}Inclua o link do produto na justificativa para acelerar a aprovação.</div>
        </div>""", unsafe_allow_html=True)

        elif wz_c == 2:
            nome_c2 = st.session_state.get("cp_nome","")
            st.markdown(f"""<div class="hzp">
          <div class="hzp-title">O que você precisa, {nome_c2}?</div>
          <div class="hzp-desc">Clique na categoria do item solicitado.</div>
        </div>""", unsafe_allow_html=True)


            cat_sel_c = st.session_state.get("cp_cat_idx", None)

            def _sel_cat_cp(_idx, _label):
                st.session_state.cp_cat_idx = _idx
                st.session_state.cp_cat = _label

            for row in [CATS_C[i:i+3] for i in range(0, len(CATS_C), 3)]:
                cols_c = st.columns(3, gap="small")
                for col_c, (ico, label, hint) in zip(cols_c, row):
                    i_c = next(i for i,(a,b,_) in enumerate(CATS_C) if b==label)
                    with col_c:
                        is_s = cat_sel_c == i_c
                        bc = "var(--c-acc)" if is_s else "#2f2f2b"
                        bg = ("linear-gradient(160deg,var(--n4) 0%,var(--n2) 100%)" if is_s
                              else "linear-gradient(160deg,var(--sup-a) 0%,var(--sup-b) 100%)")
                        _sh3d = ("inset 0 1px 0 rgba(var(--rgb-lift),.07),"
                                 "0 1px 2px var(--shc-4),"
                                 "0 8px 18px var(--shc-3)")
                        glw = (f"box-shadow:{_sh3d},0 0 22px var(--ta300);transform:translateY(-2px);"
                               if is_s else f"box-shadow:{_sh3d};")
                        lbl_col  = "var(--f-acc-l)" if is_s else "var(--t-n9)"
                        hint_col = "#ffcc4daa" if is_s else "var(--t-n7)"
                        chk = f'<div style="position:absolute;top:8px;right:9px;width:16px;height:16px;border-radius:50%;background:var(--f-acc);color:var(--t-n11);font-size:9px;font-weight:900;display:flex;align-items:center;justify-content:center">✓</div>' if is_s else ""
                        st.markdown(f"""
<div class="cat3d" style="background:{bg};border:1.5px solid {bc};border-radius:12px;
padding:20px 10px 8px;text-align:center;position:relative;{glw}
transition:transform .18s ease,box-shadow .18s ease;margin-bottom:4px">
{chk}
<div style="font-size:2rem;margin-bottom:7px;line-height:1">{ico}</div>
<div style="font-size:12px;font-weight:800;color:{lbl_col};margin-bottom:4px">{label}</div>
<div style="font-size:10px;color:{hint_col};line-height:1.4">{hint}</div>
</div>""", unsafe_allow_html=True)
                        st.button("Selecionado" if is_s else "Selecionar",
                                  key=f"ccat_{i_c}", use_container_width=True,
                                  type="primary" if is_s else "secondary",
                                  on_click=_sel_cat_cp, args=(i_c, label))

            st.write("")
            if cat_sel_c is not None:
                with st.container(border=True, key="cartao_04"):
                    col_item, col_qtd = st.columns([3, 1])
                    with col_item:
                        label_sel = CATS_C[cat_sel_c][1]
                        item_s = st.selectbox("Item específico", ITEMS_C[label_sel], key="cp_is")
                    with col_qtd:
                        qtd_s = st.number_input("Qtd", min_value=1,
                            value=st.session_state.get("cp_qtd",1), key="cp_qs")
                    spec_s = st.text_input("Modelo / Especificação *(opcional)*",
                        value=st.session_state.get("cp_spec",""),
                        placeholder="Ex: Dell Latitude 3420, Logitech G203...", key="cp_ss")
            else:
                st.caption("Selecione uma categoria acima para ver os itens.")
                item_s, qtd_s, spec_s = None, 1, ""

            cols_cn2 = st.columns([1, 3, 1])
            with cols_cn2[0]:
                if st.button("← Voltar", key="cp_b2"): st.session_state.cp_step = 1; st.rerun(scope="fragment")
            if st.button("Continuar →", use_container_width=True, type="primary", key="cp_n2", disabled=(cat_sel_c is None)):
                st.session_state.cp_item = spec_s.strip() or item_s
                st.session_state.cp_qtd  = int(qtd_s)
                st.session_state.cp_spec = spec_s.strip()
                st.session_state.cp_step = 3
                st.rerun(scope="fragment")

        elif wz_c == 3:
            nome_c3  = st.session_state.get("cp_nome","—")
            setor_c3 = st.session_state.get("cp_setor","—")
            urg_c3   = st.session_state.get("cp_urg","Moderada")
            cat_c3   = st.session_state.get("cp_cat","—")
            item_c3  = st.session_state.get("cp_item","—")
            qtd_c3   = st.session_state.get("cp_qtd",1)
            uc3,ubg3,uh3 = URG_DATA.get(urg_c3, ("var(--c-acc)","var(--ta120)",48))
            st.markdown(f"""<div class="hzp">
          <div class="hzp-crumb"><b>{nome_c3}</b> · {setor_c3} · {item_c3} ×{qtd_c3} · <b style="color:{uc3}">{urg_c3}</b></div>
          <div class="hzp-title">Justifique a necessidade da compra.</div>
        </div>""", unsafe_allow_html=True)
            with st.container(border=True, key="cartao_05"):
                forn = st.text_input("Fornecedor / link *(opcional)*",
                    value=st.session_state.get("cp_forn",""),
                    placeholder="Cole um link ou nome do fornecedor", key="cp_fs")
                st.markdown("---")
                just = st.text_area("Justificativa *",
                    value=st.session_state.get("cp_just",""),
                    placeholder="Por que esse item é necessário? Qual o impacto se não for adquirido?",
                    height=120, key="cp_js")
                nj = len(just)
                st.markdown(
                    f'<div style="font-size:9px;text-align:right;margin-top:3px;'
                    f'color:{"var(--c-ok)" if nj>=20 else "var(--c-run)" if nj>0 else "var(--n5)"}">'
                    f'{nj} chars {"✓" if nj>=20 else "— mínimo 20"}</div>',
                    unsafe_allow_html=True)
            st.markdown(
                f'<div style="background:var(--n0);border:1px solid var(--n3);border-radius:9px;'
                f'padding:10px 15px;margin:8px 0 14px;font-size:11px;color:var(--t-n5)">'
                f'<strong style="color:var(--t-n10)">{nome_c3}</strong> · {setor_c3}'
                f' · {item_c3} ×{qtd_c3} · <strong style="color:{uc3}">{urg_c3}</strong></div>',
                unsafe_allow_html=True)
            cols_c3 = st.columns([1, 2, 1])
            with cols_c3[0]:
                if st.button("← Voltar", key="cp_b3"):
                    st.session_state.cp_just = just; st.session_state.cp_forn = forn
                    st.session_state.cp_step = 2; st.rerun(scope="fragment")
            enviar_c = st.button("Enviar Solicitação →", use_container_width=True, type="primary", key="cp_env")
            if enviar_c:
                if not just.strip():
                    st.warning("Adicione uma justificativa antes de enviar.")
                else:
                    with st.status("Registrando...", expanded=True) as sc:
                        sc.write("Gravando a solicitação…")
                        protocolo_compra = gerar_protocolo("CP")
                        just_final = just.strip()
                        if forn.strip(): just_final += f"\n\nFornecedor/link: {forn.strip()}"
                        df_c = carregar_dados(DATABASE_COMPRAS, COLUNAS_COMPRAS)
                        novo_reg = pd.DataFrame([{
                            "Data": datetime.now().strftime("%d/%m/%Y %H:%M"),
                            "Protocolo": protocolo_compra, "Usuário": nome_c3, "Setor": setor_c3,
                            "Categoria": cat_c3, "Item": item_c3, "Quantidade": qtd_c3,
                            "Justificativa": just_final, "Urgência": urg_c3, "Status": "Pendente",
                            "Observação TI": "Aguardando análise", "Responsável": "Não atribuído",
                            "Atualizado em": datetime.now().strftime("%d/%m/%Y %H:%M")
                        }])
                        salvar_dataframe_seguro(pd.concat([df_c, novo_reg], ignore_index=True), DATABASE_COMPRAS, COLUNAS_COMPRAS)
                        sc.write("Notificando a equipe de TI…")
                        _corpo_c = f"Protocolo: {protocolo_compra}\nSolicitante: {nome_c3}\nItem: {item_c3} (x{qtd_c3})\nUrgência: {urg_c3}"
                        notificar_windows("🛒 NOVA SOLICITAÇÃO DE COMPRA", _corpo_c)
                        notificar_whatsapp("🛒 NOVA SOLICITAÇÃO DE COMPRA", _corpo_c, url=_url_pagina("gestor"))
                        sc.update(label=f"✅ {protocolo_compra} registrada!", state="complete", expanded=False)
                    for k in ["cp_step","cp_nome","cp_setor","cp_urg","cp_cat","cp_cat_idx","cp_item","cp_qtd","cp_spec","cp_forn","cp_just"]:
                        st.session_state.pop(k, None)
                    st.session_state.ultima_compra_registrada = protocolo_compra
                    st.rerun(scope="fragment")


    _render_solicitar()
elif pagina == "historico":

    @st.fragment
    def _render_historico():


        st.markdown("""
<style>
        .hist-scope .queue-card{
            position:relative; padding:15px 18px 15px 22px; border-radius:14px;
            background:linear-gradient(145deg,rgba(var(--rgb-n1),.94),rgba(var(--rgb-n1),.94));
            border:1px solid var(--ta100); overflow:hidden;
            transition:transform .18s ease, border-color .2s ease, box-shadow .2s ease;
        }
        .hist-scope .queue-card::before{
            content:""; position:absolute; left:0; top:0; bottom:0; width:4px;
            background:var(--hs-accent,var(--f-acc));
            box-shadow: none;
        }
        .hist-scope .queue-card:hover{
            transform:translateY(-2px); border-color:var(--ta300);
            box-shadow: 0 14px 32px var(--shc-2);
        }
        .hist-scope .hs-open{--hs-accent:var(--c-crit-l);}
        .hist-scope .hs-prog{--hs-accent:var(--c-run);}
        .hist-scope .hs-done{--hs-accent:var(--c-ok);}
        .hist-scope .work-card-head{
            display:flex; justify-content:space-between; align-items:flex-start;
            gap:14px; flex-wrap:wrap;
        }
        .hist-scope .work-card-id{
            font-family:'IBM Plex Mono',monospace; font-size:.60rem; font-weight:700;
            letter-spacing:.14em; text-transform:uppercase; color:var(--c-acc); margin-bottom:4px;
        }
        .hist-scope .work-card-title{
            font-size:1.04rem; font-weight:800; color:var(--t-n11); line-height:1.18;
            letter-spacing:-.01em;
        }
        .hist-scope .work-card-subtitle{
            font-size:.73rem; color:var(--t-n8); margin-top:4px; font-weight:500;
        }
        .hist-scope .work-badge-row{
            display:flex; gap:6px; flex-wrap:wrap; align-items:center; justify-content:flex-end;
        }
        .hist-scope .work-card-desc{
            margin-top:12px; padding-top:11px; border-top:1px solid rgba(var(--rgb-n9),.09);
            font-size:.82rem; color:var(--t-n10); line-height:1.52;
        }
        /* ── Cabeçalho de perfil da pessoa ── */
        .hist-profile{
            display:flex; align-items:center; gap:16px; flex-wrap:wrap;
            margin:14px 0 18px; padding:18px 20px; border-radius:16px;
            background:linear-gradient(135deg,var(--ta100),rgba(var(--rgb-run-l),.08));
            border:1px solid var(--ta180);
        }
        .hist-avatar{
            width:56px; height:56px; border-radius:14px; flex:0 0 auto;
            display:flex; align-items:center; justify-content:center;
            font-size:1.3rem; font-weight:800; color:var(--t-n1);
            background:linear-gradient(135deg,var(--f-acc),var(--f-run));
            box-shadow: 0 8px 20px var(--ta300);
        }
        .hist-profile-info{flex:1 1 240px; min-width:0;}
        .hist-profile-eyebrow{font-family:'IBM Plex Mono',monospace; font-size:.58rem;
            font-weight:700; letter-spacing:.16em; text-transform:uppercase; color:var(--c-acc);}
        .hist-profile-name{font-size:1.45rem; font-weight:800; color:var(--t-n11);
            line-height:1.15; margin:2px 0 4px; letter-spacing:-.01em;}
        .hist-profile-sub{font-size:.84rem; color:var(--t-n10);}
        .hist-profile-stats{display:flex; gap:18px; flex:0 0 auto;}
        .hps{text-align:center;}
        .hps-v{font-size:1.5rem; font-weight:800; color:var(--t-n11); line-height:1;}
        .hps-l{font-size:.58rem; font-weight:700; letter-spacing:.10em;
            text-transform:uppercase; color:var(--t-n8); margin-top:4px;}
        /* ── Métricas ricas dentro do card ── */
        .hist-scope .hist-metrics{display:grid; gap:8px; margin-top:12px;
            grid-template-columns:repeat(auto-fit,minmax(118px,1fr));}
        .hist-scope .hm{display:flex; flex-direction:column; gap:2px; padding:8px 10px;
            border-radius:9px; background:rgba(var(--rgb-n0),.5); border:1px solid rgba(var(--rgb-n9),.09);}
        .hist-scope .hm-wide{grid-column:1/-1;}
        .hist-scope .hm-l{font-size:.55rem; font-weight:700; letter-spacing:.09em;
            text-transform:uppercase; color:var(--t-n7);}
        .hist-scope .hm-v{font-size:.82rem; font-weight:600; color:var(--t-n10);}
        /* ── Caixa de busca de colaborador (minimal) ── */
        .hist-search-head{display:flex; align-items:center; gap:8px; margin-bottom:7px;}
        .hist-search-ico{width:24px; height:24px; border-radius:7px; flex:0 0 auto;
            display:flex; align-items:center; justify-content:center; font-size:.78rem;
            background:linear-gradient(135deg,var(--ta200),rgba(var(--rgb-run),.18));
            border:1px solid var(--ta260);}
        .hist-search-ttl{font-size:.74rem; font-weight:600; color:var(--t-n11); line-height:1;
            letter-spacing:.02em;}
        .hist-search-hint{font-size:.72rem; color:var(--t-n8); margin-top:1px;}
        .hist-hit{display:inline-flex; align-items:center; gap:6px; font-size:.74rem;
            color:var(--t-n9); background:var(--ta080); border:1px solid var(--ta180);
            padding:4px 10px; border-radius:999px;}
        .hist-hit strong{color:var(--t-n11); font-weight:700;}

        /* ═══ CABEÇALHO + BUSCA — UM CARTÃO SÓ ═══════════════════════
           Antes eram três caixas empilhadas de mesmo peso: o título, a
           busca e um painel vazio de 200px com radar e grade. Nenhuma
           mandava na outra, e as regras "high-tech" da busca miravam
           `stVerticalBlockBorderWrapper`, que o Streamlit 1.56 não emite
           mais — ou seja, a caixa aparecia crua, sem nada do desenho.
           Agora o cabeçalho e a busca são um cartão, e o estado vazio é
           uma faixa discreta em vez de um painel. */
        .hist-cab{display:flex; flex-direction:column; gap:3px; margin-bottom:13px;}
        .hist-eyebrow{font-family:var(--mono); font-size:9.5px; font-weight:800;
            letter-spacing:.16em; text-transform:uppercase; color:var(--c-ok);}
        .hist-titulo{font-size:1.32rem; font-weight:800; color:var(--t-n11);
            letter-spacing:-.015em; line-height:1.15;}
        .hist-desc{font-size:.83rem; color:var(--t-n9); line-height:1.45;}
        .hist-rot{display:flex; align-items:center; gap:7px; margin-bottom:6px;
            font-size:.78rem; font-weight:700; color:var(--t-n10);}
        .hist-rot svg{opacity:.75;}

        /* faixa do estado vazio: uma linha, não um painel */
        .hist-dica{
            display:flex; align-items:center; gap:9px; flex-wrap:wrap;
            margin:2px 0 4px; padding:11px 14px; border-radius:9px;
            border:1px dashed var(--n4); background:var(--n1);
            font-size:.79rem; color:var(--t-n9);
        }
        .hist-dica-ico{display:inline-flex; color:var(--c-acc); flex:none;}
        .hist-dica-txt{flex:1 1 260px; min-width:0;}
        .hist-dica-pills{display:flex; gap:6px; flex-wrap:wrap;}

        /* campo de busca dentro do cartão */
        input[placeholder^="Ex.: jo"]{ font-size:.9rem !important; }
        input[placeholder^="Ex.: jo"]::placeholder{
            font-family:'IBM Plex Mono',monospace !important; letter-spacing:.04em !important;
            color:rgba(var(--rgb-t-n8),.70) !important;
        }

        /* ── Hero e cabeçalho de seção compactos (só nesta página) ── */
        .hero-compact{padding:15px 22px 16px !important; margin-bottom:14px !important;}
        .hero-compact .page-hero-eyebrow{margin-bottom:6px;}
        .hero-compact .page-hero-desc{font-size:.84rem; line-height:1.45;}
        .hist-scope-sec{margin:2px 0 12px !important; padding-bottom:9px !important;}
        </style>
    """, unsafe_allow_html=True)

        import unicodedata as _ud
        def _norm_nome(s):
            """Normaliza o nome: remove acentos, caixa e espaços extras."""
            s = str(s or "").strip().lower()
            s = _ud.normalize("NFD", s)
            s = "".join(c for c in s if _ud.category(c) != "Mn")
            return " ".join(s.split())

        with st.container(border=True, key="cartao_06"):
            st.markdown(f"""
            <div class="hist-cab">
                <div class="hist-eyebrow">Self-Service &middot; Rastreamento</div>
                <div class="hist-titulo">Meu Hist&oacute;rico</div>
                <div class="hist-desc">Acompanhe seus chamados e compras em um s&oacute; lugar.</div>
            </div>
            <div class="hist-rot">{_ico("busca", 14)} Buscar pelo seu nome</div>
            """, unsafe_allow_html=True)
            _q = st.text_input(
                "Buscar pelo seu nome",
                key="hist_query",
                placeholder="Ex.: joão  ·  silva  ·  maria souza",
                label_visibility="collapsed",
            )

        _qtokens = _norm_nome(_q).split()
        nome_busca = _q.strip()

        if not _qtokens:
            st.markdown(f"""
            <div class="hist-dica">
                <span class="hist-dica-ico">{_ico("info", 15)}</span>
                <span class="hist-dica-txt">Digite seu nome no campo acima para ver seus chamados e compras.</span>
                <span class="hist-dica-pills">
                    <span class="pill pill-red">🔴 Abertos</span>
                    <span class="pill pill-amber">🟡 Em atendimento</span>
                    <span class="pill pill-green">✅ Concluídos</span>
                    <span class="pill pill-gray">🛒 Compras</span>
                </span>
            </div>
        """, unsafe_allow_html=True)

        else:
            from datetime import datetime as _dt

            def _parse_data(v):
                try:
                    return _dt.strptime(str(v).strip(), "%d/%m/%Y %H:%M")
                except Exception:
                    return None

            def _humano(dt_ini):
                if not dt_ini:
                    return "—"
                seg = max(0, (_dt.now() - dt_ini).total_seconds())
                if seg < 3600:
                    return f"{int(seg // 60)} min"
                if seg < 86400:
                    return f"{int(seg // 3600)} h"
                return f"{int(seg // 86400)} d"

            def _bate(_h):
                return all(_t in _h for _t in _qtokens)
            df_user = df_base[df_base['Usuário'].map(_norm_nome).map(_bate)] if not df_base.empty else df_base.iloc[0:0]
            df_user_c = df_compras[df_compras['Usuário'].map(_norm_nome).map(_bate)] if not df_compras.empty else df_compras.iloc[0:0]

            _variantes = []
            if not df_user.empty:
                _variantes += list(df_user['Usuário'].astype(str))
            if not df_user_c.empty:
                _variantes += list(df_user_c['Usuário'].astype(str))
            _nome_disp = (max(set(_variantes), key=_variantes.count).strip() if _variantes else nome_busca.strip())

            tot_ch = len(df_user)
            tot_cp = len(df_user_c)
            _ab = int((df_user['Status'] == 'Aberto').sum()) if tot_ch else 0
            _at = int((df_user['Status'] == 'Em Atendimento').sum()) if tot_ch else 0
            _co = int((df_user['Status'] == 'Concluído').sum()) if tot_ch else 0
            _cp_abertas = int((~df_user_c['Status'].isin(['Aprovado', 'Recusado', 'Concluído'])).sum()) if tot_cp else 0
            _taxa = f"{int(_co / tot_ch * 100)}%" if tot_ch else "—"

            _andamento = _ab + _at
            if _andamento > 0:
                _resumo = f"Você tem <strong>{_andamento}</strong> chamado(s) em andamento"
                _resumo += f" e <strong>{_cp_abertas}</strong> compra(s) aguardando." if _cp_abertas else "."
            elif tot_ch or tot_cp:
                _resumo = "Tudo em dia — nenhum chamado em aberto."
            else:
                _resumo = "Nenhum registro encontrado para este nome."
            _iniciais = ("".join(p[0] for p in _nome_disp.split()[:2]).upper() or "?")

            try:
                _taxa_num = int(str(_taxa).replace('%', '').strip())
            except Exception:
                _taxa_num = None
            def _cu_block(slot, val, suffix=""):
                t = ("@property --hcSLOT{syntax:'<integer>';initial-value:0;inherits:false;}"
                     "@keyframes hcntSLOT{from{--hcSLOT:0}to{--hcSLOT:VAL}}"
                     ".hcnt-SLOT{counter-reset:c var(--hcSLOT);"
                     "animation:hcntSLOT 1.1s cubic-bezier(.2,.7,.3,1) forwards;}"
                     ".hcnt-SLOT::after{content:counter(c) SUFFIX;}")
                return t.replace("SLOT", slot).replace("VAL", str(val)).replace("SUFFIX", suffix)
            _cu = "<style>" + _cu_block("ch", tot_ch) + _cu_block("cp", tot_cp)
            if _taxa_num is not None:
                _cu += _cu_block("tx", _taxa_num, "'%'")
            _cu += "</style>"
            st.markdown(_cu, unsafe_allow_html=True)
            _v_ch = '<span class="hcnt-ch"></span>'
            _v_cp = '<span class="hcnt-cp"></span>'
            _v_tx = '<span class="hcnt-tx"></span>' if _taxa_num is not None else escapar_html(_taxa)

            st.markdown(f"""
            <div class="hist-profile">
                <div class="hist-avatar">{escapar_html(_iniciais)}</div>
                <div class="hist-profile-info">
                    <div class="hist-profile-eyebrow">Histórico pessoal</div>
                    <div class="hist-profile-name">{escapar_html(_nome_disp)}</div>
                    <div class="hist-profile-sub">{_resumo}</div>
                </div>
                <div class="hist-profile-stats">
                    <div class="hps"><div class="hps-v">{_v_ch}</div><div class="hps-l">Chamados</div></div>
                    <div class="hps"><div class="hps-v">{_v_cp}</div><div class="hps-l">Compras</div></div>
                    <div class="hps"><div class="hps-v" style="color:var(--c-ok)">{_v_tx}</div><div class="hps-l">Resolvidos</div></div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            if tot_ch == 0 and tot_cp == 0:
                st.markdown("""
                <div class="empty-state animate-fade-in" style="padding:32px 20px;margin-top:10px;">
                    <div style="margin-bottom:12px;display:flex;justify-content:center">{_ico("busca", 34)}</div>
                    <strong>Nenhum registro encontrado</strong>
                    <span>Não há chamados ou compras vinculados a este nome.</span>
                </div>
                """, unsafe_allow_html=True)
                return

            _status_filtro = None
            if tot_ch > 0:
                _status_opcoes = ["Todos", "🔴 Aberto", "🟡 Em Atendimento", "✅ Concluído"]
                _status_map = {"🔴 Aberto": "Aberto", "🟡 Em Atendimento": "Em Atendimento", "✅ Concluído": "Concluído"}
                st.markdown('<div class="hist-search-hint" style="margin:10px 0 2px;">Filtrar seus chamados por status:</div>', unsafe_allow_html=True)
                _status_sel = st.segmented_control(
                    "Filtrar chamados por status",
                    _status_opcoes,
                    default="Todos",
                    key="hist_status_filtro",
                    label_visibility="collapsed",
                )
                _status_filtro = _status_map.get(_status_sel or "Todos")

            df_user_view = (df_user[df_user['Status'].str.strip() == _status_filtro]
                             if _status_filtro else df_user)

            _eventos = []
            for _i, _r in df_user_view.iterrows():
                _eventos.append(("chamado", _parse_data(_r.get('Data')), _i, _r))
            if not _status_filtro:
                for _i, _r in df_user_c.iterrows():
                    _eventos.append(("compra", _parse_data(_r.get('Data')), _i, _r))
            _eventos.sort(key=lambda e: (e[1] is not None, e[1] or _dt.min), reverse=True)

            st.markdown('<div class="sec-label animate-fade-in" style="margin-top:8px;"><div class="sec-dot"></div>LINHA DO TEMPO</div>', unsafe_allow_html=True)
            if _status_filtro:
                st.caption(f"{len(_eventos)} chamado(s) com status \"{_status_filtro}\".")
            else:
                st.caption(f"{len(_eventos)} registro(s) — chamados e compras juntos, do mais recente ao mais antigo.")
            _distintos = {_norm_nome(v) for v in _variantes}
            if len(_distintos) > 1:
                st.caption(f"A busca encontrou {len(_distintos)} nomes diferentes. Digite mais do nome (ex.: nome + sobrenome) para refinar.")

            if _status_filtro and not _eventos:
                st.markdown(f"""
                <div class="empty-state animate-fade-in" style="padding:26px 20px;margin-top:6px;">
                    <div style="margin-bottom:10px;display:flex;justify-content:center">{_ico("busca", 28)}</div>
                    <strong>Nenhum chamado com status "{escapar_html(_status_filtro)}"</strong>
                    <span>Tente outro status no filtro acima.</span>
                </div>
                """, unsafe_allow_html=True)

            _linhas_hist, _detalhes_hist = [], []
            for _tipo, _data_dt, idx, row in _eventos:
                status_atual = str(row.get('Status', '—')).strip()
                protocolo = primeiro_valor(row, 'Protocolo', padrao=f'#{formatar_id(idx)}')
                data_str = str(row.get('Data', '—'))
                aberto_ha = _humano(_data_dt)
                no_status_ha = _humano(_parse_data(row.get('Atualizado em')))

                if _tipo == "chamado":
                    prio = str(row.get('Prioridade', '—'))
                    categoria = extrair_categoria(row)
                    desc_texto = str(row.get('Descrição', '')).strip()
                    desc_preview = desc_texto[:160] + '...' if len(desc_texto) > 160 else desc_texto
                    tecnico = str(row.get('Técnico Responsável', 'Não atribuído'))
                    resposta = str(row.get('Resposta do Técnico', '—'))
                    _tipo_emoji, _tipo_lbl = "🎫", "Chamado"
                    status_txt = ('🔴 Aberto' if status_atual == "Aberto"
                                  else '🟡 Em Atendimento' if status_atual == "Em Atendimento"
                                  else '✅ Concluído')
                    prox = ("Aguardando triagem/atribuição da TI." if status_atual == "Aberto"
                            else "Técnico trabalhando — aguarde retorno." if status_atual == "Em Atendimento"
                            else "Resolvido. Nenhuma ação necessária.")
                    titulo = desc_preview or "(sem descrição)"
                    detalhe = {
                        "tipo": "chamado", "tecnico": tecnico, "resposta": resposta,
                        "desc_texto": desc_texto, "anexo": str(row.get('Anexo', '')),
                        "anexo_resposta": str(row.get('Anexo Resposta', '')),
                    }
                else:
                    item = str(row.get('Item', '—')).strip()
                    desc_preview = item[:160] + '...' if len(item) > 160 else item
                    _tipo_emoji, _tipo_lbl = "🛒", "Compra"
                    status_txt = ('⏳ Pendente' if "Pendente" in status_atual
                                  else '🔍 Em Análise' if "Análise" in status_atual
                                  else '✅ Aprovado' if "Aprovado" in status_atual
                                  else '❌ Recusado' if "Recusado" in status_atual
                                  else '🟢 Concluído')
                    prox = ("Aguardando análise da TI." if ("Pendente" in status_atual or "Análise" in status_atual)
                            else "Aprovada — acompanhe a entrega." if "Aprovado" in status_atual
                            else "Recusada — verifique a justificativa." if "Recusado" in status_atual
                            else "Concluída.")
                    titulo = desc_preview or "(sem item)"
                    detalhe = {
                        "tipo": "compra", "quantidade": row.get("Quantidade", "—"),
                        "justificativa": row.get("Justificativa", "—"),
                    }

                _linhas_hist.append({
                    "Tipo": f"{_tipo_emoji} {_tipo_lbl}",
                    "Protocolo": protocolo,
                    "Título": titulo,
                    "Status": status_txt,
                    "Aberto há": aberto_ha,
                    "Data": data_str,
                })
                detalhe.update({"prox": prox, "no_status_ha": no_status_ha})
                _detalhes_hist.append(detalhe)

            if _linhas_hist:
                _evt_hist_pessoal = st.dataframe(
                    pd.DataFrame(_linhas_hist),
                    use_container_width=True,
                    hide_index=True,
                    selection_mode="single-row",
                    on_select="rerun",
                    key="grid_historico_pessoal",
                    column_config={"Título": st.column_config.TextColumn(width="large")},
                )
                _sel_hist = (_evt_hist_pessoal.selection.rows
                             if _evt_hist_pessoal and _evt_hist_pessoal.selection else [])
                if _sel_hist:
                    _d = _detalhes_hist[_sel_hist[0]]
                    with st.container(border=True, key="cartao_07"):
                        st.markdown(f'<div class="meta-chip"><strong>No status há:</strong> &nbsp;{escapar_html(_d["no_status_ha"])}</div>', unsafe_allow_html=True)
                        st.write("")
                        st.markdown(f'<div class="meta-chip" style="height:auto;align-items:flex-start;padding:10px 12px;"><div><strong>Próxima ação:</strong><br><span style="color:var(--t-n9);font-size:.78rem;">{escapar_html(_d["prox"])}</span></div></div>', unsafe_allow_html=True)
                        if _d["tipo"] == "chamado":
                            _resp_txt = (html.escape(_d["resposta"]) if _d["resposta"] and _d["resposta"] not in ['—', 'nan', 'Aguardando']
                                         else 'Aguardando resposta do técnico.')
                            st.write("")
                            st.markdown(f'<div class="meta-chip"><strong>Técnico:</strong> &nbsp;{html.escape(_d["tecnico"])}</div>', unsafe_allow_html=True)
                            st.write("")
                            st.markdown(f'<div class="meta-chip" style="height:auto;align-items:flex-start;padding:10px 12px;"><div><strong>Resposta:</strong><br><span style="color:var(--t-n9);font-size:.78rem;">{_resp_txt}</span></div></div>', unsafe_allow_html=True)
                            if _d["desc_texto"]:
                                st.write("")
                                st.markdown(f'<div class="ux-note" style="margin-top:0;"><strong>Descrição completa:</strong> {html.escape(_d["desc_texto"])}</div>', unsafe_allow_html=True)
                            renderizar_anexo(_d["anexo"], key=f"dl_anexo_hist_{_sel_hist[0]}")
                            _ar = str(_d.get("anexo_resposta", "") or "").strip()
                            if _ar and _ar.lower() not in ("nenhum", "nan", "none"):
                                st.caption("Imagem enviada pela TI na resposta:")
                                renderizar_anexo(_ar, key=f"dl_anexo_resp_hist_{_sel_hist[0]}")
                        else:
                            _just = html.escape(str(_d["justificativa"] or "—"))
                            _just_txt = _just if _just and _just != "—" else "Não informada."
                            st.write("")
                            st.markdown(f'<div class="meta-chip"><strong>Quantidade:</strong> &nbsp;{html.escape(str(_d["quantidade"]))}</div>', unsafe_allow_html=True)
                            st.write("")
                            st.markdown(f'<div class="meta-chip" style="height:auto;align-items:flex-start;padding:10px 12px;"><div><strong>Justificativa:</strong><br><span style="color:var(--t-n9);font-size:.78rem;">{_just_txt}</span></div></div>', unsafe_allow_html=True)
                else:
                    st.caption("Clique numa linha da tabela para ver os detalhes completos.")
            return

    _render_historico()
elif pagina == "gestor":

    @st.fragment
    def _render_gestor():
        global df_base, df_compras

        _MAX_TENTATIVAS    = 5
        _LOCKOUT_SEGUNDOS  = 300
        _SESSION_TTL       = 3600

        if "gestor_autenticado"        not in st.session_state:
            st.session_state.gestor_autenticado   = False
        if "gestor_auth_ts"            not in st.session_state:
            st.session_state.gestor_auth_ts       = 0.0
        if "gestor_tentativas"         not in st.session_state:
            st.session_state.gestor_tentativas    = 0
        if "gestor_lockout_ate"        not in st.session_state:
            st.session_state.gestor_lockout_ate   = 0.0

        _agora_ts = time.time()
        if st.session_state.gestor_autenticado:
            if (_agora_ts - st.session_state.gestor_auth_ts) > _SESSION_TTL:
                st.session_state.gestor_autenticado = False
                st.session_state.gestor_auth_ts     = 0.0

        if not st.session_state.gestor_autenticado:
            _em_lockout   = _agora_ts < st.session_state.gestor_lockout_ate
            _seg_restantes = max(0, int(st.session_state.gestor_lockout_ate - _agora_ts))
            _tentativas_rest = max(0, _MAX_TENTATIVAS - st.session_state.gestor_tentativas)

            st.markdown("""
<style>
        @keyframes _blink_g { 0%,100%{opacity:1} 50%{opacity:.15} }
        @keyframes _lockout_scan {
            0%   { background-position: -200% 0; }
            100% { background-position:  200% 0; }
        }
        /* Card topo */
        .tg-card {
            position: relative; overflow: hidden;
            background: rgba(var(--rgb-n0),.97);
            border: 1px solid var(--ta220);
            border-radius: 12px 12px 0 0;
            font-family: 'IBM Plex Mono', ui-monospace, monospace;
            box-shadow: 0 28px 70px var(--shc-1);
        }
        .tg-card.locked {
            border-color: rgba(var(--rgb-crit-l),.35) !important;
        }
        .tg-card::before {
            content:''; position:absolute; top:0; left:0; right:0; height:1px;
            background: linear-gradient(90deg,transparent,var(--ta700),transparent);
            background-size: 200% 1px; animation: none;
        }
        .tg-card.locked::before {
            background: linear-gradient(90deg,transparent,rgba(var(--rgb-crit-l),.8),transparent);
            background-size: 200% 1px;
        }
        .tg-bar {
            display:flex; align-items:center; gap:7px; padding:9px 16px;
            background:var(--ta040); border-bottom:1px solid var(--ta080);
        }
        .tg-bar.locked { background:rgba(var(--rgb-crit-l),.04); border-bottom-color:rgba(var(--rgb-crit-l),.08); }
        .tg-dot { width:8px; height:8px; border-radius:50%; }
        .tg-bar-title {
            flex:1; text-align:center;
            font-size:.50rem; letter-spacing:.22em; text-transform:uppercase;
            color:rgba(var(--rgb-c-acc),.30); margin:0;
        }
        .tg-bar.locked .tg-bar-title { color:rgba(var(--rgb-c-crit-l),.30); }
        .tg-body { padding:22px 24px 18px; }
        .tg-icon-ring {
            width:52px; height:52px; border-radius:13px;
            margin:0 auto 14px;
            display:flex; align-items:center; justify-content:center;
            background:linear-gradient(135deg,var(--ta150),rgba(var(--rgb-run-l),.08));
            border:1px solid var(--ta220);
            font-size:1.5rem;
        }
        .tg-icon-ring.locked {
            background:linear-gradient(135deg,rgba(var(--rgb-crit-l),.15),rgba(var(--rgb-run),.06));
            border-color:rgba(var(--rgb-crit-l),.30);
        }
        .tg-badge {
            display:inline-flex; align-items:center; gap:6px;
            font-size:.50rem; letter-spacing:.16em; text-transform:uppercase;
            color:rgba(var(--rgb-c-acc),.45); margin-bottom:8px;
        }
        .tg-badge.locked { color:rgba(var(--rgb-c-crit-l),.55); }
        .tg-badge-dot { width:5px; height:5px; border-radius:50%; background:var(--f-acc); flex-shrink:0; animation: none; }
        .tg-badge-dot.locked { background:var(--f-crit-l); }
        .tg-title { font-size:.96rem; font-weight:800; color:var(--c-acc-xl); margin-bottom:4px; text-align:center; }
        .tg-sub   { font-size:.62rem; color:rgba(var(--rgb-t-n9),.62); text-align:center; }
        .tg-sep { height:1px; margin:16px 0 14px; background:linear-gradient(90deg,transparent,var(--ta200),transparent); }
        .tg-sep.locked { background:linear-gradient(90deg,transparent,rgba(var(--rgb-crit-l),.18),transparent); }
        .tg-attempts { display:flex; align-items:center; gap:6px; margin-bottom:10px; justify-content:center; }
        .tg-att-dot { width:8px; height:8px; border-radius:50%; }
        .tg-att-dot.used  { background:var(--f-crit-l); box-shadow: none; }
        .tg-att-dot.avail { background:var(--ta250); border:1px solid var(--ta180); }
        .tg-lockout-bar {
            padding:11px 13px; border:1px solid rgba(var(--rgb-crit-l),.28); border-radius:8px;
            background:rgba(var(--rgb-crit-l),.06);
        }
        .tg-lockout-title { font-size:.62rem; font-weight:800; color:var(--c-crit-l); letter-spacing:.08em; text-transform:uppercase; margin-bottom:3px; }
        .tg-lockout-msg   { font-size:.60rem; color:rgba(var(--rgb-c-crit-l),.65); }
        .tg-warn-token {
            padding:9px 12px; margin-bottom:10px;
            border:1px solid rgba(var(--rgb-run),.24); border-radius:7px;
            background:rgba(var(--rgb-run),.05);
            font-size:.60rem; color:rgba(var(--rgb-c-run),.75); line-height:1.45;
        }
        .tg-warn-token strong { color:var(--c-run); }
        /* Form integrado abaixo do card */
        div[data-testid="stForm"][data-form-key="flg"] {
            background:rgba(var(--rgb-n0),.97) !important;
            border:1px solid var(--ta220) !important;
            border-top:none !important;
            border-radius:0 0 12px 12px !important;
            padding:4px 24px 18px !important;
            margin:0 !important;
            box-shadow: 0 28px 70px var(--shc-1) !important;
        }
        div[data-testid="stForm"][data-form-key="flg"] input[type="password"] {
            background:var(--ta040) !important;
            border:1px solid var(--ta180) !important;
            border-radius:7px !important; color:var(--c-acc-l) !important;
            font-family:'IBM Plex Mono',monospace !important;
            font-size:.78rem !important; letter-spacing:.06em !important;
            height:40px !important; padding:0 14px !important;
        }
        div[data-testid="stForm"][data-form-key="flg"] input[type="password"]:focus {
            border-color:var(--ta550) !important;
            box-shadow: 0 0 0 3px var(--ta100) !important;
        }
        div[data-testid="stForm"][data-form-key="flg"] .stFormSubmitButton button {
            width:100% !important; height:34px !important;
            background:transparent !important;
            border:1px solid var(--ta180) !important;
            border-radius:6px !important; color:rgba(var(--rgb-c-acc),.55) !important;
            font-family:'IBM Plex Mono',monospace !important;
            font-size:.56rem !important; font-weight:700 !important;
            letter-spacing:.18em !important; text-transform:uppercase !important;
            padding:0 !important; margin-top:10px !important;
            transition:all .18s !important;
        }
        div[data-testid="stForm"][data-form-key="flg"] .stFormSubmitButton button:hover {
            background:var(--ta080) !important;
            border-color:var(--ta420) !important;
            color:var(--c-acc) !important;
            box-shadow: none;
        }
        .tg-footer-g {
            display:flex; justify-content:space-between; align-items:center;
            padding:10px 0 0;
            font-family:'IBM Plex Mono',monospace; font-size:.46rem; letter-spacing:.10em;
        }
        </style>
        """, unsafe_allow_html=True)

            _dots_html = '<div class="tg-attempts">'
            for i in range(_MAX_TENTATIVAS):
                _cls = "used" if i < st.session_state.gestor_tentativas else "avail"
                _dots_html += f'<div class="tg-att-dot {_cls}"></div>'
            _dots_html += '</div>'

            _locked_cls = "locked" if _em_lockout else ""

            if _em_lockout:
                _card_extra = f"""
            <div class="tg-sep {_locked_cls}"></div>
            <div class="tg-lockout-bar">
              <div class="tg-lockout-title">Acesso bloqueado</div>
              <div class="tg-lockout-msg">Aguarde <strong style="color:var(--c-crit-l);">{_seg_restantes}s</strong> para tentar novamente.</div>
            </div>"""
            elif token_padrao_ativo():
                _card_extra = """
            <div class="tg-sep"></div>
            <div class="tg-warn-token">
              ⚠️ <strong>Token padrão ativo!</strong><br>
              Configure <code>PORTAL_SUPORTE_ADMIN_TOKEN</code> em <code>secrets.toml</code>.
            </div>"""
            else:
                _card_extra = f'<div class="tg-sep"></div>{_dots_html}'

            _ga, _gb, _gc = st.columns([1, 2.4, 1])
            with _gb:
                st.markdown(f"""
            <div class="tg-card {_locked_cls}">
              <div class="tg-bar {_locked_cls}">
                <div class="tg-dot" style="background:var(--f-crit-l);opacity:.65;"></div>
                <div class="tg-dot" style="background:var(--f-run);opacity:.65;"></div>
                <div class="tg-dot" style="background:var(--f-ok);opacity:.65;"></div>
                <span class="tg-bar-title">console · admin · gestor</span>
              </div>
              <div class="tg-body">
                <div class="tg-icon-ring {_locked_cls}">{_ico("cadeado", 26) if _em_lockout else _ico("escudo", 26)}</div>
                <div style="text-align:center;">
                  <div class="tg-badge {_locked_cls}">
                    <div class="tg-badge-dot {_locked_cls}"></div>
                    {"acesso bloqueado" if _em_lockout else "acesso restrito · console admin"}
                  </div>
                </div>
                <div class="tg-title">{"Conta Bloqueada" if _em_lockout else "Autenticação Requerida"}</div>
                <div class="tg-sub">{"Muitas tentativas inválidas." if _em_lockout else "Insira o token de administrador."}</div>
                {_card_extra}
              </div>
            </div>
            """, unsafe_allow_html=True)

                if not _em_lockout:
                    with st.form("flg", clear_on_submit=True):
                        st.text_input("t", type="password",
                                      placeholder="▸ token de acesso admin",
                                      key="pwd_gestor", label_visibility="collapsed")
                        _btn_submit = st.form_submit_button("Autenticar", use_container_width=True)
                        if _btn_submit:
                            _pwd_digitado = st.session_state.get("pwd_gestor", "")
                            if validar_token_admin(_pwd_digitado):
                                st.session_state.gestor_autenticado  = True
                                st.session_state.gestor_auth_ts      = time.time()
                                st.session_state.gestor_tentativas   = 0
                                st.session_state.gestor_lockout_ate  = 0.0
                                st.rerun(scope="app")
                            else:
                                st.session_state.gestor_tentativas += 1
                                if st.session_state.gestor_tentativas >= _MAX_TENTATIVAS:
                                    st.session_state.gestor_lockout_ate = time.time() + _LOCKOUT_SEGUNDOS
                                    st.markdown(f"""<div style="font-family:'IBM Plex Mono',monospace;font-size:.58rem;
                                    color:rgba(var(--rgb-c-crit-l),.9);text-align:center;padding:4px 0 0;">
                                    Bloqueado por {_LOCKOUT_SEGUNDOS//60} min.</div>""", unsafe_allow_html=True)
                                else:
                                    _rest = _MAX_TENTATIVAS - st.session_state.gestor_tentativas
                                    st.markdown(f"""<div style="font-family:'IBM Plex Mono',monospace;font-size:.58rem;
                                    color:rgba(var(--rgb-c-crit-l),.85);text-align:center;padding:4px 0 0;">
                                    ✗ &nbsp;Token inválido — {_rest} tentativa(s) restante(s)</div>""", unsafe_allow_html=True)

                st.markdown(f"""
            <div class="tg-footer-g">
              <span style="color:rgba(var(--rgb-c-acc),.18);">TLM Service Desk · v3.8</span>
              <span style="color:rgba(var(--rgb-c-ok),.35);display:flex;align-items:center;gap:5px;">
                <span style="width:4px;height:4px;border-radius:50%;background:var(--f-ok);
                  box-shadow: none;display:inline-block;"></span>
                hmac · sha256
              </span>
            </div>
            """, unsafe_allow_html=True)

            st.stop()


        _sessao_restante = max(0, int(_SESSION_TTL - (_agora_ts - st.session_state.gestor_auth_ts)))
        _sessao_min      = _sessao_restante // 60
        _agora_str       = datetime.now().strftime("%d/%m/%Y  %H:%M")

        st.markdown("""
<style>
    /* ── GESTOR: HERO BAR PREMIUM ── */
    .gestor-hero {
        position: relative;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 16px;
        padding: 20px 26px;
        margin-bottom: 22px;
        background: linear-gradient(135deg, rgba(var(--rgb-sup-a),.98), rgba(var(--rgb-sup-b),.95));
        border: 1px solid var(--ta180);
        border-radius: 14px;
        overflow: hidden;
        box-shadow: 0 2px 40px var(--shc-2), 0 0 0 1px var(--ta050);
    }
    .gestor-hero::before {
        content: "";
        position: absolute;
        top: 0; left: 0; right: 0; height: 2px;
        background: linear-gradient(90deg, var(--f-acc), var(--f-run-l), var(--f-acc));
        background-size: 200% 100%;
        animation: none;
    }
    @keyframes heroShimmer { 0%{background-position:0} 100%{background-position:200%} }
    .gestor-hero::after {
        content: "";
        position: absolute;
        right: 0; top: 0; bottom: 0;
        width: 35%;
        background: transparent;
        pointer-events: none;
    }
    .gestor-hero-left { position: relative; z-index: 1; }
    .gestor-hero-eyebrow {
        font-family: 'IBM Plex Mono', monospace;
        font-size: .58rem; font-weight: 800;
        letter-spacing: .18em; text-transform: uppercase;
        color: rgba(var(--rgb-c-acc),.60);
        display: flex; align-items: center; gap: 8px;
        margin-bottom: 6px;
    }
    .gestor-hero-eyebrow::before {
        content: "";
        width: 6px; height: 6px; border-radius: 50%;
        background: var(--f-ok);
        box-shadow: none;
        animation: none;
    }
    @keyframes gepulse { 0%,100%{opacity:.9} 50%{opacity:.5} }
    .gestor-hero-title {
        font-size: 1.32rem; font-weight: 800;
        color: var(--t-n11); letter-spacing: -.01em;
        margin: 0 0 4px;
        text-shadow: none;
    }
    .gestor-hero-desc {
        font-size: .76rem; color: rgba(var(--rgb-t-n9),.77);
        line-height: 1.5; max-width: 500px;
    }
    .gestor-hero-right {
        position: relative; z-index: 1;
        display: flex; flex-direction: column;
        align-items: flex-end; gap: 8px;
    }
    .gestor-session-badge {
        display: flex; align-items: center; gap: 6px;
        padding: 5px 12px;
        background: rgba(var(--rgb-ok),.06);
        border: 1px solid rgba(var(--rgb-ok),.18);
        border-radius: 999px;
        font-family: 'IBM Plex Mono', monospace;
        font-size: .58rem; font-weight: 700;
        letter-spacing: .08em; color: rgba(var(--rgb-c-ok),.75);
    }
    .gestor-session-badge::before {
        content: "";
        width: 5px; height: 5px; border-radius: 50%;
        background: var(--f-ok);
        box-shadow: none;
    }
    .gestor-datetime {
        font-family: 'IBM Plex Mono', monospace;
        font-size: .60rem; color: rgba(var(--rgb-t-n9),.59);
        letter-spacing: .06em;
    }
    /* ── OPS CARDS V2 ── */
    .ops-grid-v2 {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 14px;
        margin: 0 0 20px;
    }
    .ops-card-v2 {
        position: relative;
        overflow: hidden;
        padding: 18px 20px 16px;
        border: 1px solid rgba(var(--rgb-n9),.10);
        border-radius: 12px;
        background: linear-gradient(145deg, rgba(var(--rgb-sup-a),.96), rgba(var(--rgb-sup-b),.88));
        box-shadow: 0 8px 28px var(--shc-2), inset 0 1px 0 rgba(var(--rgb-lift),.025);
        transition: border-color .2s, transform .2s, box-shadow .2s;
        cursor: default;
    }
    .ops-card-v2:hover {
        border-color: var(--tone-c, var(--ta450));
        transform: translateY(-2px);
        box-shadow: 0 18px 44px var(--shc-2);
    }
    .ops-card-v2::before {
        content: "";
        position: absolute;
        top: 0; left: 0; right: 0; height: 2.5px;
        background: var(--tone-c, var(--f-acc));
        box-shadow: none;
        opacity: .85;
    }
    .ops-card-v2::after {
        content: "";
        position: absolute;
        bottom: 0; right: 0;
        width: 70px; height: 70px;
        background: transparent;
        pointer-events: none;
    }
    .ops-card-v2-icon {
        font-size: 1.4rem; margin-bottom: 10px; display: block; line-height: 1;
    }
    .ops-card-v2-label {
        font-family: 'IBM Plex Mono', monospace;
        font-size: .60rem; font-weight: 800;
        letter-spacing: .12em; text-transform: uppercase;
        color: rgba(var(--rgb-t-n9),.72); margin-bottom: 6px;
    }
    .ops-card-v2-val {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 2.4rem; font-weight: 800;
        line-height: 1; color: var(--tone-c, var(--t-n11));
        margin-bottom: 5px;
        text-shadow: none;
    }
    .ops-card-v2-copy {
        font-size: .72rem; color: rgba(var(--rgb-t-n8),.76);
        line-height: 1.35;
    }
    /* ── INSIGHT BANNER V2 ── */
    .insight-banner {
        display: flex; align-items: center; gap: 14px;
        padding: 14px 18px;
        background: linear-gradient(135deg, var(--ta070), rgba(var(--rgb-run-l),.04));
        border: 1px solid var(--ta150);
        border-radius: 10px;
        margin-bottom: 22px;
    }
    .insight-banner-icon {
        font-size: 1.1rem; flex-shrink: 0;
    }
    .insight-banner-text {
        font-size: .82rem; color: var(--t-n10); line-height: 1.5;
    }
    .insight-banner-text strong { color: var(--c-acc-l); }
    /* ── TABS V2 ── */
    [data-testid="stTabs"] > div:first-child {
        border-bottom: 1px solid var(--ta140) !important;
        gap: 2px !important;
    }
    [data-testid="stTabs"] button[role="tab"] {
        font-family: 'IBM Plex Mono', monospace !important;
        font-size: .68rem !important;
        font-weight: 800 !important;
        letter-spacing: .10em !important;
        text-transform: uppercase !important;
        color: rgba(var(--rgb-t-n9),.72) !important;
        padding: 10px 20px !important;
        border-radius: 8px 8px 0 0 !important;
        transition: color .18s, background .18s !important;
        border: none !important;
    }
    [data-testid="stTabs"] button[role="tab"]:hover {
        color: var(--t-n11) !important;
        background: var(--ta060) !important;
    }
    [data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
        color: var(--c-acc) !important;
        background: var(--ta100) !important;
        border-bottom: 2px solid var(--c-acc) !important;
    }
    </style>
    """, unsafe_allow_html=True)

        st.markdown(f"""
    <div class="gestor-hero animate-fade-in">
        <div class="gestor-hero-left">
            <div class="gestor-hero-eyebrow">Acesso restrito · Gestão de TI</div>
            <div class="gestor-hero-title">Painel Administrativo</div>
            <div class="gestor-hero-desc">Gerencie a fila de trabalho, analise métricas em tempo real e decida sobre requisições logísticas.</div>
        </div>
        <div class="gestor-hero-right">
            <div class="gestor-session-badge">sessão ativa · {_sessao_min}min restantes</div>
            <div class="gestor-datetime">{_ico("relogio", 12, extra="margin-right:5px")}{_agora_str}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

        _col_logout_g = st.columns([6, 1])[1]
        with _col_logout_g:
            if st.button("Encerrar Sessão", key="gestor_logout", use_container_width=True, help="Encerrar sessão administrativa"):
                st.session_state.gestor_autenticado = False
                st.session_state.gestor_auth_ts     = 0.0
                st.rerun(scope="app")

        df_base_live = carregar_dados_cacheados(DATABASE, COLUNAS)
        df_compras_live = carregar_dados_cacheados(DATABASE_COMPRAS, COLUNAS_COMPRAS)

        if "qtd_chamados_antiga" not in st.session_state:
            st.session_state.qtd_chamados_antiga = len(df_base_live)
        if "qtd_compras_antiga" not in st.session_state:
            st.session_state.qtd_compras_antiga = len(df_compras_live)

        if len(df_base_live) > st.session_state.qtd_chamados_antiga:
            ultimo_c = df_base_live.iloc[-1]
            st.toast(f"Novo Chamado de {ultimo_c['Usuário']}", icon="🔔")
            st.session_state.qtd_chamados_antiga = len(df_base_live)

        if len(df_compras_live) > st.session_state.qtd_compras_antiga:
            ultima_c = df_compras_live.iloc[-1]
            st.toast(f"Nova Compra de {ultima_c['Usuário']}", icon="🛍️")
            st.session_state.qtd_compras_antiga = len(df_compras_live)

        df_base = df_base_live
        df_compras = df_compras_live

        df_ch_ativos = df_base[df_base["Status"].isin(["Aberto", "Em Atendimento"])] if not df_base.empty else pd.DataFrame(columns=COLUNAS)
        df_cp_ativos = df_compras[df_compras["Status"].isin(["Pendente", "Em Análise"])] if not df_compras.empty else pd.DataFrame(columns=COLUNAS_COMPRAS)
        crit_ch = int((df_ch_ativos["Prioridade"] == "Crítica").sum()) if not df_ch_ativos.empty else 0
        alta_ch = int((df_ch_ativos["Prioridade"] == "Alta").sum()) if not df_ch_ativos.empty else 0
        fila_ordenada = ordenar_por_prioridade_e_data(df_ch_ativos, "Prioridade") if not df_ch_ativos.empty else df_ch_ativos
        proximo = fila_ordenada.iloc[0] if not fila_ordenada.empty else None
        if proximo is not None:
            prox_texto = (
                f"#{formatar_id(proximo.name)} {primeiro_valor(proximo, 'Protocolo', padrao='sem protocolo')} - "
                f"{primeiro_valor(proximo, 'Usuário')} / {primeiro_valor(proximo, 'Setor')} - "
                f"{primeiro_valor(proximo, 'Prioridade')}"
            )
        else:
            prox_texto = "Sem chamado ativo no momento"

        st.markdown('<div class="sec-label animate-fade-in"><div class="sec-dot"></div>SITUAÇÃO ATUAL</div>', unsafe_allow_html=True)
        _crit_color = "var(--c-crit-l)" if crit_ch > 0 else "var(--c-ok)"
        st.markdown(f"""
        <div class="ops-grid-v2 animate-fade-in">
            <div class="ops-card-v2" style="--tone-c:{_crit_color}">
                <span class="ops-card-v2-icon">{_ico("alerta", 20)}</span>
                <div class="ops-card-v2-label">Chamados Críticos</div>
                <div class="ops-card-v2-val">{crit_ch}</div>
                <div class="ops-card-v2-copy">Chamados com prioridade crítica na fila ativa.</div>
            </div>
            <div class="ops-card-v2" style="--tone-c:var(--c-run)">
                <span class="ops-card-v2-icon">{_ico("raio", 20)}</span>
                <div class="ops-card-v2-label">Alta Prioridade</div>
                <div class="ops-card-v2-val">{alta_ch}</div>
                <div class="ops-card-v2-copy">Chamados de alta prioridade aguardando atendimento.</div>
            </div>
            <div class="ops-card-v2" style="--tone-c:var(--c-acc)">
                <span class="ops-card-v2-icon">{_ico("chamado", 20)}</span>
                <div class="ops-card-v2-label">Fila Ativa</div>
                <div class="ops-card-v2-val">{len(df_ch_ativos)}</div>
                <div class="ops-card-v2-copy">Tickets abertos ou em atendimento agora.</div>
            </div>
            <div class="ops-card-v2" style="--tone-c:var(--c-run-l)">
                <span class="ops-card-v2-icon">{_ico("carrinho", 20)}</span>
                <div class="ops-card-v2-label">Compras Pendentes</div>
                <div class="ops-card-v2-val">{len(df_cp_ativos)}</div>
                <div class="ops-card-v2-copy">Requisições aguardando análise ou decisão de TI.</div>
            </div>
        </div>
        <div class="insight-banner animate-fade-in">
            <span class="insight-banner-icon">{_ico("alvo", 18)}</span>
            <div class="insight-banner-text">
                <strong>Próxima ação recomendada:</strong> {html.escape(prox_texto)}
            </div>
        </div>
    """, unsafe_allow_html=True)

        @st.dialog("Analisar Requisição de Compra")
        def modal_responder_compra(idx_c, row_c):
            st.markdown(f"### Requisição {primeiro_valor(row_c, 'Protocolo', padrao=f'#{formatar_id(idx_c)}')}")
            st.write(f"**Solicitante:** {row_c['Usuário']} | **Setor:** {row_c['Setor']}")
            st.markdown(f"**Item Solicitado:** `{row_c['Item']}` (Quantidade: {row_c['Quantidade']})")
            st.write(f"**Categoria:** {row_c['Categoria']} | **Urgência:** {row_c['Urgência']}")
            st.markdown(f"> **Justificativa:** {row_c['Justificativa']}")
            st.markdown("---")
        
            novo_status_c = st.selectbox(
                "Decisão de TI / Governança:",
                ["Pendente", "Em Análise", "Aprovado", "Recusado", "Concluído"],
                index=["Pendente", "Em Análise", "Aprovado", "Recusado", "Concluído"].index(row_c["Status"])
            )
            obs_ti = st.text_area(
                "Despacho Administrativo / Orçamentário:",
                value=row_c["Observação TI"] if row_c["Observação TI"] != "Aguardando análise" else "",
                placeholder="Insira os motivos da decisão ou cotações de valores..."
            )
            responsavel_compra = st.text_input(
                "Responsável pelo acompanhamento:",
                value=primeiro_valor(row_c, "Responsável", padrao=""),
                placeholder="Ex: Samuel / Hiago"
            )
        
            if st.button("Registrar Parecer e Fechar", use_container_width=True, type="primary"):
                df_compras.at[idx_c, "Status"] = novo_status_c
                df_compras.at[idx_c, "Observação TI"] = obs_ti
                df_compras.at[idx_c, "Responsável"] = responsavel_compra.strip() or "Não atribuído"
                df_compras.at[idx_c, "Atualizado em"] = datetime.now().strftime("%d/%m/%Y %H:%M")
                salvar_dataframe_seguro(df_compras, DATABASE_COMPRAS, COLUNAS_COMPRAS)
                st.success("Parecer gravado com sucesso!")
                st.rerun(scope="app")

        st.markdown("""
<style>
    .refresh-bar {
        display: flex; align-items: center; justify-content: flex-end;
        gap: 10px; margin: 0 0 10px 0;
    }
    .refresh-ts {
        font-family: var(--mono, monospace);
        font-size: 0.72rem;
        color: rgba(var(--rgb-t-n8),.79);
        letter-spacing: 0.04em;
    }
    div[data-testid="stButton"] button[kind="secondary"].refresh-btn {
        background: transparent !important;
        border: 1px solid var(--ta220) !important;
        color: rgba(var(--rgb-c-acc),.85) !important;
        font-family: var(--mono, monospace) !important;
        font-size: 0.72rem !important;
        letter-spacing: 0.08em !important;
        padding: 2px 14px !important;
        border-radius: 6px !important;
        transition: border-color .2s, box-shadow .2s, color .2s !important;
    }
    div[data-testid="stButton"] button[kind="secondary"].refresh-btn:hover {
        border-color: var(--ta700) !important;
        box-shadow: none;
        color: var(--c-acc) !important;
    }
    </style>
    """, unsafe_allow_html=True)
        _ref_col_l, _ref_col_r = st.columns([6, 1])
        with _ref_col_r:
            if st.button("Atualizar dados", key="btn_refresh_manual", help="Recarregar dados do banco SQLite agora", use_container_width=True):
                carregar_dados_cacheados.clear()
                st.rerun(scope="fragment")
        with _ref_col_l:
            _ts_atual = datetime.now().strftime("%d/%m/%Y  %H:%M:%S")
            st.markdown(f'<div class="refresh-ts">última atualização · {_ts_atual}</div>', unsafe_allow_html=True)

        st.markdown('<div class="sec-label animate-fade-in"><div class="sec-dot"></div>ÁREAS DE GESTÃO</div>', unsafe_allow_html=True)
        tab_analytics, tab_chamados, tab_compras = st.tabs([
            "Indicadores",
            "Governança de Chamados",
            "Auditoria de Compras",
        ])

        _ICONE_PRIO = {"Baixa": "🔵", "Média": "🟡", "Alta": "🟠", "Crítica": "🔴"}
        _ICONE_STATUS_CH = {"Aberto": "🔴 Aberto", "Em Atendimento": "🟡 Em Atendimento", "Concluído": "🟢 Concluído"}
        _ICONE_STATUS_CP = {
            "Pendente": "⏳ Pendente", "Em Análise": "🔍 Em Análise", "Aprovado": "✅ Aprovado",
            "Recusado": "❌ Recusado", "Concluído": "🟢 Concluído",
        }

        def _dt_abertura(v):
            """Data/hora de abertura como datetime real.

            Guardamos a data como texto "dd/mm/aaaa hh:mm"; se ela fosse para a
            grade assim, a ordenação por essa coluna seria alfabética (01/12
            viria antes de 02/01). Convertida, o st.dataframe ordena de fato por
            tempo e formata a exibição pelo column_config."""
            return pd.to_datetime(str(v).strip(), format="%d/%m/%Y %H:%M",
                                  errors="coerce")

        _COL_ABERTO = st.column_config.DatetimeColumn(
            "Aberto em", format="DD/MM/YYYY HH:mm", width="medium",
            help="Data e hora em que o registro foi aberto",
        )

        def _truncar(texto, tam=120):
            texto = str(texto).strip()
            return texto[:tam] + "…" if len(texto) > tam else texto

        def _tabela_fila_chamados(df):
            linhas = [{
                "Protocolo": primeiro_valor(row, "Protocolo", padrao=f"#{formatar_id(idx)}"),
                "Aberto em": _dt_abertura(row.get("Data", "")),
                "Usuário": row.get("Usuário", ""),
                "Setor": row.get("Setor", ""),
                "Categoria": extrair_categoria(row),
                "Prioridade": f"{_ICONE_PRIO.get(str(row.get('Prioridade', '')).strip(), '⚪')} {row.get('Prioridade', '—')}",
                "Status": _ICONE_STATUS_CH.get(str(row.get("Status", "")).strip(), row.get("Status", "—")),
                "Em fila": tempo_decorrido(row.get("Data", "")),
                "Anexo": "📎 Sim" if str(row.get("Anexo", "Nenhum")) not in ["Nenhum", "", "nan"] else "—",
                "Descrição": _truncar(row.get("Descrição", "")),
            } for idx, row in df.iterrows()]
            return pd.DataFrame(linhas, index=df.index)

        def _tabela_historico_chamados(df):
            linhas = [{
                "Protocolo": primeiro_valor(row, "Protocolo", padrao=f"#{formatar_id(idx)}"),
                "Aberto em": _dt_abertura(row.get("Data", "")),
                "Usuário": row.get("Usuário", ""),
                "Setor": row.get("Setor", ""),
                "Categoria": extrair_categoria(row),
                "Prioridade": str(row.get("Prioridade", "—")).strip() or "—",
                "Concluído em": _dt_abertura(row.get("Atualizado em") or row.get("Data", "")),
            } for idx, row in df.iterrows()]
            return pd.DataFrame(linhas, index=df.index)

        def _tabela_fila_compras(df):
            linhas = [{
                "Protocolo": primeiro_valor(row, "Protocolo", padrao=f"#{formatar_id(idx)}"),
                "Aberto em": _dt_abertura(row.get("Data", "")),
                "Item": row.get("Item", ""),
                "Usuário": row.get("Usuário", ""),
                "Setor": row.get("Setor", ""),
                "Categoria": row.get("Categoria", ""),
                "Urgência": str(row.get("Urgência", "—")).strip() or "—",
                "Qtd": row.get("Quantidade", ""),
                "Status": _ICONE_STATUS_CP.get(str(row.get("Status", "")).strip(), row.get("Status", "—")),
                "Justificativa": _truncar(row.get("Justificativa", "")),
            } for idx, row in df.iterrows()]
            return pd.DataFrame(linhas, index=df.index)

        def _tabela_historico_compras(df):
            linhas = [{
                "Protocolo": primeiro_valor(row, "Protocolo", padrao=f"#{formatar_id(idx)}"),
                "Aberto em": _dt_abertura(row.get("Data", "")),
                "Item": row.get("Item", ""),
                "Usuário": row.get("Usuário", ""),
                "Setor": row.get("Setor", ""),
                "Categoria": row.get("Categoria", ""),
                "Qtd": row.get("Quantidade", ""),
                "Status": _ICONE_STATUS_CP.get(str(row.get("Status", "")).strip(), row.get("Status", "—")),
            } for idx, row in df.iterrows()]
            return pd.DataFrame(linhas, index=df.index)

        def _linha_selecionada(evento, df):
            """(idx, linha) marcados na grade, mapeados de volta ao índice real do DataFrame de origem."""
            if not evento or not evento.selection or not evento.selection.rows:
                return None, None
            idx = df.index[evento.selection.rows[0]]
            return idx, df.loc[idx]

        @st.fragment
        def renderizar_analytics_realtime():
            df_b_all = carregar_dados_cacheados(DATABASE, COLUNAS)
            df_c_all = carregar_dados_cacheados(DATABASE_COMPRAS, COLUNAS_COMPRAS)

            import io as _io, zipfile as _zf
            _agora_f = datetime.now()
            _meses_pt = {
                1:"Janeiro",  2:"Fevereiro", 3:"Março",    4:"Abril",
                5:"Maio",     6:"Junho",     7:"Julho",     8:"Agosto",
                9:"Setembro", 10:"Outubro",  11:"Novembro", 12:"Dezembro"
            }
            _anos_disp = list(range(max(_agora_f.year - 3, 2022), _agora_f.year + 1))

            st.markdown('<div class="sec-label animate-fade-in"><div class="sec-dot"></div>FILTRO DE PERÍODO</div>', unsafe_allow_html=True)
            st.caption("Exibe apenas o mês selecionado — os contadores reiniciam a cada período.")

            _fc1, _fc2, _fc3, _fc4 = st.columns([1.3, 0.9, 1.8, 1.1])
            with _fc1:
                _mes_sel = st.selectbox(
                    "Mês",
                    list(_meses_pt.keys()),
                    format_func=lambda m: _meses_pt[m],
                    index=_agora_f.month - 1,
                    key="an_mes_sel",
                    label_visibility="visible"
                )
            with _fc2:
                _ano_sel = st.selectbox(
                    "Ano",
                    _anos_disp,
                    index=len(_anos_disp) - 1,
                    key="an_ano_sel",
                    label_visibility="visible"
                )
            with _fc3:
                _label_periodo = f"{_meses_pt[_mes_sel]} / {_ano_sel}"
                st.markdown(
                    f"""<div style="display:flex;align-items:flex-end;height:100%;padding-bottom:8px;">
                <span style="font-family:var(--mono,monospace);font-size:.72rem;color:var(--c-acc);
                border:1px solid var(--ta280);padding:6px 14px;border-radius:7px;
                background:var(--ta060);white-space:nowrap;">
                {_ico("calendario", 13, extra="margin-right:6px")}Exibindo: <strong>{_label_periodo}</strong> — contagem mensal</span></div>""",
                    unsafe_allow_html=True
                )
            with _fc4:
                _zip_bytes = _gerar_export_periodo_zip(
                    df_b_all, df_c_all, _mes_sel, _ano_sel, _meses_pt[_mes_sel]
                )
                st.download_button(
                    "Exportar período",
                    data=_zip_bytes,
                    file_name=f"analytics_{_meses_pt[_mes_sel].lower()}_{_ano_sel}.zip",
                    mime="application/zip",
                    use_container_width=True,
                    key=f"dl_period_{_mes_sel}_{_ano_sel}"
                )

            df_b_live = _filtrar_periodo_cacheado(df_b_all, _mes_sel, _ano_sel)
            df_c_live = _filtrar_periodo_cacheado(df_c_all, _mes_sel, _ano_sel)

            if df_b_live.empty and df_c_live.empty:
                st.markdown(
                    f"""<div style="padding:14px 18px;border:1px dashed var(--ta280);
                border-radius:10px;background:var(--ta040);margin:8px 0 16px;">
                <strong style="color:var(--c-acc);">Sem dados em {_label_periodo}</strong>
                <span style="color:var(--t-n8);font-size:.82rem;display:block;margin-top:4px;">
                Nenhum chamado ou compra registrado neste período. Os contadores iniciam em zero a cada mês.</span>
                </div>""",
                    unsafe_allow_html=True
                )

            total_ch    = len(df_b_live)
            abertos_ch  = len(df_b_live[df_b_live['Status'] == 'Aberto']) if not df_b_live.empty else 0
            em_atend_ch = len(df_b_live[df_b_live['Status'] == 'Em Atendimento']) if not df_b_live.empty else 0
            concluidos  = len(df_b_live[df_b_live['Status'] == 'Concluído']) if not df_b_live.empty else 0
            total_comp  = len(df_c_live)
            taxa_resol  = f"{int(concluidos/total_ch*100)}%" if total_ch > 0 else "—"

            import streamlit.components.v1 as _comp

            taxa_resol_num = int(concluidos / total_ch * 100) if total_ch > 0 else 0
            pressao_fila = round((abertos_ch + em_atend_ch) / max(total_ch, 1) * 100)
            saude_label = "Crítica" if pressao_fila > 70 else ("Atenção" if pressao_fila > 40 else "Saudável")
            saude_cor = "var(--c-crit-l)" if pressao_fila > 70 else ("var(--c-run)" if pressao_fila > 40 else "var(--c-ok)")
            saude_icon = "🔴" if pressao_fila > 70 else ("🟡" if pressao_fila > 40 else "🟢")

            kpi_html = f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
*{{box-sizing:border-box;margin:0;padding:0;font-family:'Archivo','Inter',sans-serif}}
body{{background:transparent;padding:0 0 4px}}
.row{{display:grid;grid-template-columns:1.8fr 1fr 1fr 1fr 1fr 1fr;gap:10px}}

/* ── HEALTH CARD (first) ── */
.health{{
  background:linear-gradient(145deg,rgba(var(--rgb-sup-a),.98),rgba(var(--rgb-sup-b),.94));
  border:1px solid var(--ta140);border-radius:14px;
  padding:16px 18px;position:relative;overflow:hidden;
  display:flex;flex-direction:column;gap:8px;
}}
.health::before{{content:'';position:absolute;top:0;left:0;right:0;height:2px;
  background:linear-gradient(90deg,var(--f-acc),var(--f-run-l));
  box-shadow: none}}
.health-top{{display:flex;align-items:center;justify-content:space-between}}
.health-title{{font-size:.58rem;text-transform:uppercase;letter-spacing:.14em;color:var(--t-n6);font-weight:800}}
.health-badge{{
  font-size:.6rem;font-weight:800;letter-spacing:.08em;text-transform:uppercase;
  padding:3px 10px;border-radius:999px;
  background:color-mix(in srgb,{saude_cor} 12%,transparent);
  color:{saude_cor};border:1px solid color-mix(in srgb,{saude_cor} 35%,transparent);
}}
.health-body{{display:flex;align-items:center;gap:14px}}
.ring-wrap{{position:relative;flex-shrink:0;width:68px;height:68px}}
.ring-wrap svg{{transform:rotate(-90deg)}}
.ring-label{{
  position:absolute;inset:0;display:flex;flex-direction:column;
  align-items:center;justify-content:center;
}}
.ring-pct{{font-size:.95rem;font-weight:800;color:var(--t-n11);font-family:monospace;line-height:1}}
.ring-sub{{font-size:.48rem;text-transform:uppercase;letter-spacing:.1em;color:var(--t-n5);margin-top:1px}}
.health-stats{{flex:1;display:flex;flex-direction:column;gap:5px}}
.stat-row{{display:flex;align-items:center;justify-content:space-between;gap:6px}}
.stat-label{{font-size:.62rem;color:var(--t-n8);font-weight:600}}
.stat-bar{{flex:1;height:4px;background:rgba(var(--rgb-lift),.06);border-radius:4px;overflow:hidden}}
.stat-fill{{height:100%;border-radius:4px;transition:width 1s cubic-bezier(.4,0,.2,1)}}
.stat-val{{font-size:.62rem;font-weight:800;font-family:monospace;min-width:20px;text-align:right}}

/* ── KPI CARD (others) ── */
.k{{
  background:linear-gradient(145deg,rgba(var(--rgb-sup-a),.96),rgba(var(--rgb-sup-b),.90));
  border:1px solid var(--ta100);border-radius:14px;
  padding:14px 14px 12px;position:relative;overflow:hidden;
  transition:transform .25s,border-color .25s,box-shadow .25s;cursor:default;
  display:flex;flex-direction:column;justify-content:space-between;
}}
.k:hover{{
  transform:translateY(-3px);
  border-color:color-mix(in srgb,var(--c) 50%,transparent);
  box-shadow: 0 12px 28px var(--shc-1);
}}
.k::before{{content:'';position:absolute;top:0;left:0;right:0;height:2px;
  background:var(--c);opacity:.7;
  box-shadow: none}}
.k-top{{display:flex;align-items:flex-start;justify-content:space-between}}
.k-ico{{font-size:1.1rem;line-height:1}}
.k-trend{{
  font-size:.55rem;font-weight:800;letter-spacing:.06em;padding:2px 7px;border-radius:999px;
  background:rgba(var(--rgb-lift),.04);border:1px solid rgba(var(--rgb-lift),.07);color:var(--t-n5);
}}
.k-v{{
  font-family:'IBM Plex Mono','Courier New',monospace;font-size:2rem;font-weight:800;
  color:var(--c);line-height:1;margin:6px 0 2px;letter-spacing:-.03em;
}}
.k-l{{font-size:.6rem;text-transform:uppercase;letter-spacing:.11em;color:var(--t-n5);font-weight:700}}
.k-sub{{font-size:.58rem;color:var(--t-n6);margin-top:4px;font-weight:500}}
</style></head><body>
<div class="row">

  <!-- ── HEALTH CARD ── -->
  <div class="health">
    <div class="health-top">
      <span class="health-title">Saúde da Fila {saude_icon}</span>
      <span class="health-badge">{saude_label}</span>
    </div>
    <div class="health-body">
      <div class="ring-wrap">
        <svg width="68" height="68" viewBox="0 0 68 68">
          <circle cx="34" cy="34" r="28" fill="none" stroke="rgba(var(--rgb-lift),.06)" stroke-width="7"/>
          <circle cx="34" cy="34" r="28" fill="none"
            stroke="{saude_cor}" stroke-width="7"
            stroke-linecap="round"
            stroke-dasharray="{round(2*3.14159*28*taxa_resol_num/100, 1)} {round(2*3.14159*28*(1-taxa_resol_num/100), 1)}"
            style="filter:"/>
        </svg>
        <div class="ring-label">
          <span class="ring-pct" id="rpct">0%</span>
          <span class="ring-sub">resolvido</span>
        </div>
      </div>
      <div class="health-stats">
        <div class="stat-row">
          <span class="stat-label">Fila ativa</span>
          <div class="stat-bar"><div class="stat-fill" style="width:{min(abertos_ch/max(total_ch,1)*100,100):.0f}%;background:var(--f-crit-l)"></div></div>
          <span class="stat-val" style="color:var(--c-crit-l)">{abertos_ch}</span>
        </div>
        <div class="stat-row">
          <span class="stat-label">Em atend.</span>
          <div class="stat-bar"><div class="stat-fill" style="width:{min(em_atend_ch/max(total_ch,1)*100,100):.0f}%;background:var(--f-run)"></div></div>
          <span class="stat-val" style="color:var(--c-run)">{em_atend_ch}</span>
        </div>
        <div class="stat-row">
          <span class="stat-label">Concluídos</span>
          <div class="stat-bar"><div class="stat-fill" style="width:{min(concluidos/max(total_ch,1)*100,100):.0f}%;background:var(--f-ok)"></div></div>
          <span class="stat-val" style="color:var(--c-ok)">{concluidos}</span>
        </div>
      </div>
    </div>
  </div>

  <!-- ── KPI: TOTAL CHAMADOS ── -->
  <div class="k" style="--c:var(--c-acc)">
    <div class="k-top">
      <span class="k-ico">{_ico("chamado", 16)}</span>
      <span class="k-trend">PERÍODO</span>
    </div>
    <div class="k-v" data-t="{total_ch}">0</div>
    <div class="k-l">Total de chamados</div>
    <div class="k-sub">volume do mês</div>
  </div>

  <!-- ── KPI: FILA ABERTA ── -->
  <div class="k" style="--c:var(--c-crit-l)">
    <div class="k-top">
      <span class="k-ico">{_ico("alerta", 16)}</span>
      <span class="k-trend">URGENTE</span>
    </div>
    <div class="k-v" data-t="{abertos_ch}">0</div>
    <div class="k-l">Aguardando</div>
    <div class="k-sub">{"acima do ideal" if abertos_ch > 5 else "dentro do esperado"}</div>
  </div>

  <!-- ── KPI: EM ATENDIMENTO ── -->
  <div class="k" style="--c:var(--c-run)">
    <div class="k-top">
      <span class="k-ico">{_ico("engrenagem", 16)}</span>
      <span class="k-trend">ATIVO</span>
    </div>
    <div class="k-v" data-t="{em_atend_ch}">0</div>
    <div class="k-l">Em atendimento</div>
    <div class="k-sub">técnico alocado</div>
  </div>

  <!-- ── KPI: CONCLUÍDOS ── -->
  <div class="k" style="--c:var(--c-ok)">
    <div class="k-top">
      <span class="k-ico">{_ico("ok", 16)}</span>
      <span class="k-trend">OK</span>
    </div>
    <div class="k-v" data-t="{concluidos}">0</div>
    <div class="k-l">Encerrados</div>
    <div class="k-sub">resolvidos no período</div>
  </div>

  <!-- ── KPI: COMPRAS ── -->
  <div class="k" style="--c:var(--c-run-l)">
    <div class="k-top">
      <span class="k-ico">{_ico("carrinho", 16)}</span>
      <span class="k-trend">LOGÍSTICA</span>
    </div>
    <div class="k-v" data-t="{total_comp}">0</div>
    <div class="k-l">Requisições</div>
    <div class="k-sub">solicitações de compra</div>
  </div>

</div>

<script>
// Animação dos contadores
document.querySelectorAll('.k-v[data-t]').forEach(el=>{{
  const target=+el.dataset.t,dur=1100;let start=null;
  const ease=t=>t<.5?2*t*t:-1+(4-2*t)*t;
  (function step(ts){{
    if(!start)start=ts;
    const p=Math.min((ts-start)/dur,1);
    el.textContent=Math.round(ease(p)*target);
    if(p<1)requestAnimationFrame(step);
  }})(performance.now());
}});
// Animação do anel
(()=>{{
  const el=document.getElementById('rpct');
  const target={taxa_resol_num},dur=1200;let start=null;
  const ease=t=>1-Math.pow(1-t,3);
  (function step(ts){{
    if(!start)start=ts;
    const p=Math.min((ts-start)/dur,1);
    el.textContent=Math.round(ease(p)*target)+'%';
    if(p<1)requestAnimationFrame(step);
  }})(performance.now());
}})();
</script></body></html>"""
            _comp.html(_resolver_vars(kpi_html), height=142, scrolling=False)

            import streamlit.components.v1 as components
            import json

            st.markdown('<div class="sec-label animate-fade-in"><div class="sec-dot"></div>INDICADORES DO PERÍODO</div>', unsafe_allow_html=True)

            if not df_b_live.empty:
                df_setor = df_b_live['Setor'].value_counts().reset_index()
                df_setor.columns = ['Setor', 'Chamados']
                df_setor = df_setor.sort_values('Chamados', ascending=True)

                df_prio = df_b_live['Prioridade'].value_counts().reset_index()
                df_prio.columns = ['Prioridade', 'Qtd']
                todas_prio = pd.DataFrame({'Prioridade': ['Baixa','Média','Alta','Crítica']})
                df_prio = todas_prio.merge(df_prio, on='Prioridade', how='left').fillna(0)
                df_prio['Qtd'] = df_prio['Qtd'].astype(int)

                df_time2 = df_b_live.copy()
                df_time2['DataParsed'] = pd.to_datetime(df_time2['Data'], format="%d/%m/%Y %H:%M", errors='coerce')
                df_time2 = df_time2.dropna(subset=['DataParsed'])
                df_time2['Dia'] = df_time2['DataParsed'].dt.strftime('%d/%m')
                df_dia2 = df_time2.groupby('Dia').size().reset_index(name='Chamados')
                df_dia2['_s'] = pd.to_datetime(df_time2.groupby('Dia')['DataParsed'].min().values)
                df_dia2 = df_dia2.sort_values('_s').drop(columns=['_s']).tail(15)

                df_top2 = df_b_live['Usuário'].value_counts().head(8).reset_index()
                df_top2.columns = ['Usuário','Chamados']
                df_top2['Usuário'] = df_top2['Usuário'].str.title().str[:22]
                df_top2 = df_top2.sort_values('Chamados', ascending=True)

                SL  = json.dumps(df_setor['Setor'].tolist())
                SV  = json.dumps(df_setor['Chamados'].tolist())
                PL  = json.dumps(df_prio['Prioridade'].tolist())
                PV  = json.dumps(df_prio['Qtd'].tolist())
                TL  = json.dumps(df_dia2['Dia'].tolist())
                TV  = json.dumps(df_dia2['Chamados'].tolist())
                UL  = json.dumps(df_top2['Usuário'].tolist())
                UV  = json.dumps(df_top2['Chamados'].tolist())

                if not df_c_live.empty:
                    df_cat2 = df_c_live['Categoria'].value_counts().reset_index()
                    df_cat2.columns = ['Categoria','Pedidos']
                    df_cat2 = df_cat2.sort_values('Pedidos', ascending=True)
                    CL = json.dumps(df_cat2['Categoria'].tolist())
                    CV = json.dumps(df_cat2['Pedidos'].tolist())
                else:
                    CL = '[]'; CV = '[]'

                top_setor = df_setor.iloc[-1]['Setor'] if not df_setor.empty else "—"
                top_setor_v = int(df_setor.iloc[-1]['Chamados']) if not df_setor.empty else 0
                top_prio = df_prio.sort_values('Qtd',ascending=False).iloc[0] if not df_prio.empty else None
                top_prio_label = top_prio['Prioridade'] if top_prio is not None else "—"
                top_prio_v = int(top_prio['Qtd']) if top_prio is not None else 0
                max_dia = df_dia2.loc[df_dia2['Chamados'].idxmax(),'Dia'] if not df_dia2.empty else "—"
                max_dia_v = int(df_dia2['Chamados'].max()) if not df_dia2.empty else 0
                top_user = df_top2.iloc[-1]['Usuário'] if not df_top2.empty else "—"
                top_user_v = int(df_top2.iloc[-1]['Chamados']) if not df_top2.empty else 0
                status_vals = [abertos_ch, em_atend_ch, concluidos]
                status_labels = ["Aberto", "Em Atendimento", "Concluído"]
                STV = json.dumps(status_vals)
                STL = json.dumps(status_labels)

                DASH_HTML = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
{_tokens_vars(_tema_atual())}
*{{box-sizing:border-box;margin:0;padding:0}}
html,body{{background:transparent;font-family:'IBM Plex Mono',ui-monospace,monospace;overflow-x:hidden}}
.dash{{display:grid;grid-template-columns:1fr 1fr;gap:12px;padding:2px;align-content:start}}
.wide{{grid-column:1/-1}}
.card{{
  background:linear-gradient(145deg,rgba(var(--rgb-sup2-a),.98),rgba(var(--rgb-sup2-b),.94));
  border:1px solid var(--ta100);
  border-radius:12px;padding:15px 17px 13px;
  position:relative;overflow:hidden;
  box-shadow: 0 8px 28px var(--shc-2), inset 0 1px 0 rgba(var(--rgb-lift),.03);
}}
.card::before{{
  content:'';position:absolute;top:0;left:0;right:0;height:2px;
  background:linear-gradient(90deg,var(--accent,var(--f-acc)),color-mix(in srgb,var(--accent,var(--c-acc)) 30%,transparent));
  box-shadow: none;opacity:.7;
}}
.card-hdr{{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;gap:8px}}
.card-title{{font-size:9.5px;font-weight:800;letter-spacing:.12em;text-transform:uppercase;color:var(--t-n9)}}
.card-right{{display:flex;align-items:center;gap:6px}}
.badge{{
  font-size:8.5px;padding:2px 7px;border-radius:999px;
  background:color-mix(in srgb,var(--accent,var(--c-acc)) 10%,transparent);
  color:color-mix(in srgb,var(--accent,var(--c-acc)) 90%,white);
  border:1px solid color-mix(in srgb,var(--accent,var(--c-acc)) 25%,transparent);
  letter-spacing:.06em;font-weight:700;
}}
.insight{{
  display:inline-flex;align-items:center;gap:5px;
  font-size:8.5px;color:var(--t-n6);font-weight:600;letter-spacing:.04em;
  padding:2px 8px;border-radius:6px;
  background:rgba(var(--rgb-lift),.03);border:1px solid rgba(var(--rgb-lift),.05);
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:200px;
}}
.insight strong{{color:var(--c-acc-l);font-weight:800}}
/* donut row */
.donut-wrap{{display:flex;align-items:center;gap:20px;height:130px}}
.donut-box{{position:relative;width:130px;height:130px;flex-shrink:0}}
.donut-box canvas{{position:absolute;inset:0;width:100%!important;height:100%!important}}
.donut-legend{{flex:1;display:flex;flex-direction:column;gap:10px}}
.leg-row{{display:flex;align-items:center;gap:8px}}
.leg-dot{{width:8px;height:8px;border-radius:50%;flex-shrink:0}}
.leg-label{{font-size:10px;color:var(--t-n8);font-weight:600;flex:1}}
.leg-val{{font-size:11px;font-weight:800;color:var(--t-n11);font-family:monospace}}
.leg-pct{{font-size:9px;color:var(--t-n5);font-weight:700;margin-left:3px}}
/* chart containers com altura fixa */
.ch-box{{position:relative;width:100%}}
.ch-box canvas{{width:100%!important}}
</style>
</head>
<body>
<div class="dash">

  <div class="card" style="--accent:var(--c-acc)">
    <div class="card-hdr">
      <span class="card-title">{_ico("grafico", 15)}<span style="margin-left:7px">Status dos Chamados</span></span>
      <span class="badge">LIVE</span>
    </div>
    <div class="donut-wrap">
      <div class="donut-box"><canvas id="c0"></canvas></div>
      <div class="donut-legend">
        <div class="leg-row">
          <div class="leg-dot" style="background:var(--f-crit-l)"></div>
          <span class="leg-label">Abertos</span>
          <span class="leg-val">{abertos_ch}</span>
          <span class="leg-pct">({round(abertos_ch/max(abertos_ch+em_atend_ch+concluidos,1)*100)}%)</span>
        </div>
        <div class="leg-row">
          <div class="leg-dot" style="background:var(--f-run)"></div>
          <span class="leg-label">Em atendimento</span>
          <span class="leg-val">{em_atend_ch}</span>
          <span class="leg-pct">({round(em_atend_ch/max(abertos_ch+em_atend_ch+concluidos,1)*100)}%)</span>
        </div>
        <div class="leg-row">
          <div class="leg-dot" style="background:var(--f-ok)"></div>
          <span class="leg-label">Concluídos</span>
          <span class="leg-val">{concluidos}</span>
          <span class="leg-pct">({round(concluidos/max(abertos_ch+em_atend_ch+concluidos,1)*100)}%)</span>
        </div>
      </div>
    </div>
  </div>

  <div class="card" style="--accent:var(--c-crit)">
    <div class="card-hdr">
      <span class="card-title">{_ico("alerta", 15)}<span style="margin-left:7px">Nível de Criticidade</span></span>
      <div class="card-right">
        <span class="insight">Maior: <strong>{top_prio_label} ({top_prio_v})</strong></span>
        <span class="badge" style="background:rgba(var(--rgb-crit),.1);color:var(--c-crit-l);border-color:rgba(var(--rgb-crit),.3)">RISCO</span>
      </div>
    </div>
    <div class="ch-box" style="height:160px"><canvas id="c2"></canvas></div>
  </div>

  <div class="card wide" style="--accent:var(--c-acc)">
    <div class="card-hdr">
      <span class="card-title">{_ico("calendario", 15)}<span style="margin-left:7px">Volume de Chamados por Dia</span></span>
      <div class="card-right">
        <span class="insight">Pico: <strong>{max_dia} ({max_dia_v} tickets)</strong></span>
        <span class="badge">TENDÊNCIA</span>
      </div>
    </div>
    <div class="ch-box" style="height:110px"><canvas id="c3"></canvas></div>
  </div>

  <div class="card" style="--accent:var(--c-acc)">
    <div class="card-hdr">
      <span class="card-title">{_ico("predio", 15)}<span style="margin-left:7px">Incidentes por Setor</span></span>
      <div class="card-right">
        <span class="insight">Líder: <strong>{top_setor} ({top_setor_v})</strong></span>
        <span class="badge">VOLUME</span>
      </div>
    </div>
    <div class="ch-box" style="height:200px"><canvas id="c1"></canvas></div>
  </div>

  <div class="card" style="--accent:var(--c-run-l)">
    <div class="card-hdr">
      <span class="card-title">{_ico("pessoa", 15)}<span style="margin-left:7px">Top Solicitantes</span></span>
      <div class="card-right">
        <span class="insight" style="color:var(--c-run-l)">+ ativo: <strong style="color:var(--c-run-l)">{top_user} ({top_user_v}×)</strong></span>
        <span class="badge" style="background:rgba(var(--rgb-run-l),.1);color:var(--c-run-l);border-color:rgba(var(--rgb-run-l),.3)">RANKING</span>
      </div>
    </div>
    <div class="ch-box" style="height:200px"><canvas id="c4"></canvas></div>
  </div>

  <div class="card wide" style="--accent:var(--c-run-l)">
    <div class="card-hdr">
      <span class="card-title">{_ico("caixa", 15)}<span style="margin-left:7px">Demandas por Categoria de Compras</span></span>
      <span class="badge" style="background:rgba(var(--rgb-run-l),.1);color:var(--c-run-l);border-color:rgba(var(--rgb-run-l),.3)">LOGÍSTICA</span>
    </div>
    <div class="ch-box" style="height:100px"><canvas id="c5"></canvas></div>
  </div>

</div>
<script>
Chart.defaults.color='var(--n8)';
Chart.defaults.font.family="'IBM Plex Mono',monospace";
Chart.defaults.font.size=10.5;
const SL={SL},SV={SV};
const PL={PL},PV={PV};
const TL={TL},TV={TV};
const UL={UL},UV={UV};
const CL={CL},CV={CV};
const STL={STL},STV={STV};
const tt={{
  backgroundColor:'rgba(var(--rgb-n0),0.97)',borderWidth:1,padding:10,
  titleFont:{{weight:'800',size:11}},bodyColor:'var(--n11)',cornerRadius:8,displayColors:false,
}};

/* ── C0: DONUT STATUS ── */
(()=>{{
  const canvas=document.getElementById('c0');
  const ctx=canvas.getContext('2d');
  const CORES=['var(--c-crit-l)','var(--c-run)','var(--c-ok)'];
  const total=STV.reduce((a,b)=>a+b,0);
  new Chart(ctx,{{
    type:'doughnut',
    data:{{labels:STL,datasets:[{{
      data:STV,
      backgroundColor:CORES.map(c=>c+'cc'),
      borderColor:CORES,borderWidth:2,
      hoverBackgroundColor:CORES,hoverOffset:6,
    }}]}},
    options:{{
      cutout:'68%',
      responsive:true,maintainAspectRatio:false,
      animation:{{duration:300,easing:'easeOutQuart'}},
      plugins:{{
        legend:{{display:false}},
        tooltip:{{...tt,borderColor:'var(--ta250)',titleColor:'var(--c-acc)',
          callbacks:{{label:c=>` ${{c.parsed}} (${{total?Math.round(c.parsed/total*100):0}}%)`}}
        }}
      }}
    }},
    plugins:[{{
      afterDraw(chart){{
        const {{ctx:cx,chartArea:{{left,right,top,bottom}}}}=chart;
        const cxc=(left+right)/2,cyc=(top+bottom)/2;
        cx.save();
        cx.textAlign='center';cx.textBaseline='middle';
        cx.fillStyle='var(--n11)';cx.font='bold 22px Geist Mono,monospace';
        cx.fillText(total,cxc,cyc-8);
        cx.fillStyle='var(--n5)';cx.font='700 9px Geist Mono,monospace';
        cx.fillText('TOTAL',cxc,cyc+12);
        cx.restore();
      }}
    }}]
  }});
}})();

/* ── C1: SETOR (horizontal bars) ── */
(()=>{{
  const ctx=document.getElementById('c1').getContext('2d');
  new Chart(ctx,{{
    type:'bar',
    data:{{labels:SL,datasets:[{{
      data:SV,
      backgroundColor:'var(--ta750)',
      hoverBackgroundColor:'var(--c-acc-l)',
      borderRadius:{{topRight:5,bottomRight:5}},borderSkipped:false,barThickness:16,
    }}]}},
    options:{{
      indexAxis:'y',
      responsive:true,maintainAspectRatio:false,
      animation:{{duration:300,easing:'easeOutQuart'}},
      plugins:{{
        legend:{{display:false}},
        tooltip:{{...tt,borderColor:'var(--ta300)',titleColor:'var(--c-acc)',
          callbacks:{{label:c=>` ${{c.parsed.x}} chamados`}}}}
      }},
      scales:{{
        x:{{display:false}},
        y:{{grid:{{display:false}},border:{{display:false}},
          ticks:{{color:'var(--t-n10)',font:{{weight:'600',size:11}},padding:6}}}}
      }}
    }},
    plugins:[{{
      afterDatasetsDraw(chart){{
        const c2=chart.ctx,meta=chart.getDatasetMeta(0);
        meta.data.forEach((bar,i)=>{{
          const v=SV[i];if(!v)return;
          c2.save();c2.fillStyle='rgba(var(--rgb-acc-l),.85)';
          c2.font='bold 11px Geist Mono,monospace';
          c2.textAlign='left';c2.textBaseline='middle';
          c2.fillText(v,bar.x+7,bar.y);c2.restore();
        }});
      }}
    }}]
  }});
}})();

/* ── C2: CRITICIDADE (vertical bars) ── */
(()=>{{
  const ctx=document.getElementById('c2').getContext('2d');
  const CORES=['var(--c-acc)','var(--c-run)','var(--c-run)','var(--c-crit)','var(--c-crit)'];
  new Chart(ctx,{{
    type:'bar',
    data:{{labels:PL,datasets:[{{
      data:PV,
      backgroundColor:CORES.map(c=>c+'bb'),
      borderColor:CORES,borderWidth:0,
      borderRadius:{{topLeft:6,topRight:6}},borderSkipped:'bottom',
      barThickness:32,
    }}]}},
    options:{{
      responsive:true,maintainAspectRatio:false,
      animation:{{duration:300,easing:'easeOutExpo'}},
      plugins:{{
        legend:{{display:false}},
        tooltip:{{...tt,borderColor:'rgba(var(--rgb-crit),.3)',titleColor:'var(--c-crit-l)',
          callbacks:{{label:c=>` ${{c.parsed.y}} chamados`}}}}
      }},
      scales:{{
        x:{{grid:{{display:false}},border:{{display:false}},
          ticks:{{color:'var(--t-n9)',font:{{weight:'800',size:10}}}}}},
        y:{{grid:{{color:'rgba(var(--rgb-lift),.04)',drawTicks:false}},border:{{display:false}},
          ticks:{{color:'var(--t-n5)',padding:4,stepSize:1}},
          min:0}}
      }}
    }},
    plugins:[{{
      afterDatasetsDraw(chart){{
        const c2=chart.ctx,meta=chart.getDatasetMeta(0);
        meta.data.forEach((bar,i)=>{{
          const v=PV[i];if(!v)return;
          c2.save();
          c2.fillStyle='var(--n11)';c2.font='bold 12px Geist Mono,monospace';
          c2.textAlign='center';c2.textBaseline='bottom';
          c2.fillText(v,bar.x,bar.y-4);c2.restore();
        }});
      }}
    }}]
  }});
}})();

/* ── C3: LINHA TEMPORAL ── */
(()=>{{
  const ctx=document.getElementById('c3').getContext('2d');
  const maxV=Math.max(...TV,1);
  new Chart(ctx,{{
    type:'line',
    data:{{labels:TL,datasets:[{{
      data:TV,borderColor:'var(--c-acc)',borderWidth:2.5,
      backgroundColor:'var(--ta120)',fill:true,tension:0.38,
      pointBackgroundColor:TV.map(v=>v===maxV?'var(--c-run)':'var(--c-acc)'),
      pointBorderColor:TV.map(v=>v===maxV?'var(--c-run-d)':'var(--n1)'),
      pointBorderWidth:TV.map(v=>v===maxV?2.5:1.5),
      pointRadius:TV.map(v=>v===maxV?6:3.5),
      pointHoverRadius:7,pointHoverBackgroundColor:'var(--c-acc-l)',
    }}]}},
    options:{{
      responsive:true,maintainAspectRatio:false,
      animation:{{duration:300,easing:'easeOutCubic'}},
      plugins:{{
        legend:{{display:false}},
        tooltip:{{...tt,borderColor:'var(--ta300)',titleColor:'var(--c-acc)',
          callbacks:{{label:c=>` ${{c.parsed.y}} tickets`}}}}
      }},
      scales:{{
        x:{{grid:{{color:'rgba(var(--rgb-lift),.03)',drawTicks:false}},border:{{display:false}},
          ticks:{{color:'var(--t-n8)',maxRotation:30,font:{{size:9.5}}}}}},
        y:{{grid:{{color:'rgba(var(--rgb-lift),.03)',drawTicks:false}},border:{{display:false}},
          ticks:{{color:'var(--t-n5)',padding:6,stepSize:1}},min:0}}
      }}
    }}
  }});
}})();

/* ── C4: TOP USUÁRIOS ── */
(()=>{{
  const ctx=document.getElementById('c4').getContext('2d');
  new Chart(ctx,{{
    type:'bar',
    data:{{labels:UL,datasets:[{{
      data:UV,
      backgroundColor:'rgba(var(--rgb-run),0.75)',
      hoverBackgroundColor:'var(--c-run-l)',
      borderRadius:{{topRight:5,bottomRight:5}},borderSkipped:false,barThickness:15,
    }}]}},
    options:{{
      indexAxis:'y',
      responsive:true,maintainAspectRatio:false,
      animation:{{duration:300,easing:'easeOutQuart'}},
      plugins:{{
        legend:{{display:false}},
        tooltip:{{...tt,borderColor:'rgba(var(--rgb-run-l),.3)',titleColor:'var(--c-run-l)',
          callbacks:{{label:c=>` ${{c.parsed.x}} chamados`}}}}
      }},
      scales:{{
        x:{{display:false}},
        y:{{grid:{{display:false}},border:{{display:false}},
          ticks:{{color:'var(--t-n10)',font:{{weight:'600',size:11}},padding:6}}}}
      }}
    }},
    plugins:[{{
      afterDatasetsDraw(chart){{
        const c2=chart.ctx,meta=chart.getDatasetMeta(0);
        meta.data.forEach((bar,i)=>{{
          c2.save();c2.fillStyle='rgba(var(--rgb-run-l),.9)';
          c2.font='bold 11px Geist Mono,monospace';
          c2.textAlign='left';c2.textBaseline='middle';
          c2.fillText(UV[i],bar.x+7,bar.y);c2.restore();
        }});
      }}
    }}]
  }});
}})();

/* ── C5: CATEGORIAS DE COMPRAS ── */
(()=>{{
  if(!CL.length)return;
  const ctx=document.getElementById('c5').getContext('2d');
  new Chart(ctx,{{
    type:'bar',
    data:{{labels:CL,datasets:[{{
      data:CV,
      backgroundColor:'rgba(var(--rgb-run-d),0.75)',
      hoverBackgroundColor:'var(--c-run-l)',
      borderRadius:{{topRight:5,bottomRight:5}},borderSkipped:false,barThickness:14,
    }}]}},
    options:{{
      indexAxis:'y',
      responsive:true,maintainAspectRatio:false,
      animation:{{duration:300,easing:'easeOutQuart'}},
      plugins:{{
        legend:{{display:false}},
        tooltip:{{...tt,borderColor:'rgba(var(--rgb-run-l),.3)',titleColor:'var(--c-run-l)',
          callbacks:{{label:c=>` ${{c.parsed.x}} pedidos`}}}}
      }},
      scales:{{
        x:{{display:false}},
        y:{{grid:{{display:false}},border:{{display:false}},
          ticks:{{color:'var(--t-n10)',font:{{weight:'600',size:11}},padding:6}}}}
      }}
    }},
    plugins:[{{
      afterDatasetsDraw(chart){{
        const c2=chart.ctx,meta=chart.getDatasetMeta(0);
        meta.data.forEach((bar,i)=>{{
          c2.save();c2.fillStyle='rgba(var(--rgb-run-l),.85)';
          c2.font='bold 11px Geist Mono,monospace';
          c2.textAlign='left';c2.textBaseline='middle';
          c2.fillText(CV[i],bar.x+7,bar.y);c2.restore();
        }});
      }}
    }}]
  }});
}})();
</script>
<script>
(function(){{}})();
</script>
</body></html>"""

                components.html(_resolver_vars(DASH_HTML), height=880, scrolling=True)

        with tab_analytics:
            renderizar_analytics_realtime()
        with tab_chamados:
            if not df_base.empty:
                prioridades_disponiveis = sorted([
                    x for x in df_base['Prioridade'].dropna().unique()
                    if str(x).strip().lower() not in ['nan', 'none', '']
                ], key=lambda p: {'Baixa':0,'Média':1,'Alta':2,'Crítica':3,'Urgente':4}.get(p, 99))
                status_disponiveis = ["Aberto", "Em Atendimento", "Concluído"]
                setores_disponiveis = sorted([
                    x for x in df_base['Setor'].dropna().unique()
                    if str(x).strip().lower() not in ['nan', 'none', '']
                ])

                if st.session_state.get("_reset_filters_ch"):
                    for k in ["f_prio_ch", "f_status_ch", "f_setor_ch", "f_txt_mgr", "f_data_ini_ch", "f_data_fim_ch"]:
                        if k in st.session_state:
                            del st.session_state[k]
                    st.session_state["_reset_filters_ch"] = False
                    st.rerun(scope="fragment")

                st.caption("Filtros")
                with st.container(border=True, key="cartao_08"):
                    _fa, _fb, _fc, _fd, _fe, _ff = st.columns([1.4, 1.2, 1.8, 1, 1, 0.7])

                    with _fa:
                        filtro_prio = st.multiselect(
                            "Urgência",
                            options=prioridades_disponiveis,
                            default=[],
                            key="f_prio_ch",
                            placeholder="Todas",
                        )
                        if not filtro_prio:
                            filtro_prio = prioridades_disponiveis

                    with _fb:
                        filtro_status = st.multiselect(
                            "Status",
                            options=status_disponiveis,
                            default=[],
                            key="f_status_ch",
                            placeholder="Todos",
                        )
                        if not filtro_status:
                            filtro_status = status_disponiveis

                    with _fc:
                        filtro_setor = st.multiselect(
                            "Setor / Unidade",
                            options=setores_disponiveis,
                            default=[],
                            key="f_setor_ch",
                            placeholder="Todos",
                        )
                        if not filtro_setor:
                            filtro_setor = setores_disponiveis

                    with _fd:
                        filtro_data_ini_ch = st.date_input("De", value=None, key="f_data_ini_ch", format="DD/MM/YYYY")

                    with _fe:
                        filtro_data_fim_ch = st.date_input("Até", value=None, key="f_data_fim_ch", format="DD/MM/YYYY")

                    with _ff:
                        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
                        if st.button("Limpar", key="btn_reset_filters", use_container_width=True, help="Limpar filtros"):
                            st.session_state["_reset_filters_ch"] = True
                            st.rerun(scope="fragment")

                busca_texto = st.text_input(
                    "Busca livre",
                    key="f_txt_mgr",
                    placeholder="Nome, protocolo, setor ou palavra-chave…",
                    label_visibility="visible",
                )

                df_filtrado = df_base[
                    (df_base['Prioridade'].fillna('').isin(filtro_prio) | (df_base['Prioridade'].fillna('') == '')) &
                    df_base['Status'].fillna('').str.strip().isin(filtro_status) &
                    (df_base['Setor'].fillna('').isin(filtro_setor) | (df_base['Setor'].fillna('') == ''))
                ]

                if filtro_data_ini_ch or filtro_data_fim_ch:
                    _dts_ch = pd.to_datetime(df_filtrado['Data'], format="%d/%m/%Y %H:%M", errors="coerce")
                    if filtro_data_ini_ch:
                        df_filtrado = df_filtrado[_dts_ch >= pd.Timestamp(filtro_data_ini_ch)]
                        _dts_ch = pd.to_datetime(df_filtrado['Data'], format="%d/%m/%Y %H:%M", errors="coerce")
                    if filtro_data_fim_ch:
                        df_filtrado = df_filtrado[_dts_ch <= pd.Timestamp(filtro_data_fim_ch) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)]

                if busca_texto:
                    df_filtrado = df_filtrado[
                        df_filtrado['Usuário'].str.contains(busca_texto, case=False, na=False) |
                        df_filtrado['Protocolo'].str.contains(busca_texto, case=False, na=False) |
                        df_filtrado['Descrição'].str.contains(busca_texto, case=False, na=False) |
                        df_filtrado['Setor'].str.contains(busca_texto, case=False, na=False)
                    ]

                df_ativa = ordenar_por_prioridade_e_data(
                    df_filtrado[df_filtrado['Status'].str.strip().isin(["Aberto", "Em Atendimento"])],
                    "Prioridade"
                )
                df_historico = df_filtrado[df_filtrado['Status'].str.strip() == "Concluído"].copy()
                if not df_historico.empty:
                    _quando = df_historico.get("Atualizado em", pd.Series("", index=df_historico.index)).astype(str).str.strip()
                    _quando = _quando.where(_quando != "", df_historico["Data"].astype(str))
                    df_historico = (
                        df_historico
                        .assign(_ord=pd.to_datetime(_quando, format="%d/%m/%Y %H:%M", errors="coerce"))
                        .sort_values("_ord", ascending=False, na_position="last")
                        .drop(columns="_ord")
                    )

                c_space, c_export = st.columns([3, 1])
                with c_export:
                    st.download_button("Exportar PDF filtrado", data=gerar_pdf(df_filtrado), file_name="chamados_ti_filtrado.pdf", use_container_width=True, key="dl_pdf_ch")

                abertos_fil = len(df_filtrado[df_filtrado["Status"] == "Aberto"])
                atend_fil = len(df_filtrado[df_filtrado["Status"] == "Em Atendimento"])
                concl_fil = len(df_filtrado[df_filtrado["Status"] == "Concluído"])
                st.markdown(f"""
                <div class="mini-kanban animate-fade-in">
                    <div class="kanban-lane"><strong>🔴 Aberto: {abertos_fil}</strong><span>Entradas aguardando primeiro despacho.</span></div>
                    <div class="kanban-lane"><strong>🟡 Em atendimento: {atend_fil}</strong><span>Itens já assumidos pela equipe.</span></div>
                    <div class="kanban-lane"><strong>🟢 Concluído: {concl_fil}</strong><span>Registros fechados no filtro atual.</span></div>
                </div>
            """, unsafe_allow_html=True)

                st.markdown("#### Fila de Trabalho Ativa")
                if not df_ativa.empty:
                    _evt_fila_ch = st.dataframe(
                        _tabela_fila_chamados(df_ativa),
                        use_container_width=True,
                        hide_index=True,
                        selection_mode="single-row",
                        on_select="rerun",
                        key="grid_fila_chamados",
                        column_config={"Descrição": st.column_config.TextColumn(width="large"),
                                       "Aberto em": _COL_ABERTO},
                    )
                    _idx_sel_ch, _row_sel_ch = _linha_selecionada(_evt_fila_ch, df_ativa)
                    if _idx_sel_ch is not None:
                        if st.button("Despachar chamado selecionado", key="btn_despachar_sel_ch",
                                     use_container_width=True, type="primary"):
                            modal_responder_chamado(_idx_sel_ch, _row_sel_ch)
                    else:
                        st.caption("Clique numa linha da tabela para selecionar o chamado e despachá-lo.")
                else:
                    st.markdown('<div class="empty-state"><strong>Fila de chamados limpa</strong><span>Nenhum chamado aberto ou em atendimento no filtro atual.</span></div>', unsafe_allow_html=True)

                st.write("")
            
                with st.expander("Histórico e Edição de Chamados Concluídos", expanded=False):
                    if not df_historico.empty:
                        st.markdown(f"""
                        <div class="history-panel-head">
                            <div class="history-panel-title">
                                <span class="history-title-dot"></span>
                                <span>Base de Concluídos</span>
                            </div>
                            <span class="history-count">{len(df_historico)} registros</span>
                        </div>
                    """, unsafe_allow_html=True)

                        @st.fragment
                        def _lista_concluidos(_df):
                            _cb, _cq = st.columns([4, 1], vertical_alignment="bottom")
                            with _cb:
                                _busca = st.text_input(
                                    "Buscar", key="hcc_busca", label_visibility="collapsed",
                                    placeholder="Buscar por protocolo, pessoa, setor ou categoria…",
                                )
                            with _cq:
                                _so_crit = st.toggle("Só críticos", key="hcc_crit")

                            _v = _df
                            if _so_crit:
                                _v = _v[_v["Prioridade"].isin(["Crítica", "Urgente", "Alta"])]
                            if _busca:
                                _t = normalizar_busca(_busca)
                                _mask = False
                                for _c in ("Protocolo", "Usuário", "Setor", "Categoria", "Descrição"):
                                    if _c in _v.columns:
                                        _m = _v[_c].astype(str).map(lambda x: _t in normalizar_busca(x))
                                        _mask = _m if _mask is False else (_mask | _m)
                                if _mask is not False:
                                    _v = _v[_mask]

                            _n = len(_v)
                            if _n == 0:
                                st.markdown('<div class="hcc-nada">Nenhum chamado corresponde ao filtro.</div>',
                                            unsafe_allow_html=True)
                                return

                            _qtd = st.session_state.get("hcc_qtd", 10)
                            _aberto = st.session_state.get("hcc_aberto")

                            for _i, _r in _v.head(_qtd).iterrows():
                                _p = str(_r.get("Prioridade", "")).strip().lower()
                                _pc = ("crit" if ("cr" in _p or "urg" in _p)
                                       else "alta" if "alt" in _p
                                       else "media" if _p.startswith("m") else "baixa")
                                _proto = escapar_html(primeiro_valor(_r, "Protocolo", padrao=f"#{formatar_id(_i)}"), "—")
                                _quando = str(_r.get("Atualizado em", "") or _r.get("Data", "")).strip()
                                _desc = str(_r.get("Descrição", "")).strip() or "—"
                                _desc = escapar_html(_desc[:110] + "…" if len(_desc) > 110 else _desc, "—")
                                _eh_aberto = (_aberto == _i)

                                _c1, _c2 = st.columns([9, 1.4], vertical_alignment="center")
                                with _c1:
                                    st.markdown(
                                        f'<div class="hcc-row{" hcc-on" if _eh_aberto else ""}">'
                                        f'<div class="hcc-stripe hcc-{_pc}"></div>'
                                        f'<div class="hcc-quem"><div class="hcc-nome">{escapar_html(_r.get("Usuário",""),"—")}</div>'
                                        f'<div class="hcc-proto">{_proto}</div></div>'
                                        f'<div><div class="hcc-desc">{_desc}</div>'
                                        f'<div class="hcc-sub">{escapar_html(extrair_categoria(_r),"—")} &middot; '
                                        f'{escapar_html(_r.get("Setor",""),"—")}</div></div>'
                                        f'<div><span class="hcc-chip hcc-c-{_pc}">{escapar_html(_r.get("Prioridade","—"),"—")}</span></div>'
                                        f'<div class="hcc-quando">{escapar_html(_quando,"—")}</div>'
                                        f'</div>',
                                        unsafe_allow_html=True,
                                    )
                                with _c2:
                                    if st.button("Fechar" if _eh_aberto else "Abrir",
                                                 key=f"hcc_btn_{_i}", use_container_width=True,
                                                 type="primary" if _eh_aberto else "secondary"):
                                        st.session_state["hcc_aberto"] = None if _eh_aberto else _i
                                        st.rerun(scope="fragment")

                                if _eh_aberto:
                                    _detalhe_chamado_concluido(_r)
                                    _ce1, _ce2 = st.columns([3, 1])
                                    with _ce2:
                                        if st.button("Editar chamado", key=f"hcc_edit_{_i}",
                                                     use_container_width=True, type="primary"):
                                            modal_responder_chamado(_i, _r)
                                    _ax = str(_r.get("Anexo", "") or "").strip()
                                    _ar = str(_r.get("Anexo Resposta", "") or "").strip()
                                    if _ax and _ax.lower() not in ("nenhum", "nan", "none"):
                                        st.caption("Anexo do solicitante")
                                        renderizar_anexo(_ax, key=f"hcc_anx_{_i}")
                                    if _ar and _ar.lower() not in ("nenhum", "nan", "none"):
                                        st.caption("Anexo da resposta")
                                        renderizar_anexo(_ar, key=f"hcc_anr_{_i}")

                            if _n > _qtd:
                                if st.button(f"Mostrar mais ({_n - _qtd} restantes)",
                                             key="hcc_mais", use_container_width=True):
                                    st.session_state["hcc_qtd"] = _qtd + 10
                                    st.rerun(scope="fragment")
                            elif _qtd > 10:
                                if st.button("Recolher lista", key="hcc_menos", use_container_width=True):
                                    st.session_state["hcc_qtd"] = 10
                                    st.rerun(scope="fragment")

                        _lista_concluidos(df_historico)
                    else:
                        st.markdown('<div class="empty-state"><strong>Sem histórico concluído</strong><span>Nenhum chamado concluído dentro do filtro atual.</span></div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="empty-state"><strong>Nenhum ticket registrado</strong><span>A base de chamados ainda não recebeu solicitações.</span></div>', unsafe_allow_html=True)

        with tab_compras:
            if not df_compras.empty:
                urgencias_disponiveis_cp = sorted([
                    x for x in df_compras['Urgência'].dropna().unique()
                    if str(x).strip().lower() not in ['nan', 'none', '']
                ], key=lambda u: {'Sem pressa':0,'Moderada':1,'Alta':2,'Crítica':3}.get(u, 99))
                status_disponiveis_cp = ["Pendente", "Em Análise", "Aprovado", "Recusado", "Concluído"]
                setores_disponiveis_cp = sorted([
                    x for x in df_compras['Setor'].dropna().unique()
                    if str(x).strip().lower() not in ['nan', 'none', '']
                ])
                categorias_disponiveis_cp = sorted([
                    x for x in df_compras['Categoria'].dropna().unique()
                    if str(x).strip().lower() not in ['nan', 'none', '']
                ])

                if st.session_state.get("_reset_filters_cp"):
                    for k in ["f_urg_cp", "f_status_cp", "f_setor_cp", "f_cat_cp", "f_data_ini_cp", "f_data_fim_cp", "f_txt_cp"]:
                        if k in st.session_state:
                            del st.session_state[k]
                    st.session_state["_reset_filters_cp"] = False
                    st.rerun(scope="fragment")

                st.caption("Filtros")
                with st.container(border=True, key="cartao_09"):
                    _ca, _cb, _cc, _cd, _ce, _cf, _cg = st.columns([1.2, 1.4, 1.4, 1.4, 0.9, 0.9, 0.6])

                    with _ca:
                        filtro_urg_cp = st.multiselect(
                            "Urgência",
                            options=urgencias_disponiveis_cp,
                            default=[],
                            key="f_urg_cp",
                            placeholder="Todas",
                        )
                        if not filtro_urg_cp:
                            filtro_urg_cp = urgencias_disponiveis_cp

                    with _cb:
                        filtro_status_cp = st.multiselect(
                            "Status",
                            options=status_disponiveis_cp,
                            default=[],
                            key="f_status_cp",
                            placeholder="Todos",
                        )
                        if not filtro_status_cp:
                            filtro_status_cp = status_disponiveis_cp

                    with _cc:
                        filtro_setor_cp = st.multiselect(
                            "Setor",
                            options=setores_disponiveis_cp,
                            default=[],
                            key="f_setor_cp",
                            placeholder="Todos",
                        )
                        if not filtro_setor_cp:
                            filtro_setor_cp = setores_disponiveis_cp

                    with _cd:
                        filtro_cat_cp = st.multiselect(
                            "Categoria",
                            options=categorias_disponiveis_cp,
                            default=[],
                            key="f_cat_cp",
                            placeholder="Todas",
                        )
                        if not filtro_cat_cp:
                            filtro_cat_cp = categorias_disponiveis_cp

                    with _ce:
                        filtro_data_ini_cp = st.date_input("De", value=None, key="f_data_ini_cp", format="DD/MM/YYYY")

                    with _cf:
                        filtro_data_fim_cp = st.date_input("Até", value=None, key="f_data_fim_cp", format="DD/MM/YYYY")

                    with _cg:
                        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
                        if st.button("Limpar", key="btn_reset_filters_cp", use_container_width=True, help="Limpar filtros"):
                            st.session_state["_reset_filters_cp"] = True
                            st.rerun(scope="fragment")

                busca_texto_cp = st.text_input(
                    "Busca livre",
                    key="f_txt_cp",
                    placeholder="Item, usuário, setor ou protocolo…",
                    label_visibility="visible",
                )

                _urg_mask = (
                    df_compras['Urgência'].fillna('').isin(filtro_urg_cp) |
                    (df_compras['Urgência'].fillna('') == '')
                )
                df_compras_filtrado = df_compras[
                    _urg_mask &
                    df_compras['Status'].fillna('').str.strip().isin(filtro_status_cp) &
                    (df_compras['Setor'].fillna('').isin(filtro_setor_cp) | (df_compras['Setor'].fillna('') == '')) &
                    (df_compras['Categoria'].fillna('').isin(filtro_cat_cp) | (df_compras['Categoria'].fillna('') == ''))
                ]

                if filtro_data_ini_cp or filtro_data_fim_cp:
                    _dts_cp = pd.to_datetime(df_compras_filtrado['Data'], format="%d/%m/%Y %H:%M", errors="coerce")
                    if filtro_data_ini_cp:
                        df_compras_filtrado = df_compras_filtrado[_dts_cp >= pd.Timestamp(filtro_data_ini_cp)]
                        _dts_cp = pd.to_datetime(df_compras_filtrado['Data'], format="%d/%m/%Y %H:%M", errors="coerce")
                    if filtro_data_fim_cp:
                        df_compras_filtrado = df_compras_filtrado[_dts_cp <= pd.Timestamp(filtro_data_fim_cp) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)]

                if busca_texto_cp:
                    df_compras_filtrado = df_compras_filtrado[
                        df_compras_filtrado['Usuário'].str.contains(busca_texto_cp, case=False, na=False) |
                        df_compras_filtrado['Item'].str.contains(busca_texto_cp, case=False, na=False) |
                        df_compras_filtrado['Protocolo'].str.contains(busca_texto_cp, case=False, na=False) |
                        df_compras_filtrado['Setor'].str.contains(busca_texto_cp, case=False, na=False) |
                        df_compras_filtrado['Justificativa'].str.contains(busca_texto_cp, case=False, na=False)
                    ]

                df_comp_ativa = ordenar_por_prioridade_e_data(
                    df_compras_filtrado[df_compras_filtrado['Status'].isin(["Pendente", "Em Análise"])],
                    "Urgência"
                )
                df_comp_historico = df_compras_filtrado[df_compras_filtrado['Status'].isin(["Aprovado", "Recusado", "Concluído"])]

                c_space_c, c_export_c = st.columns([3, 1])
                with c_export_c:
                    st.download_button("Exportar PDF Compras", data=gerar_pdf_compras(df_compras_filtrado), file_name="compras_ti.pdf", use_container_width=True, key="dl_pdf_cp")

                _cp_pend_fil  = len(df_compras_filtrado[df_compras_filtrado["Status"] == "Pendente"])
                _cp_anal_fil  = len(df_compras_filtrado[df_compras_filtrado["Status"] == "Em Análise"])
                _cp_aprov_fil = len(df_compras_filtrado[df_compras_filtrado["Status"] == "Aprovado"])
                _cp_recus_fil = len(df_compras_filtrado[df_compras_filtrado["Status"] == "Recusado"])
                _cp_conc_fil  = len(df_compras_filtrado[df_compras_filtrado["Status"] == "Concluído"])
                st.markdown(f"""
                <div class="mini-kanban animate-fade-in">
                    <div class="kanban-lane"><strong>⏳ Pendente: {_cp_pend_fil}</strong><span>Aguardando análise de TI.</span></div>
                    <div class="kanban-lane"><strong>🔍 Em Análise: {_cp_anal_fil}</strong><span>Em avaliação técnica ou orçamentária.</span></div>
                    <div class="kanban-lane"><strong>✅ Aprovado: {_cp_aprov_fil}</strong><span>Liberadas para aquisição.</span></div>
                    <div class="kanban-lane"><strong>❌ Recusado: {_cp_recus_fil}</strong><span>Negadas no filtro atual.</span></div>
                    <div class="kanban-lane"><strong>🟢 Concluído: {_cp_conc_fil}</strong><span>Finalizadas e entregues.</span></div>
                </div>
            """, unsafe_allow_html=True)

                st.markdown("#### Requisições de Compra Pendentes")
                if not df_comp_ativa.empty:
                    _evt_fila_cp = st.dataframe(
                        _tabela_fila_compras(df_comp_ativa),
                        use_container_width=True,
                        hide_index=True,
                        selection_mode="single-row",
                        on_select="rerun",
                        key="grid_fila_compras",
                        column_config={"Justificativa": st.column_config.TextColumn(width="large"),
                                       "Aberto em": _COL_ABERTO},
                    )
                    _idx_sel_cp, _row_sel_cp = _linha_selecionada(_evt_fila_cp, df_comp_ativa)
                    if _idx_sel_cp is not None:
                        if st.button("Analisar requisição selecionada", key="btn_analisar_sel_cp",
                                     use_container_width=True, type="primary"):
                            modal_responder_compra(_idx_sel_cp, _row_sel_cp)
                    else:
                        st.caption("Clique numa linha da tabela para selecionar a requisição e analisá-la.")
                else:
                    st.markdown('<div class="empty-state"><strong>Compras auditadas</strong><span>Não há solicitações pendentes ou em análise no momento.</span></div>', unsafe_allow_html=True)

                st.write("")

                with st.expander("Ver Arquivo de Solicitações Finalizadas", expanded=False):
                    if not df_comp_historico.empty:
                        st.markdown(f"""
                        <div class="history-panel-head">
                            <div class="history-panel-title">
                                <span class="history-title-dot" style="background:var(--f-acc);box-shadow: none;"></span>
                                <span>Arquivo de Compras Finalizadas</span>
                            </div>
                            <span class="history-count">{len(df_comp_historico)} registros</span>
                        </div>
                    """, unsafe_allow_html=True)

                        _evt_hist_cp = st.dataframe(
                            _tabela_historico_compras(df_comp_historico),
                            use_container_width=True,
                            hide_index=True,
                            selection_mode="single-row",
                            on_select="rerun",
                            key="grid_historico_compras",
                            column_config={"Aberto em": _COL_ABERTO},
                        )
                        _idx_sel_hcp, _row_sel_hcp = _linha_selecionada(_evt_hist_cp, df_comp_historico)
                        if _idx_sel_hcp is not None:
                            if st.button("Rever requisição selecionada", key="btn_rever_sel_cp",
                                         use_container_width=True):
                                modal_responder_compra(_idx_sel_hcp, _row_sel_hcp)
                    else:
                        st.markdown('<div class="empty-state"><strong>Sem arquivo finalizado</strong><span>Nenhuma requisição aprovada, recusada ou concluída até o momento.</span></div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="empty-state"><strong>Nenhuma requisição registrada</strong><span>A base de compras ainda não recebeu solicitações.</span></div>', unsafe_allow_html=True)

    _render_gestor()
elif pagina == "relatorios":

    @st.fragment
    def _render_relatorios():
        import io, zipfile, os
        from datetime import datetime, date
        import calendar
        import pandas as pd

        _MAX_TENT_R   = 5
        _LOCKOUT_R    = 300
        _SESSION_TTL_R = 3600

        if "autenticado_relatorios"    not in st.session_state:
            st.session_state.autenticado_relatorios  = False
        if "relat_auth_ts"             not in st.session_state:
            st.session_state.relat_auth_ts           = 0.0
        if "relat_tentativas"          not in st.session_state:
            st.session_state.relat_tentativas        = 0
        if "relat_lockout_ate"         not in st.session_state:
            st.session_state.relat_lockout_ate       = 0.0

        _ts_r = time.time()
        if st.session_state.autenticado_relatorios:
            if (_ts_r - st.session_state.relat_auth_ts) > _SESSION_TTL_R:
                st.session_state.autenticado_relatorios = False
                st.session_state.relat_auth_ts          = 0.0

        if not st.session_state.autenticado_relatorios:
            _em_lock_r   = _ts_r < st.session_state.relat_lockout_ate
            _seg_rest_r  = max(0, int(st.session_state.relat_lockout_ate - _ts_r))
            _tent_rest_r = max(0, _MAX_TENT_R - st.session_state.relat_tentativas)

            st.markdown("""
<style>
        @keyframes _scan_r {
            0%   { background-position: -200% 0; }
            100% { background-position:  200% 0; }
        }
        @keyframes _blink_r { 0%,100%{opacity:1} 50%{opacity:.15} }

        /* Estreita o bloco central para ser um card compacto */
        .auth-rel-wrap {
            max-width: 400px;
            margin: 6vh auto 0;
        }
        /* Card topo */
        .tr-card {
            position: relative; overflow: hidden;
            background: rgba(var(--rgb-n0),.97);
            border: 1px solid rgba(var(--rgb-run-l),.22);
            border-radius: 12px 12px 0 0;
            font-family: 'IBM Plex Mono', ui-monospace, monospace;
            box-shadow: 0 28px 70px var(--shc-1);
        }
        .tr-card.locked {
            border-color: rgba(var(--rgb-crit-l),.35) !important;
        }
        .tr-card::before {
            content:''; position:absolute; top:0; left:0; right:0; height:1px;
            background: linear-gradient(90deg,transparent,rgba(var(--rgb-run-l),.7),transparent);
            background-size: 200% 1px; animation: none;
        }
        .tr-card.locked::before {
            background: linear-gradient(90deg,transparent,rgba(var(--rgb-crit-l),.8),transparent);
            background-size: 200% 1px;
        }
        .tr-bar {
            display:flex; align-items:center; gap:7px; padding:9px 16px;
            background:rgba(var(--rgb-run-l),.04); border-bottom:1px solid rgba(var(--rgb-run-l),.08);
        }
        .tr-bar.locked { background:rgba(var(--rgb-crit-l),.04); border-bottom-color:rgba(var(--rgb-crit-l),.08); }
        .tr-dot { width:8px; height:8px; border-radius:50%; }
        .tr-bar-title {
            flex:1; text-align:center;
            font-size:.50rem; letter-spacing:.22em; text-transform:uppercase;
            color:rgba(var(--rgb-c-run-l),.30); margin:0;
        }
        .tr-bar.locked .tr-bar-title { color:rgba(var(--rgb-c-crit-l),.30); }
        .tr-body { padding:22px 24px 18px; }
        .tr-icon-ring {
            width:52px; height:52px; border-radius:13px;
            margin:0 auto 14px;
            display:flex; align-items:center; justify-content:center;
            background:linear-gradient(135deg,rgba(var(--rgb-run-l),.15),var(--ta080));
            border:1px solid rgba(var(--rgb-run-l),.22);
            font-size:1.5rem;
        }
        .tr-icon-ring.locked {
            background:linear-gradient(135deg,rgba(var(--rgb-crit-l),.15),rgba(var(--rgb-run),.06));
            border-color:rgba(var(--rgb-crit-l),.30);
        }
        .tr-badge {
            display:inline-flex; align-items:center; gap:6px;
            font-size:.50rem; letter-spacing:.16em; text-transform:uppercase;
            color:rgba(var(--rgb-c-run-l),.50); margin-bottom:8px;
        }
        .tr-badge.locked { color:rgba(var(--rgb-c-crit-l),.55); }
        .tr-badge-dot { width:5px; height:5px; border-radius:50%; background:var(--f-run-l); flex-shrink:0; animation: none; }
        .tr-badge-dot.locked { background:var(--f-crit-l); }
        .tr-title { font-size:.96rem; font-weight:800; color:var(--c-acc-xl); letter-spacing:-.01em; margin-bottom:4px; text-align:center; }
        .tr-sub   { font-size:.62rem; color:rgba(var(--rgb-t-n9),.62); text-align:center; }
        .tr-sep   { height:1px; margin:16px 0 14px; background:linear-gradient(90deg,transparent,rgba(var(--rgb-run-l),.20),transparent); }
        .tr-sep.locked { background:linear-gradient(90deg,transparent,rgba(var(--rgb-crit-l),.18),transparent); }
        .tr-attempts { display:flex; align-items:center; gap:6px; margin-bottom:10px; justify-content:center; }
        .tr-att-dot { width:8px; height:8px; border-radius:50%; }
        .tr-att-dot.used  { background:var(--f-crit-l); box-shadow: none; }
        .tr-att-dot.avail { background:rgba(var(--rgb-run-l),.25); border:1px solid rgba(var(--rgb-run-l),.18); }
        .tr-lockout-bar {
            padding:11px 13px; border:1px solid rgba(var(--rgb-crit-l),.28); border-radius:8px;
            background:rgba(var(--rgb-crit-l),.06);
        }
        .tr-lockout-title { font-size:.62rem; font-weight:800; color:var(--c-crit-l); letter-spacing:.08em; text-transform:uppercase; margin-bottom:3px; }
        .tr-lockout-msg   { font-size:.60rem; color:rgba(var(--rgb-c-crit-l),.65); }
        .tr-warn-token {
            padding:9px 12px; margin-bottom:10px;
            border:1px solid rgba(var(--rgb-run),.24); border-radius:7px;
            background:rgba(var(--rgb-run),.05);
            font-size:.60rem; color:rgba(var(--rgb-c-run),.75); line-height:1.45;
        }
        .tr-warn-token strong { color:var(--c-run); }

        /* Form Streamlit — integrado embaixo do card */
        div[data-testid="stForm"][data-form-key="flr"] {
            background:rgba(var(--rgb-n0),.97) !important;
            border:1px solid rgba(var(--rgb-run-l),.22) !important;
            border-top:none !important;
            border-radius:0 0 12px 12px !important;
            padding:4px 24px 18px !important;
            margin:0 !important;
            box-shadow: 0 28px 70px var(--shc-1) !important;
        }
        div[data-testid="stForm"][data-form-key="flr"] input[type="password"] {
            background:rgba(var(--rgb-run-l),.04) !important;
            border:1px solid rgba(var(--rgb-run-l),.18) !important;
            border-radius:7px !important; color:var(--c-run-l) !important;
            font-family:'IBM Plex Mono',monospace !important;
            font-size:.78rem !important; letter-spacing:.06em !important;
            height:40px !important; padding:0 14px !important;
        }
        div[data-testid="stForm"][data-form-key="flr"] input[type="password"]:focus {
            border-color:rgba(var(--rgb-run-l),.55) !important;
            box-shadow: 0 0 0 3px rgba(var(--rgb-run-l),.10) !important;
        }
        div[data-testid="stForm"][data-form-key="flr"] .stFormSubmitButton button {
            width:100% !important; height:34px !important;
            background:transparent !important;
            border:1px solid rgba(var(--rgb-run-l),.18) !important;
            border-radius:6px !important; color:rgba(var(--rgb-c-run-l),.55) !important;
            font-family:'IBM Plex Mono',monospace !important;
            font-size:.56rem !important; font-weight:700 !important;
            letter-spacing:.18em !important; text-transform:uppercase !important;
            padding:0 !important; margin-top:10px !important;
            transition:all .18s !important;
        }
        div[data-testid="stForm"][data-form-key="flr"] .stFormSubmitButton button:hover {
            background:rgba(var(--rgb-run-l),.08) !important;
            border-color:rgba(var(--rgb-run-l),.42) !important;
            color:var(--c-run-l) !important;
            box-shadow: none;
        }
        .tr-footer-r {
            display:flex; justify-content:space-between; align-items:center;
            padding:10px 0 0;
            font-family:'IBM Plex Mono',monospace; font-size:.46rem; letter-spacing:.10em;
        }
        </style>
        """, unsafe_allow_html=True)

            _rdots = '<div class="tr-attempts">'
            for i in range(_MAX_TENT_R):
                _cls = "used" if i < st.session_state.relat_tentativas else "avail"
                _rdots += f'<div class="tr-att-dot {_cls}"></div>'
            _rdots += '</div>'

            _locked_r = "locked" if _em_lock_r else ""

            if _em_lock_r:
                _rcard_extra = f"""
            <div class="tr-sep {_locked_r}"></div>
            <div class="tr-lockout-bar">
              <div class="tr-lockout-title">Acesso bloqueado</div>
              <div class="tr-lockout-msg">Aguarde <strong style="color:var(--c-crit-l);">{_seg_rest_r}s</strong> para tentar novamente.</div>
            </div>"""
            elif token_padrao_ativo():
                _rcard_extra = """
            <div class="tr-sep"></div>
            <div class="tr-warn-token">
              ⚠️ <strong>Token padrão ativo!</strong><br>
              Configure <code>PORTAL_SUPORTE_ADMIN_TOKEN</code> em <code>secrets.toml</code>.
            </div>"""
            else:
                _rcard_extra = f'<div class="tr-sep"></div>{_rdots}'

            _ca, _cb, _cc = st.columns([1, 2.4, 1])
            with _cb:
                st.markdown(f"""
            <div class="tr-card {_locked_r}">
              <div class="tr-bar {_locked_r}">
                <div class="tr-dot" style="background:var(--f-crit-l);opacity:.65;"></div>
                <div class="tr-dot" style="background:var(--f-run);opacity:.65;"></div>
                <div class="tr-dot" style="background:var(--f-ok);opacity:.65;"></div>
                <span class="tr-bar-title">central · relatórios · exportação</span>
              </div>
              <div class="tr-body">
                <div class="tr-icon-ring {_locked_r}">{_ico("cadeado", 26) if _em_lock_r else _ico("grafico", 26)}</div>
                <div style="text-align:center;">
                  <div class="tr-badge {_locked_r}">
                    <div class="tr-badge-dot {_locked_r}"></div>
                    {"acesso bloqueado" if _em_lock_r else "acesso restrito · relatórios"}
                  </div>
                </div>
                <div class="tr-title">{"Conta Bloqueada" if _em_lock_r else "Autenticação Requerida"}</div>
                <div class="tr-sub">{"Muitas tentativas inválidas." if _em_lock_r else "Insira o token para continuar."}</div>
                {_rcard_extra}
              </div>
            </div>
            """, unsafe_allow_html=True)

                if not _em_lock_r:
                    with st.form("flr", clear_on_submit=True):
                        st.text_input("t", type="password",
                                      placeholder="▸ token de acesso",
                                      key="pwd_relatorios", label_visibility="collapsed")
                        _btn_r = st.form_submit_button("Autenticar", use_container_width=True)
                        if _btn_r:
                            if validar_token_admin(st.session_state.get("pwd_relatorios", "")):
                                st.session_state.autenticado_relatorios = True
                                st.session_state.relat_auth_ts          = time.time()
                                st.session_state.relat_tentativas       = 0
                                st.session_state.relat_lockout_ate      = 0.0
                                st.rerun(scope="app")
                            else:
                                st.session_state.relat_tentativas += 1
                                if st.session_state.relat_tentativas >= _MAX_TENT_R:
                                    st.session_state.relat_lockout_ate = time.time() + _LOCKOUT_R
                                    st.markdown(f"""<div style="font-family:'IBM Plex Mono',monospace;font-size:.58rem;
                                    color:rgba(var(--rgb-c-crit-l),.9);text-align:center;padding:4px 0 0;">
                                    Bloqueado por {_LOCKOUT_R//60} min.</div>""", unsafe_allow_html=True)
                                else:
                                    _rest_r = _MAX_TENT_R - st.session_state.relat_tentativas
                                    st.markdown(f"""<div style="font-family:'IBM Plex Mono',monospace;font-size:.58rem;
                                    color:rgba(var(--rgb-c-crit-l),.85);text-align:center;padding:4px 0 0;">
                                    ✗ &nbsp;Token inválido — {_rest_r} tentativa(s) restante(s)</div>""", unsafe_allow_html=True)

                st.markdown(f"""
            <div class="tr-footer-r">
              <span style="color:rgba(var(--rgb-c-run-l),.18);">TLM Service Desk · v3.8</span>
              <span style="color:rgba(var(--rgb-c-ok),.35);display:flex;align-items:center;gap:5px;">
                <span style="width:4px;height:4px;border-radius:50%;background:var(--f-ok);
                  box-shadow: none;display:inline-block;"></span>
                hmac · sha256
              </span>
            </div>
            """, unsafe_allow_html=True)

            st.stop()


        _sessao_rest_r  = max(0, int(_SESSION_TTL_R - (_ts_r - st.session_state.relat_auth_ts)))
        _sessao_min_r   = _sessao_rest_r // 60
        _agora_str_r    = datetime.now().strftime("%d/%m/%Y  %H:%M")

        st.markdown("""
<style>
    /* ── RELATÓRIOS: HERO BAR ── */
    .rel-hero {
        position: relative;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 16px;
        padding: 20px 26px;
        margin-bottom: 22px;
        background: linear-gradient(135deg, rgba(var(--rgb-sup-a),.98), rgba(var(--rgb-sup-b),.95));
        border: 1px solid rgba(var(--rgb-run-l),.18);
        border-radius: 14px;
        overflow: hidden;
        box-shadow: 0 2px 40px var(--shc-2), 0 0 0 1px rgba(var(--rgb-run-l),.05);
    }
    .rel-hero::before {
        content: "";
        position: absolute;
        top: 0; left: 0; right: 0; height: 2px;
        background: linear-gradient(90deg, var(--f-run-l), var(--f-acc), var(--f-run-l));
        background-size: 200% 100%;
        animation: none;
    }
    @keyframes relHeroShimmer { 0%{background-position:0} 100%{background-position:200%} }
    .rel-hero::after {
        content: "";
        position: absolute;
        right: 0; top: 0; bottom: 0;
        width: 35%;
        background: transparent;
        pointer-events: none;
    }
    .rel-hero-left { position: relative; z-index: 1; }
    .rel-hero-eyebrow {
        font-family: 'IBM Plex Mono', monospace;
        font-size: .58rem; font-weight: 800;
        letter-spacing: .18em; text-transform: uppercase;
        color: rgba(var(--rgb-c-run-l),.60);
        display: flex; align-items: center; gap: 8px;
        margin-bottom: 6px;
    }
    .rel-hero-eyebrow::before {
        content: "";
        width: 6px; height: 6px; border-radius: 50%;
        background: var(--f-run-l);
        box-shadow: none;
    }
    .rel-hero-title {
        font-size: 1.32rem; font-weight: 800;
        color: var(--t-n11); letter-spacing: -.01em;
        margin: 0 0 4px;
        text-shadow: none;
    }
    .rel-hero-desc {
        font-size: .76rem; color: rgba(var(--rgb-t-n9),.77);
        line-height: 1.5;
    }
    .rel-hero-right {
        position: relative; z-index: 1;
        display: flex; flex-direction: column;
        align-items: flex-end; gap: 8px;
    }
    .rel-session-badge {
        display: flex; align-items: center; gap: 6px;
        padding: 5px 12px;
        background: rgba(var(--rgb-run-l),.06);
        border: 1px solid rgba(var(--rgb-run-l),.22);
        border-radius: 999px;
        font-family: 'IBM Plex Mono', monospace;
        font-size: .58rem; font-weight: 700;
        letter-spacing: .08em; color: rgba(var(--rgb-c-run-l),.75);
    }
    .rel-session-badge::before {
        content: "";
        width: 5px; height: 5px; border-radius: 50%;
        background: var(--f-run-l);
        box-shadow: none;
    }
    /* ── KPI GRID V2 ── */
    .kpi-grid-v2 {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(185px, 1fr));
        gap: 14px;
        margin-bottom: 26px;
    }
    .kpi-v2 {
        position: relative;
        overflow: hidden;
        padding: 18px 20px 16px;
        border-radius: 12px;
        background: linear-gradient(145deg, rgba(var(--rgb-sup-a),.96), rgba(var(--rgb-sup-b),.90));
        border: 1px solid rgba(var(--rgb-n9),.12);
        box-shadow: 0 8px 24px var(--shc-2), inset 0 1px 0 rgba(var(--rgb-lift),.02);
        transition: transform .2s, border-color .2s, box-shadow .2s;
    }
    .kpi-v2:hover {
        transform: translateY(-2px);
        border-color: var(--kc, var(--ta400));
        box-shadow: 0 18px 44px var(--shc-2);
    }
    .kpi-v2::before {
        content: "";
        position: absolute;
        left: 0; top: 0; bottom: 0; width: 3px;
        background: var(--kc, var(--f-acc));
        box-shadow: none;
    }
    .kpi-v2::after {
        content: "";
        position: absolute;
        inset: 0;
        background: linear-gradient(135deg, rgba(var(--rgb-lift),.025), transparent 50%);
        pointer-events: none;
    }
    .kpi-v2-icon { font-size: 1.1rem; margin-bottom: 8px; display: block; }
    .kpi-v2-val {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 2rem; font-weight: 800; line-height: 1;
        color: var(--kc, var(--t-n11));
        margin-bottom: 5px;
        text-shadow: none;
    }
    .kpi-v2-label {
        font-size: .62rem; text-transform: uppercase;
        letter-spacing: .10em; color: rgba(var(--rgb-t-n9),.79); font-weight: 700;
    }
    /* ── EXEC RECOMENDATION BANNER ── */
    .exec-banner {
        display: flex; align-items: flex-start; gap: 14px;
        padding: 16px 20px;
        background: linear-gradient(135deg, rgba(var(--rgb-run-l),.07), var(--ta040));
        border: 1px solid rgba(var(--rgb-run-l),.16);
        border-radius: 10px;
        margin-bottom: 22px;
    }
    .exec-banner-icon { font-size: 1.2rem; flex-shrink: 0; margin-top: 1px; }
    .exec-banner-content { flex: 1; }
    .exec-banner-title {
        font-size: .62rem; font-weight: 800;
        font-family: 'IBM Plex Mono', monospace;
        letter-spacing: .12em; text-transform: uppercase;
        color: rgba(var(--rgb-c-run-l),.70); margin-bottom: 4px;
    }
    .exec-banner-text { font-size: .82rem; color: var(--t-n10); line-height: 1.5; }
    </style>
    """, unsafe_allow_html=True)

        st.markdown(f"""
    <div class="rel-hero animate-fade-in">
        <div class="rel-hero-left">
            <div class="rel-hero-eyebrow">Exportação e auditoria</div>
            <div class="rel-hero-title">Relatórios e Indicadores</div>
            <div class="rel-hero-desc">Visualize indicadores consolidados, exporte dados e monitore o desempenho operacional de TI.</div>
        </div>
        <div class="rel-hero-right">
            <div class="rel-session-badge">sessão segura · {_sessao_min_r}min</div>
            <div style="font-family:'IBM Plex Mono',monospace;font-size:.60rem;color:rgba(var(--rgb-t-n9),.59);">{_ico("relogio", 11, extra="margin-right:5px")}{_agora_str_r}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

        _logout_col = st.columns([6, 1])[1]
        with _logout_col:
            if st.button("Sair", key="relatorios_logout", use_container_width=True, help="Encerrar sessão de relatórios"):
                st.session_state.autenticado_relatorios = False
                st.session_state.relat_auth_ts          = 0.0
                st.rerun(scope="app")

        df_rel   = carregar_dados_cacheados(DATABASE, COLUNAS)
        df_comp  = carregar_dados_cacheados(DATABASE_COMPRAS, COLUNAS_COMPRAS)
        df_rel_raw  = df_rel.copy()
        df_comp_raw = df_comp.copy()
        agora    = datetime.now()
        hoje     = agora.date()
        ult_dia  = calendar.monthrange(hoje.year, hoje.month)[1]
        is_last  = (hoje.day == ult_dia)
        is_hora  = (agora.hour == 18)

        st.markdown("""<div class="sec-label animate-fade-in"><div class="sec-dot"></div>FILTROS DO RELATÓRIO</div>""", unsafe_allow_html=True)

        def _parse_data_col(_df):
            return pd.to_datetime(_df["Data"], format="%d/%m/%Y %H:%M", errors="coerce")

        _dt_ch_all   = _parse_data_col(df_rel).dropna()  if not df_rel.empty  else pd.Series([], dtype="datetime64[ns]")
        _dt_comp_all = _parse_data_col(df_comp).dropna() if not df_comp.empty else pd.Series([], dtype="datetime64[ns]")
        _todas_dt    = pd.concat([_dt_ch_all, _dt_comp_all])
        _min_dt_rel  = _todas_dt.min().date() if not _todas_dt.empty else hoje
        _max_dt_rel  = _todas_dt.max().date() if not _todas_dt.empty else hoje

        _setores_opts = sorted(set(
            (df_rel["Setor"].dropna().unique().tolist()  if not df_rel.empty  else []) +
            (df_comp["Setor"].dropna().unique().tolist() if not df_comp.empty else [])
        ))
        _status_ch_opts   = sorted(df_rel["Status"].dropna().unique().tolist())       if not df_rel.empty  else []
        _prio_opts        = sorted(df_rel["Prioridade"].dropna().unique().tolist())    if not df_rel.empty  else []
        _status_comp_opts = sorted(df_comp["Status"].dropna().unique().tolist())       if not df_comp.empty else []

        with st.container(border=True, key="cartao_10"):
            ff1, ff2, ff3, ff4, ff5 = st.columns([1.6, 1.3, 1.3, 1.3, 0.7])
            with ff1:
                _periodo_sel = st.date_input(
                    "Período (Data)", value=(_min_dt_rel, _max_dt_rel),
                    min_value=_min_dt_rel, max_value=_max_dt_rel,
                    key="rel_filtro_periodo",
                )
            with ff2:
                _f_setor = st.multiselect("Setor", options=_setores_opts, key="rel_filtro_setor", placeholder="Todos")
            with ff3:
                _f_status = st.multiselect("Status (Chamados)", options=_status_ch_opts, key="rel_filtro_status", placeholder="Todos")
            with ff4:
                _f_prio = st.multiselect("Prioridade", options=_prio_opts, key="rel_filtro_prioridade", placeholder="Todas")
            with ff5:
                st.markdown("<div style='height:1.7rem'></div>", unsafe_allow_html=True)
                if st.button("Limpar", use_container_width=True, key="rel_filtro_limpar"):
                    for _k in ("rel_filtro_periodo", "rel_filtro_setor", "rel_filtro_status",
                               "rel_filtro_prioridade", "rel_filtro_status_compra"):
                        st.session_state.pop(_k, None)
                    st.rerun(scope="fragment")
            _f_status_compra = st.multiselect("Status (Compras)", options=_status_comp_opts, key="rel_filtro_status_compra", placeholder="Todos")

        if isinstance(_periodo_sel, (tuple, list)) and len(_periodo_sel) == 2:
            _dt_ini_rel, _dt_fim_rel = _periodo_sel
        else:
            _dt_ini_rel, _dt_fim_rel = _min_dt_rel, _max_dt_rel

        if not df_rel.empty:
            _dt_ch_f = _parse_data_col(df_rel)
            _mask_ch = ((_dt_ch_f.dt.date >= _dt_ini_rel) & (_dt_ch_f.dt.date <= _dt_fim_rel)) | _dt_ch_f.isna()
            df_rel = df_rel[_mask_ch]
            if _f_setor:
                df_rel = df_rel[df_rel["Setor"].isin(_f_setor)]
            if _f_status:
                df_rel = df_rel[df_rel["Status"].isin(_f_status)]
            if _f_prio:
                df_rel = df_rel[df_rel["Prioridade"].isin(_f_prio)]

        if not df_comp.empty:
            _dt_comp_f = _parse_data_col(df_comp)
            _mask_comp = ((_dt_comp_f.dt.date >= _dt_ini_rel) & (_dt_comp_f.dt.date <= _dt_fim_rel)) | _dt_comp_f.isna()
            df_comp = df_comp[_mask_comp]
            if _f_setor:
                df_comp = df_comp[df_comp["Setor"].isin(_f_setor)]
            if _f_status_compra:
                df_comp = df_comp[df_comp["Status"].isin(_f_status_compra)]

        _filtros_ativos = bool(_f_setor or _f_status or _f_prio or _f_status_compra or
                                _dt_ini_rel != _min_dt_rel or _dt_fim_rel != _max_dt_rel)
        if _filtros_ativos:
            st.caption(f"Filtros ativos — exibindo {len(df_rel)} de {len(df_rel_raw)} chamado(s) e {len(df_comp)} de {len(df_comp_raw)} compra(s).")

        total_ch    = len(df_rel)
        abertos     = len(df_rel[df_rel['Status'] == 'Aberto'])          if not df_rel.empty else 0
        atendimento = len(df_rel[df_rel['Status'] == 'Em Atendimento'])  if not df_rel.empty else 0
        concluidos  = len(df_rel[df_rel['Status'] == 'Concluído'])       if not df_rel.empty else 0
        taxa        = f"{int(concluidos/total_ch*100)}%" if total_ch else "—"
        urgentes    = len(df_rel[df_rel['Prioridade'].isin(['Alta','Crítica','Urgente'])]) if not df_rel.empty else 0
        total_comp  = len(df_comp)
        comp_aprov  = len(df_comp[df_comp['Status'] == 'Aprovado'])      if not df_comp.empty else 0
        comp_pend   = len(df_comp[df_comp['Status'] == 'Pendente'])      if not df_comp.empty else 0

        top_setor = df_rel['Setor'].value_counts().idxmax() if not df_rel.empty else "—"
        top_prio  = df_rel['Prioridade'].value_counts().idxmax() if not df_rel.empty else "—"
        df_rel_ativos = df_rel[df_rel["Status"].isin(["Aberto", "Em Atendimento"])] if not df_rel.empty else pd.DataFrame(columns=COLUNAS)
        df_comp_ativos = df_comp[df_comp["Status"].isin(["Pendente", "Em Análise"])] if not df_comp.empty else pd.DataFrame(columns=COLUNAS_COMPRAS)
        crit_ativos = int((df_rel_ativos["Prioridade"] == "Crítica").sum()) if not df_rel_ativos.empty else 0
        alta_ativos = int((df_rel_ativos["Prioridade"] == "Alta").sum()) if not df_rel_ativos.empty else 0
        if crit_ativos:
            recomendacao_exec = f"Priorizar {crit_ativos} chamado(s) crítico(s) antes de novas demandas administrativas."
        elif alta_ativos:
            recomendacao_exec = "Atenção aos chamados de alta prioridade na fila para evitar acúmulo."
        else:
            recomendacao_exec = "Operação saudável no momento: manter monitoramento e fechar registros antigos."

        st.markdown("""<div class="sec-label animate-fade-in"><div class="sec-dot"></div>INDICADORES DO PERÍODO</div>""", unsafe_allow_html=True)

        st.markdown(f"""
        <div class="kpi-grid-v2 animate-fade-in">
            <div class="kpi-v2" style="--kc:var(--c-acc)">
                <span class="kpi-v2-icon">{_ico("chamado", 18)}</span>
                <div class="kpi-v2-val">{total_ch}</div>
                <div class="kpi-v2-label">Total de Chamados</div>
            </div>
            <div class="kpi-v2" style="--kc:var(--c-ok)">
                <span class="kpi-v2-icon">{_ico("ok", 18)}</span>
                <div class="kpi-v2-val">{taxa}</div>
                <div class="kpi-v2-label">Taxa de Resolução</div>
            </div>
            <div class="kpi-v2" style="--kc:var(--c-crit-l)">
                <span class="kpi-v2-icon">{_ico("alerta", 18)}</span>
                <div class="kpi-v2-val">{urgentes}</div>
                <div class="kpi-v2-label">Alta / Crítica</div>
            </div>
            <div class="kpi-v2" style="--kc:var(--c-run-l)">
                <span class="kpi-v2-icon">{_ico("carrinho", 18)}</span>
                <div class="kpi-v2-val">{total_comp}</div>
                <div class="kpi-v2-label">Requisições de Compra</div>
            </div>
            <div class="kpi-v2" style="--kc:var(--c-run)">
                <span class="kpi-v2-icon">{_ico("ampulheta", 18)}</span>
                <div class="kpi-v2-val">{abertos}</div>
                <div class="kpi-v2-label">Chamados Abertos</div>
            </div>
            <div class="kpi-v2" style="--kc:var(--c-run)">
                <span class="kpi-v2-icon">{_ico("engrenagem", 18)}</span>
                <div class="kpi-v2-val">{atendimento}</div>
                <div class="kpi-v2-label">Em Atendimento</div>
            </div>
            <div class="kpi-v2" style="--kc:var(--c-ok)">
                <span class="kpi-v2-icon">{_ico("caixa", 18)}</span>
                <div class="kpi-v2-val">{comp_aprov}</div>
                <div class="kpi-v2-label">Compras Aprovadas</div>
            </div>
            <div class="kpi-v2" style="--kc:var(--c-crit-l)">
                <span class="kpi-v2-icon">{_ico("relogio", 18)}</span>
                <div class="kpi-v2-val">{comp_pend}</div>
                <div class="kpi-v2-label">Compras Pendentes</div>
            </div>
        </div>
    """, unsafe_allow_html=True)

        st.markdown(f"""
        <div class="exec-banner animate-fade-in">
            <span class="exec-banner-icon">{_ico("prancheta", 18)}</span>
            <div class="exec-banner-content">
                <div class="exec-banner-title">Leitura Executiva</div>
                <div class="exec-banner-text">
                    {html.escape(recomendacao_exec)}
                    &nbsp;·&nbsp; <strong>Críticos:</strong> {crit_ativos} chamado(s)
                    &nbsp;·&nbsp; <strong>Alta prioridade:</strong> {alta_ativos} chamado(s)
                    &nbsp;·&nbsp; <strong>Fila ativa:</strong> {len(df_rel_ativos)} chamado(s) e {len(df_comp_ativos)} compra(s)
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)

        c_inf1, c_inf2, c_inf3, c_inf4 = st.columns(4)
        _info_cards = [
            (c_inf1, "trofeu", "Setor Mais Demandante", top_setor, "var(--c-acc)"),
            (c_inf2, "raio",   "Prioridade Mais Frequente", top_prio, "var(--c-run)"),
            (c_inf3, "alerta", "Chamados Críticos (ativos)", str(crit_ativos), "var(--c-crit-l)"),
            (c_inf4, "ok",     "Concluídos no Período", str(concluidos), "var(--c-ok)"),
        ]
        for _ci, (col, ico, lbl, val, color) in enumerate(_info_cards):
            with col:
                with st.container(border=True, key=f"cartao_kpi_{_ci}"):
                    st.markdown(f"""
                    <div style="padding:2px 0;">
                        <div style="font-size:.60rem;color:rgba(var(--rgb-t-n9),.72);
                            text-transform:uppercase;letter-spacing:.10em;
                            font-weight:800;margin-bottom:6px;display:flex;align-items:center;gap:6px;">{_ico(ico, 13)}<span>{lbl}</span></div>
                        <div style="font-size:1.55rem;font-weight:800;color:{color};
                            font-family:'IBM Plex Mono',monospace;line-height:1;">{val}</div>
                    </div>
                """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("""<div class="sec-label animate-fade-in"><div class="sec-dot"></div>GRÁFICOS</div>""", unsafe_allow_html=True)

        import plotly.graph_objects as go

        _claro = _tema_atual() == "claro"
        _C_ACCENT = "#ffb700" if not _claro else "#b37f00"
        _C_GREEN  = "#4ec27c" if not _claro else "#1f9950"
        _C_BLUE   = "#6ba7f5" if not _claro else "#2f6fc0"
        _C_ORANGE = "#e8804a" if not _claro else "#c25a1e"
        _C_RED    = "#f5707f" if not _claro else "#c02434"
        _C_GREY   = "#78746c" if not _claro else "#8e8981"
        _STATUS_COLORS = {
            "Aberto": _C_BLUE, "Em Atendimento": _C_ACCENT,
            "Concluído": _C_GREEN, "Cancelado": _C_GREY, "Outros": _C_GREY,
        }
        _PRIO_COLORS = {
            "Baixa": _C_GREY, "Média": _C_BLUE, "Alta": _C_ORANGE,
            "Crítica": _C_RED, "Urgente": _C_RED,
        }

        def _aplicar_tema(fig, altura=300):
            fig.update_layout(
                template=("plotly_white" if _claro else "plotly_dark"),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font=dict(family="Archivo, sans-serif", color=_cor("n11"), size=12),
                margin=dict(l=10, r=10, t=14, b=10),
                height=altura,
                legend=dict(orientation="h", yanchor="bottom", y=-0.22, x=0, font=dict(size=11)),
                hoverlabel=dict(bgcolor=_cor("n2"), bordercolor=_cor("acc", .4), font_size=12),
            )
            fig.update_xaxes(showgrid=True, gridcolor=_cor("n9", .08), zeroline=False, color=_cor("n8"))
            fig.update_yaxes(showgrid=True, gridcolor=_cor("n9", .08), zeroline=False, color=_cor("n8"))
            return fig

        _PLOT_CFG = {"displayModeBar": False, "responsive": True}

        if df_rel.empty and df_comp.empty:
            st.info("Sem dados suficientes para gerar os gráficos analíticos.")
        else:
            g1c1, g1c2 = st.columns(2)
            with g1c1:
                with st.container(border=True, key="cartao_12"):
                    st.markdown("###### Distribuição de chamados por status")
                    if not df_rel.empty:
                        _CANON = ["Aberto", "Em Atendimento", "Concluído", "Cancelado"]
                        _sraw = df_rel["Status"].astype(str).str.strip()
                        _snorm = _sraw.where(_sraw.isin(_CANON), "Outros")
                        _vc = _snorm.value_counts()
                        _ordem = [x for x in _CANON if x in _vc.index] + (["Outros"] if "Outros" in _vc.index else [])
                        sc = _vc.reindex(_ordem)
                        fig = go.Figure(go.Pie(
                            labels=sc.index.tolist(), values=sc.values.tolist(),
                            hole=.62, sort=False, direction="clockwise",
                            marker=dict(colors=[_STATUS_COLORS.get(x, _cor("n7")) for x in sc.index],
                                        line=dict(color=_cor("n0"), width=2)),
                            textinfo="percent", textposition="inside",
                            insidetextorientation="horizontal",
                            textfont=dict(size=12, color=_cor("n0")),
                        ))
                        fig.update_layout(annotations=[dict(
                            text=f"<b>{total_ch}</b><br><span style='font-size:11px;color:{_cor('t-n8')}'>chamados</span>",
                            x=0.5, y=0.5, xref="paper", yref="paper",
                            showarrow=False, font=dict(size=22, color=_cor("n11")))])
                        st.plotly_chart(_aplicar_tema(fig, 300), use_container_width=True, config=_PLOT_CFG)
                    else:
                        st.caption("Sem chamados registrados.")
            with g1c2:
                with st.container(border=True, key="cartao_13"):
                    st.markdown("###### Chamados por setor")
                    if not df_rel.empty:
                        ss = df_rel["Setor"].value_counts().head(8).sort_values()
                        fig = go.Figure(go.Bar(
                            x=ss.values.tolist(), y=ss.index.tolist(), orientation="h",
                            marker=dict(color=ss.values.tolist(),
                                        colorscale=[[0, _cor("c-acc-d")], [1, _C_ACCENT]],
                                        line=dict(width=0)),
                            text=ss.values.tolist(), textposition="outside",
                            textfont=dict(color=_cor("n11"), size=11),
                        ))
                        st.plotly_chart(_aplicar_tema(fig, 300), use_container_width=True, config=_PLOT_CFG)
                    else:
                        st.caption("Sem chamados registrados.")

            with st.container(border=True, key="cartao_14"):
                st.markdown("###### Chamados por nível de criticidade")
                if not df_rel.empty:
                    ordem = ["Baixa", "Média", "Alta", "Crítica"]
                    pc = df_rel["Prioridade"].value_counts()
                    cats = [p for p in ordem if p in pc.index] + [p for p in pc.index if p not in ordem]
                    vals = [int(pc.get(p, 0)) for p in cats]
                    fig = go.Figure(go.Bar(
                        x=cats, y=vals,
                        marker=dict(color=[_PRIO_COLORS.get(p, _C_ACCENT) for p in cats], line=dict(width=0)),
                        text=vals, textposition="outside", textfont=dict(color=_cor("n11"), size=11),
                    ))
                    st.plotly_chart(_aplicar_tema(fig, 300), use_container_width=True, config=_PLOT_CFG)
                else:
                    st.caption("Sem chamados registrados.")

            with st.container(border=True, key="cartao_15"):
                st.markdown("###### Evolução de chamados no tempo")
                if not df_rel.empty:
                    _dft = df_rel.copy()
                    _dft["__dt"] = pd.to_datetime(_dft["Data"], format="%d/%m/%Y %H:%M", errors="coerce")
                    _dft = _dft.dropna(subset=["__dt"])
                    if not _dft.empty:
                        serie = _dft.groupby(_dft["__dt"].dt.date).size().sort_index()
                        fig = go.Figure(go.Scatter(
                            x=list(serie.index), y=serie.values.tolist(),
                            mode="lines", line=dict(color=_C_ACCENT, width=2.5, shape="spline"),
                            fill="tozeroy", fillcolor=_cor("acc", .12),
                            hovertemplate="%{x}<br><b>%{y}</b> chamado(s)<extra></extra>",
                        ))
                        st.plotly_chart(_aplicar_tema(fig, 280), use_container_width=True, config=_PLOT_CFG)
                    else:
                        st.caption("Datas insuficientes para a série temporal.")
                else:
                    st.caption("Sem chamados registrados.")

        def _xlsx_bytes(frames: dict) -> bytes:
            """Recebe {nome_aba: dataframe} e devolve bytes de um .xlsx formatado."""
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter
            HDR_FILL   = PatternFill("solid", fgColor="1E3A5F")
            HDR_FONT   = Font(bold=True, color="FFFFFF", name="Arial", size=10)
            HDR_ALIGN  = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ROW_FILL_E = PatternFill("solid", fgColor="EAF2FB")
            ROW_FILL_O = PatternFill("solid", fgColor="FFFFFF")
            CELL_ALIGN = Alignment(vertical="center", wrap_text=True)
            THIN       = Side(style="thin", color="CCCCCC")
            BORDER     = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
            wb = Workbook()
            wb.remove(wb.active)
            for sheet_name, df in frames.items():
                ws = wb.create_sheet(title=sheet_name[:31])
                if df.empty:
                    ws.append(["(sem dados neste período)"])
                    continue
                ws.append(list(df.columns))
                for cell in ws[1]:
                    cell.font      = HDR_FONT
                    cell.fill      = HDR_FILL
                    cell.alignment = HDR_ALIGN
                    cell.border    = BORDER
                ws.row_dimensions[1].height = 28
                for r_idx, row in enumerate(df.itertuples(index=False), start=2):
                    ws.append(list(row))
                    fill = ROW_FILL_E if r_idx % 2 == 0 else ROW_FILL_O
                    for cell in ws[r_idx]:
                        cell.fill      = fill
                        cell.alignment = CELL_ALIGN
                        cell.border    = BORDER
                        cell.font      = Font(name="Arial", size=9)
                for col_idx, col_cells in enumerate(ws.columns, start=1):
                    max_len = max((len(str(c.value or "")) for c in col_cells), default=8)
                    ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 52)
                ws.freeze_panes = "A2"
            buf_xl = io.BytesIO()
            wb.save(buf_xl)
            buf_xl.seek(0)
            return buf_xl.read()

        def gerar_xlsx_kpis(dfr=None, dfc=None) -> bytes:
            """Sem argumentos, usa os dados FILTRADOS em tela (df_rel/df_comp).
            Passando dfr/dfc explícitos (ex.: df_rel_raw/df_comp_raw), recalcula os
            KPIs a partir deles — usado no fechamento mensal automático, que deve
            sempre refletir o dado oficial completo, independente do filtro em tela."""
            _dfr = df_rel if dfr is None else dfr
            _dfc = df_comp if dfc is None else dfc
            _total_ch    = len(_dfr)
            _abertos     = len(_dfr[_dfr['Status'] == 'Aberto'])          if not _dfr.empty else 0
            _atendimento = len(_dfr[_dfr['Status'] == 'Em Atendimento'])  if not _dfr.empty else 0
            _concluidos  = len(_dfr[_dfr['Status'] == 'Concluído'])       if not _dfr.empty else 0
            _urgentes    = len(_dfr[_dfr['Prioridade'].isin(['Alta','Crítica','Urgente'])]) if not _dfr.empty else 0
            _total_comp  = len(_dfc)
            _comp_aprov  = len(_dfc[_dfc['Status'] == 'Aprovado'])        if not _dfc.empty else 0
            _comp_pend   = len(_dfc[_dfc['Status'] == 'Pendente'])       if not _dfc.empty else 0
            _top_setor   = _dfr['Setor'].value_counts().idxmax()         if not _dfr.empty else "—"
            _top_prio    = _dfr['Prioridade'].value_counts().idxmax()    if not _dfr.empty else "—"
            _dfr_ativos  = _dfr[_dfr["Status"].isin(["Aberto", "Em Atendimento"])] if not _dfr.empty else pd.DataFrame(columns=COLUNAS)
            _crit_ativos = int((_dfr_ativos["Prioridade"] == "Crítica").sum()) if not _dfr_ativos.empty else 0
            _alta_ativos = int((_dfr_ativos["Prioridade"] == "Alta").sum())    if not _dfr_ativos.empty else 0
            if _crit_ativos:
                _rec = f"Priorizar {_crit_ativos} chamado(s) crítico(s) antes de novas demandas administrativas."
            elif _alta_ativos:
                _rec = "Atenção aos chamados de alta prioridade na fila para evitar acúmulo."
            else:
                _rec = "Operação saudável no momento: manter monitoramento e fechar registros antigos."
            kpis = {
                "Métrica": [
                    "Total de Chamados","Chamados Abertos","Em Atendimento","Concluídos",
                    "Taxa de Resolução (%)","Alta/Crítica",
                    "Setor Mais Demandante","Prioridade Mais Frequente",
                    "Total Requisições Compra","Compras Aprovadas","Compras Pendentes",
                    "Chamados Críticos (ativos)","Alta Prioridade (ativos)","Recomendação Executiva"
                ],
                "Valor": [
                    _total_ch, _abertos, _atendimento, _concluidos,
                    f"{int(_concluidos/_total_ch*100) if _total_ch else 0}%",
                    _urgentes, _top_setor, _top_prio,
                    _total_comp, _comp_aprov, _comp_pend,
                    _crit_ativos, _alta_ativos, _rec
                ]
            }
            return _xlsx_bytes({"KPIs TI": pd.DataFrame(kpis)})

        def gerar_zip_completo(dfr=None, dfc=None):
            """Sem argumentos, exporta os dados FILTRADOS em tela. Passando dfr/dfc
            explícitos, exporta esses dataframes (usado com df_rel_raw/df_comp_raw
            no fechamento mensal automático)."""
            _dfr = df_rel if dfr is None else dfr
            _dfc = df_comp if dfc is None else dfc
            buf = io.BytesIO()
            ts  = agora.strftime("%Y%m%d_%H%M")
            with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
                z.writestr(f"kpis_{ts}.xlsx", gerar_xlsx_kpis(_dfr, _dfc))
                if not _dfr.empty:
                    z.writestr(f"chamados_completo_{ts}.xlsx",
                               _xlsx_bytes({"Chamados": _dfr}))
                if not _dfc.empty:
                    z.writestr(f"compras_completo_{ts}.xlsx",
                               _xlsx_bytes({"Compras": _dfc}))
                if not _dfr.empty:
                    setor_sum = _dfr.groupby('Setor').agg(
                        Chamados=('Status','count'),
                        Concluidos=('Status', lambda x: (x=='Concluído').sum()),
                        Abertos=('Status', lambda x: (x=='Aberto').sum())
                    ).reset_index()
                    prio_sum = _dfr['Prioridade'].value_counts().reset_index()
                    prio_sum.columns = ['Prioridade','Qtd']
                    z.writestr(f"chamados_analitico_{ts}.xlsx",
                               _xlsx_bytes({"Por Setor": setor_sum, "Por Prioridade": prio_sum}))
                pdf_bytes = gerar_pdf(_dfr)
                z.writestr(f"relatorio_chamados_{ts}.pdf", pdf_bytes)
                if not _dfc.empty:
                    pdf_comp_bytes = gerar_pdf_compras(_dfc)
                    z.writestr(f"relatorio_compras_{ts}.pdf", pdf_comp_bytes)
            buf.seek(0)
            return buf.read()

        st.markdown("""<div class="sec-label animate-fade-in"><div class="sec-dot"></div>EXPORTAR DADOS</div>""", unsafe_allow_html=True)

        with st.container(border=True, key="cartao_16"):
            st.markdown("""
            <div style='font-size:.76rem;color:rgba(var(--rgb-t-n9),.72);
                margin-bottom:16px;line-height:1.5;'>
                Exporte qualquer módulo operacional para análise externa, auditoria ou arquivo.
                Arquivos gerados com timestamp automático.
            </div>
        """, unsafe_allow_html=True)
            d1, d2, d3, d4 = st.columns(4)

            with d1:
                st.download_button(
                    "KPIs Consolidados (XLSX)",
                    data=gerar_xlsx_kpis(),
                    file_name=f"kpis_ti_{agora.strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with d2:
                if not df_rel.empty:
                    st.download_button(
                        "Banco de Chamados (XLSX)",
                        data=_xlsx_bytes({"Chamados": df_rel}),
                        file_name=f"chamados_{agora.strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.button("Banco de Chamados (XLSX)", disabled=True, use_container_width=True)
            with d3:
                if not df_rel.empty:
                    st.download_button(
                        "Relatório de Incidentes (PDF)",
                        data=gerar_pdf(df_rel),
                        file_name=f"chamados_{agora.strftime('%Y%m%d_%H%M')}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                else:
                    st.button("Relatório de Incidentes (PDF)", disabled=True, use_container_width=True)
            with d4:
                st.download_button(
                "Pacote Completo (.ZIP)",
                data=gerar_zip_completo(),
                file_name=f"relatorio_ti_completo_{agora.strftime('%Y%m%d_%H%M')}.zip",
                mime="application/zip",
                use_container_width=True
        )

            if not df_comp.empty:
                st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
                d5, d6, _, _ = st.columns(4)
                with d5:
                    st.download_button(
                        "Fluxo de Compras (XLSX)",
                        data=_xlsx_bytes({"Compras": df_comp}),
                        file_name=f"compras_{agora.strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                with d6:
                    st.download_button(
                        "Auditoria Logística (PDF)",
                        data=gerar_pdf_compras(df_comp),
                        file_name=f"compras_{agora.strftime('%Y%m%d_%H%M')}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )

        st.markdown("---")
        st.markdown("""<div class="sec-label animate-fade-in"><div class="sec-dot"></div>FECHAMENTO MENSAL E BACKUPS</div>""", unsafe_allow_html=True)

        EXPORT_DIR = os.path.join(_DADOS, "exports")
        if not os.path.exists(EXPORT_DIR):
            os.makedirs(EXPORT_DIR)

        flag_file = os.path.join(EXPORT_DIR, f"auto_{hoje.strftime('%Y%m')}.flag")
        ja_exportou = os.path.exists(flag_file)

        with st.container(border=True, key="cartao_17"):
            col_ag1, col_ag2 = st.columns([1.8, 1.2])
            with col_ag1:
                st.markdown(f"""
                <div style='display:flex;align-items:center;gap:10px;margin-bottom:10px;'>
                    <span>{_ico("engrenagem", 17)}</span>
                    <div>
                        <div style='font-size:.82rem;font-weight:800;color:var(--t-n11);'>Motor de Fechamento Mensal</div>
                        <div style='font-size:.68rem;color:rgba(var(--rgb-t-n9),.68);'>Executa automaticamente no último dia do mês às 18h00</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
                st.markdown(f"""
                <p style='color:rgba(var(--rgb-t-n9),.76); font-size:.76rem; line-height:1.6; margin-bottom:15px;'>
                O microsserviço de auditoria monitora em tempo real a janela de fechamento do ciclo. 
                Ao atingir o limite, compila os logs automaticamente, salva em <code style='color:var(--c-acc);background:var(--ta080);padding:1px 5px;border-radius:4px;'>exports/</code> e notifica a equipe.
                </p>
            """, unsafe_allow_html=True)

                if is_last and is_hora and not ja_exportou:
                    zip_data = gerar_zip_completo(df_rel_raw, df_comp_raw)
                    ts_auto  = agora.strftime("%Y%m%d_%H%M")
                    zip_path = os.path.join(EXPORT_DIR, f"relatorio_mensal_{ts_auto}.zip")
                    with open(zip_path, 'wb') as f:
                        f.write(zip_data)
                    with open(flag_file, 'w') as f:
                        f.write(agora.isoformat())
                    _corpo_rel = f"Export automático de {hoje.strftime('%m/%Y')} salvo em exports/"
                    notificar_windows("📊 RELATÓRIO MENSAL GERADO", _corpo_rel)
                    notificar_whatsapp(
                        "RELATÓRIO MENSAL GERADO",
                        _corpo_rel,
                        url=_url_pagina("relatorios"),
                    )
                    st.success(f"Relatório mensal de {hoje.strftime('%m/%Y')} gerado e salvo automaticamente!")
                elif ja_exportou:
                    st.success(f"Status: Sincronização mensal de {hoje.strftime('%m/%Y')} validada e armazenada no servidor.")
                else:
                    prox = date(hoje.year if hoje.month < 12 else hoje.year+1,
                                hoje.month+1 if hoje.month < 12 else 1, 1)
                    dias_restantes = (prox - hoje).days - 1
                    st.info(f"Monitoramento em execução. Próximo fechamento em **{dias_restantes} dia(s)** (dia {ult_dia:02d}/{hoje.month:02d}/{hoje.year} às 18h00).")

            with col_ag2:
                st.markdown(f"""
                <div style='display:flex;align-items:center;gap:8px;margin-bottom:12px;'>
                    <span>{_ico("pasta", 16)}</span>
                    <div style='font-size:.80rem;font-weight:800;color:var(--t-n11);'>Repositório de Backups</div>
                </div>
            """, unsafe_allow_html=True)
                arquivos = sorted([f for f in os.listdir(EXPORT_DIR) if f.endswith('.zip')], reverse=True)
                if arquivos:
                    @st.cache_data(show_spinner=False)
                    def _ler_zip_cache(caminho: str, _mtime: float):
                        """Lê o ZIP do disco apenas quando o arquivo muda (mtime como cache key)."""
                        with open(caminho, 'rb') as _f:
                            return _f.read()

                    for arq in arquivos[:5]:
                        arq_path = os.path.join(EXPORT_DIR, arq)
                        _mtime = os.path.getmtime(arq_path)
                        st.download_button(
                            f"{arq[:24]}...",
                            data=_ler_zip_cache(arq_path, _mtime),
                            file_name=arq,
                            mime="application/zip",
                            use_container_width=True,
                            key=f"dl_saved_{arq}"
                        )
                else:
                    st.caption("Nenhum snapshot automático localizado no cluster.")

        st.markdown("---")
        st.markdown("""<div class="sec-label animate-fade-in"><div class="sec-dot"></div>MATRIZ DE DADOS DE INFRAESTRUTURA E AUDITORIA</div>""", unsafe_allow_html=True)

        st.markdown('''<style>
        .tbl-pro{width:100%;border-collapse:collapse;font-family:var(--font);font-size:.78rem;margin-top:2px;}
        .tbl-pro thead th{text-align:left;padding:8px 14px;color:var(--t-n9);font-weight:700;text-transform:uppercase;letter-spacing:.05em;font-size:.62rem;border-bottom:1px solid var(--ta160);white-space:nowrap;}
        .tbl-pro tbody td{padding:8px 14px;color:var(--t-n11);border-bottom:1px solid rgba(var(--rgb-n9),.05);vertical-align:middle;}
        .tbl-pro tbody tr:hover{background:var(--ta050);}
        .tbl-pro tbody tr:last-child td{border-bottom:none;}
        .tp-bar{position:relative;background:var(--ta080);border-radius:5px;height:18px;min-width:78px;display:flex;align-items:center;overflow:hidden;}
        .tp-bar-fill{position:absolute;left:0;top:0;height:100%;border-radius:5px;background:linear-gradient(90deg,var(--f-acc-d),var(--f-acc));}
        .tp-bar-val{position:relative;padding-left:8px;font-family:var(--mono);font-weight:700;font-size:.70rem;color:var(--t-n11);}
        .tp-badge{display:inline-block;padding:2px 10px;border-radius:99px;font-size:.68rem;font-weight:700;background:rgba(var(--rgb-n9),.12);color:var(--t-n10);}
        .tp-badge.b-red{background:rgba(var(--rgb-crit-l),.16);color:var(--c-crit-l);}
        .tp-badge.b-amber{background:rgba(var(--rgb-run),.16);color:var(--c-run);}
        .tp-badge.b-green{background:rgba(var(--rgb-ok),.16);color:var(--c-ok);}
        .tp-badge.b-blue{background:var(--ta160);color:var(--c-acc-l);}
        .tp-badge.b-orange{background:rgba(var(--rgb-run),.16);color:var(--c-run-l);}
        </style>''', unsafe_allow_html=True)

        _PRIO_BADGE = {"Baixa":"🔵", "Média":"🟡", "Alta":"🟠", "Crítica":"🔴"}
        _STATUS_BADGE = {"Aberto":"🟡","Em Atendimento":"🔵","Concluído":"🟢","Pendente":"🟡","Aprovado":"🟢","Em Análise":"🔵","Cancelado":"⚪","Recusado":"🔴"}

        def _tabela_pro(df, bar_col=None, badge_map=None, headers=None):
            """Grade nativa (st.dataframe): barra de volume via ProgressColumn,
            emoji como indicador de status/prioridade no lugar de badge HTML."""
            df_show = df.copy()
            if badge_map:
                for col, icones in badge_map.items():
                    if col in df_show.columns:
                        df_show[col] = df_show[col].map(lambda v: f"{icones.get(str(v), '⚪')} {v}")
            column_config = {}
            if headers:
                df_show = df_show.rename(columns=headers)
            if bar_col:
                bar_label = (headers or {}).get(bar_col, bar_col)
                maxv = float(df[bar_col].max()) if len(df) else 0.0
                column_config[bar_label] = st.column_config.ProgressColumn(
                    bar_label, format="%d", min_value=0, max_value=max(maxv, 1)
                )
            st.dataframe(df_show, use_container_width=True, hide_index=True, column_config=column_config)

        tab_t1, tab_t2, tab_t3, tab_t4 = st.tabs([
            "Volumetria por Setor",
            "Criticidade e Prioridades",
            "Demandas Logísticas",
            "Categorias de Chamados"
        ])

        with tab_t1:
            with st.container(border=True, key="cartao_18"):
                if not df_rel.empty:
                    setor_tab = df_rel.groupby('Setor').agg(
                        Total=('Status','count'),
                        Abertos=('Status', lambda x: (x=='Aberto').sum()),
                        Em_Atendimento=('Status', lambda x: (x=='Em Atendimento').sum()),
                        Concluidos=('Status', lambda x: (x=='Concluído').sum()),
                    ).reset_index().sort_values('Total', ascending=False)
                    setor_tab['Resolutividade'] = setor_tab.apply(
                        lambda r: f"{int(r['Concluidos']/r['Total']*100)}%" if r['Total'] else '—', axis=1)
                    _tabela_pro(setor_tab, bar_col="Total", headers={
                        "Setor": "Setor", "Total": "Total",
                        "Abertos": "🔴 Abertos", "Em_Atendimento": "🟡 Atendimento",
                        "Concluidos": "✅ Concluídos", "Resolutividade": "Resolução"})
                else:
                    st.info("Nenhum indicador setorial indexado no banco de dados.")

        with tab_t2:
            with st.container(border=True, key="cartao_19"):
                if not df_rel.empty:
                    prio_tab = df_rel.groupby('Prioridade').agg(
                        Total=('Status','count'),
                        Abertos=('Status', lambda x: (x=='Aberto').sum()),
                        Concluidos=('Status', lambda x: (x=='Concluído').sum()),
                    ).reset_index()
                    ordem = {'Baixa':0,'Média':1,'Alta':2,'Crítica':3}
                    prio_tab['_ord'] = prio_tab['Prioridade'].map(lambda p: ordem.get(p, 99))
                    prio_tab = prio_tab.sort_values('_ord').drop(columns=['_ord'])
                    _tabela_pro(prio_tab, bar_col="Total", badge_map={"Prioridade": _PRIO_BADGE}, headers={
                        "Prioridade": "Criticidade", "Total": "Volume",
                        "Abertos": "Fila Ativa", "Concluidos": "Encerrados"})
                else:
                    st.info("Nenhuma métrica de severidade registrada para o período atual.")

        with tab_t3:
            with st.container(border=True, key="cartao_20"):
                if not df_comp.empty:
                    comp_tab = df_comp.groupby(['Categoria','Status']).agg(
                        Itens=('Item','count'),
                        Qtd_Total=('Quantidade','sum'),
                    ).reset_index().sort_values('Itens', ascending=False)
                    _tabela_pro(comp_tab, bar_col="Itens", badge_map={"Status": _STATUS_BADGE}, headers={
                        "Categoria": "Categoria", "Status": "Fase",
                        "Itens": "Modelos", "Qtd_Total": "Unidades"})
                else:
                    st.info("Nenhuma requisição logística catalogada para consolidação.")

        with tab_t4:
            with st.container(border=True, key="cartao_21"):
                if not df_rel.empty:
                    df_cat = df_rel.copy()
                    df_cat["Categoria Operacional"] = df_cat.apply(extrair_categoria, axis=1)
                    cat_tab = df_cat.groupby("Categoria Operacional").agg(
                        Total=("Status", "count"),
                        Abertos=("Status", lambda x: (x == "Aberto").sum()),
                        Concluidos=("Status", lambda x: (x == "Concluído").sum()),
                    ).reset_index().sort_values("Total", ascending=False)
                    _tabela_pro(cat_tab, bar_col="Total", headers={
                        "Categoria Operacional": "Categoria", "Total": "Total",
                        "Abertos": "🔴 Abertos", "Concluidos": "✅ Concluídos"})
                else:
                    st.info("Sem chamados para leitura por categoria.")
    _render_relatorios()


