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
# NOTA DE PERFORMANCE: fpdf e openpyxl NÃO são importados aqui.
# São bibliotecas pesadas usadas apenas em recursos específicos (exportar PDF,
# exportar Excel e análise por IA). Importá-las no topo penaliza o PRIMEIRO
# carregamento da página inicial em ~1,5-2,5s sem necessidade. Por isso o
# import é feito "lazy", dentro da função que realmente usa cada uma.
# (altair foi removido: estava importado mas nunca era usado.)

# ── AUTO-INSTALL: garante dependências extras no venv do usuário ──
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


# --- CONTROLE DE UPTIME ---
@st.cache_resource
def get_server_start_time():
    return time.time()

server_start = get_server_start_time()

@st.cache_data(show_spinner=False)  # cache invalidado manualmente em salvar_dataframe_seguro (sem ttl: nada de recarga periódica nem caixa "Running...")
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


# --- CONFIGURAÇÃO INICIAL DA PÁGINA ---
# Ícone do portal (favicon da aba). Usa a logo do portal se existir; senão, 💻.
_PORTAL_ICON = os.path.join(os.path.dirname(os.path.abspath(__file__)), "portal_logo.png")
st.set_page_config(
    page_title="IT ServiceDesk Pro",
    page_icon=(_PORTAL_ICON if os.path.exists(_PORTAL_ICON) else "💻"),
    layout="wide",
)

# ── CSS CRÍTICO ANTI-FOUC ─────────────────────────────────────────────
# Primeira coisa pintada na página. Garante o fundo escuro e a cor de texto
# IMEDIATAMENTE (antes do CSS global e antes das fontes Geist baixarem),
# eliminando o "flash de letras sem estilo" ao recarregar (F5).
# O fade só roda no 1º carregamento da sessão — assim os cliques seguintes
# (reruns) NÃO re-animam a tela, o que dava sensação de lentidão a cada ação.
_primeira_carga = "_app_booted" not in st.session_state
st.session_state["_app_booted"] = True
_fade_css = ("""
[data-testid="stMain"] { animation: antiFoucReveal .32s ease-out both; }
@keyframes antiFoucReveal { from { opacity: 0; } to { opacity: 1; } }
""" if _primeira_carga else "")
st.markdown(f"""
<style id="anti-fouc-critical">
html, body, .stApp,
[data-testid="stAppViewContainer"], [data-testid="stMain"] {{
    background-color: #02040b !important;
    color: #e2e8f0 !important;
    /* Pilha de fontes definida JÁ na primeira pintura: enquanto a Geist baixa,
       o texto aparece numa fonte de sistema próxima — elimina o "flash de
       letras sem estilo" (FOUT) e o reflow quando a Geist troca. */
    font-family: 'Geist', system-ui, -apple-system, 'Segoe UI', Roboto,
                 'Helvetica Neue', Arial, sans-serif !important;
}}
/* Suaviza a substituição da fonte sem piscar */
html {{ -webkit-font-smoothing: antialiased; text-rendering: optimizeLegibility; }}
{_fade_css}
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


@st.cache_data(show_spinner=False)
def _get_global_css() -> str:
    """CSS global cacheado — montado uma única vez por sessão do servidor."""
    return """
<style>
/* Fontes carregadas via <link> (preconnect) fora deste bloco — um @import aqui
   bloquearia a aplicação de TODO o CSS até a fonte baixar, causando FOUC
   (texto sem estilo antes do design carregar). */

/* ── CORREÇÕES ANTI-MODO CLARO (FORÇAR DARK THEME) ── */
[data-testid="stAppViewContainer"] {
    background:
        radial-gradient(ellipse at 10% -5%, rgba(56,189,248,.22), transparent 32%),
        radial-gradient(ellipse at 90% 5%, rgba(56,189,248,.14), transparent 28%),
        radial-gradient(ellipse at 50% 100%, rgba(56,189,248,.08), transparent 40%),
        linear-gradient(180deg, #020408 0%, #040810 48%, #02040b 100%) !important;
    color: #e2e8f0 !important; /* var(--txt-1) */
}
[data-testid="stHeader"] {
    background: transparent !important;
    backdrop-filter: none !important;
    -webkit-backdrop-filter: none !important;
}
[data-testid="stSidebar"] {
    background-color: #080d18 !important; /* var(--bg-1) */
}
/* Força títulos, textos e labels nativos a não ficarem pretos */
.st-emotion-cache-16txtl3, .st-emotion-cache-10trblm, h1, h2, h3, h4, h5, h6, p, label, span {
    color: #e2e8f0 !important;
}
/* Evita que inputs nativos fiquem com texto escuro no modo claro */
input, textarea, select {
    color: #e2e8f0 !important;
    -webkit-text-fill-color: #e2e8f0 !important;
}

/* ── TOKENS ── */
:root {
    --bg-0:       #02040b;
    --bg-1:       #060b15;
    --bg-2:       #0a1020;
    --bg-card:    #0c1828;
    --panel:      rgba(12,24,40,.92);
    --panel-2:    rgba(6,11,21,.85);
    --bg-hover:   #111e35;
    --border:     rgba(56,189,248,0.08);
    --border-soft:rgba(148,163,184,.10);
    --border-hi:  rgba(56,189,248,0.55);
    --accent:     #38bdf8;
    --accent-dim: rgba(56,189,248,0.14);
    --accent-soft:rgba(56,189,248,.07);
    --violet:     #a78bfa;
    --violet-soft:rgba(167,139,250,.08);
    --green:      #34d399;
    --amber:      #fbbf24;
    --red:        #f87171;
    --txt-1:      #e2e8f0;
    --txt-2:      #7d8fa8;
    --txt-3:      #3d5068;
    --font:       'Geist', sans-serif;
    --mono:       'Geist Mono', monospace;
    --radius:     14px;
    /* Typography scale */
    --text-xs:    0.70rem;
    --text-sm:    0.78rem;
    --text-md:    0.86rem;
    /* Shadows */
    --shadow:     0 20px 50px rgba(0,0,0,.52), 0 0 0 1px rgba(56,189,248,.04);
    --shadow-soft:0 10px 26px rgba(0,0,0,.32);
    --glow-accent:0 0 24px rgba(56,189,248,.35), 0 0 60px rgba(56,189,248,.12);
    --glow-violet:0 0 24px rgba(167,139,250,.35), 0 0 60px rgba(167,139,250,.12);
    --glow-green: 0 0 20px rgba(52,211,153,.4), 0 0 50px rgba(52,211,153,.12);
    --glow-red:   0 0 20px rgba(248,113,113,.4), 0 0 50px rgba(248,113,113,.12);
}

/* ── RESET / BASE ── */
.stApp {
    background:
        linear-gradient(rgba(56,189,248,.015) 1px, transparent 1px),
        linear-gradient(90deg, rgba(56,189,248,.015) 1px, transparent 1px);
    background-size: 44px 44px;
    font-family: var(--font);
}
.stApp::before {
    content: "";
    position: fixed;
    inset: 0;
    pointer-events: none;
    background:
        linear-gradient(180deg, rgba(255,255,255,.03), transparent 18%),
        linear-gradient(90deg, transparent, rgba(56,189,248,.08), transparent);
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
    background: linear-gradient(90deg, transparent, rgba(56,189,248,.45), transparent) !important;
    box-shadow: 0 0 8px rgba(56,189,248,.25) !important;
    margin: 22px 0 !important;
}
/* ── SCROLLBARS GLOBAIS — azul ciano combinando com a sidebar ── */
/* Firefox */
*, html, body {
    scrollbar-width: thin;
    scrollbar-color: rgba(56,189,248,.55) transparent;
}

/* Webkit (Chrome, Edge, Safari) — todos os elementos */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: rgba(6,11,21,.40); border-radius: 99px; }
::-webkit-scrollbar-thumb {
    background: linear-gradient(180deg,
        #38bdf8 0%,
        rgba(56,189,248,.65) 60%,
        rgba(56,189,248,.35) 100%);
    border-radius: 99px;
    border: none;
    box-shadow: 0 0 6px rgba(56,189,248,.45);
    transition: background .2s, box-shadow .2s;
}
::-webkit-scrollbar-thumb:hover {
    background: linear-gradient(180deg,
        #7dd3fc 0%,
        #38bdf8 50%,
        rgba(56,189,248,.80) 100%);
    box-shadow: 0 0 12px rgba(56,189,248,.70);
}
::-webkit-scrollbar-corner { background: transparent; }

/* Sidebar scrollbar — mesma cor, ligeiramente mais fina */
[data-testid="stSidebar"] ::-webkit-scrollbar { width: 3px; }
[data-testid="stSidebar"] ::-webkit-scrollbar-track { background: transparent; }
[data-testid="stSidebar"] ::-webkit-scrollbar-thumb {
    background: linear-gradient(180deg,
        #38bdf8 0%,
        rgba(56,189,248,.60) 100%);
    border-radius: 99px;
    border: none;
    box-shadow: 0 0 5px rgba(56,189,248,.40);
}
[data-testid="stSidebar"] ::-webkit-scrollbar-thumb:hover {
    background: #38bdf8;
    box-shadow: 0 0 10px rgba(56,189,248,.65);
}
[data-testid="stSidebar"] > div {
    scrollbar-width: thin;
    scrollbar-color: rgba(56,189,248,.55) transparent;
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
    scrollbar-color: rgba(56,189,248,.55) transparent !important;
}
[data-testid="stDataFrame"] ::-webkit-scrollbar,
.dvn-scroller::-webkit-scrollbar,
.glideDataEditor ::-webkit-scrollbar { width: 5px !important; height: 5px !important; }
[data-testid="stDataFrame"] ::-webkit-scrollbar-track,
.dvn-scroller::-webkit-scrollbar-track { background: rgba(6,11,21,.40) !important; border-radius: 99px !important; }
[data-testid="stDataFrame"] ::-webkit-scrollbar-thumb,
.dvn-scroller::-webkit-scrollbar-thumb {
    background: linear-gradient(180deg, #38bdf8 0%, rgba(56,189,248,.55) 100%) !important;
    border-radius: 99px !important;
    box-shadow: 0 0 6px rgba(56,189,248,.45) !important;
}
[data-testid="stDataFrame"] ::-webkit-scrollbar-thumb:hover,
.dvn-scroller::-webkit-scrollbar-thumb:hover {
    background: #7dd3fc !important;
    box-shadow: 0 0 12px rgba(56,189,248,.70) !important;
}

/* Injeta o estilo dentro dos iframes de dataframe via CSS variable trick */
iframe { color-scheme: dark; }
#MainMenu, footer, .stDeployButton, [data-testid="stDecoration"] { display: none !important; visibility: hidden !important; }

/* ── SIDEBAR / NAV SHELL ── */
[data-testid="stSidebar"] {
    position: relative !important;
    height: 100vh !important;
    top: 0 !important;
    background:
        radial-gradient(circle at 18% 4%, rgba(56,189,248,.16), transparent 24%),
        linear-gradient(180deg, #08101d 0%, #050912 54%, #04070f 100%) !important;
    border-right: 1px solid rgba(56,189,248,.13) !important;
    box-shadow: inset -1px 0 0 rgba(255,255,255,.03), 18px 0 60px rgba(0,0,0,.22) !important;
    transition: transform .22s ease, width .22s ease, box-shadow .22s ease !important;
}
[data-testid="stSidebar"]::before {
    content: "";
    position: absolute;
    inset: 0;
    pointer-events: none;
    background-image:
        linear-gradient(rgba(56,189,248,.045) 1px, transparent 1px),
        linear-gradient(90deg, rgba(56,189,248,.045) 1px, transparent 1px);
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
    border: 1px solid rgba(56,189,248,.22) !important;
    background: rgba(8,13,24,.72) !important;
    box-shadow: 0 10px 28px rgba(0,0,0,.35), 0 0 0 1px rgba(56,189,248,.06) !important;
    backdrop-filter: blur(10px) !important;
}
[data-testid="collapsedControl"]:hover {
    transform: translateY(-1px) scale(1.02) !important;
    border-color: rgba(56,189,248,.45) !important;
    box-shadow: 0 16px 40px rgba(0,0,0,.44), 0 0 26px rgba(56,189,248,.22) !important;
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
    background: linear-gradient(135deg, rgba(15,25,41,.72), rgba(9,15,28,.52)) !important;
    border: 1px solid rgba(148,163,184,.10) !important;
    box-shadow: inset 0 1px 0 rgba(255,255,255,.035) !important;
    cursor: pointer !important;
    /* Transições de alta performance */
    transition: transform 0.25s cubic-bezier(0.4, 0, 0.2, 1), 
                border-color 0.25s cubic-bezier(0.4, 0, 0.2, 1), 
                background 0.25s cubic-bezier(0.4, 0, 0.2, 1), 
                box-shadow 0.25s cubic-bezier(0.4, 0, 0.2, 1) !important;
}

/* Oculta bolinha padrão do Streamlit */
[data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child {
    display: none !important;
}

[data-testid="stSidebar"] div[role="radiogroup"] > label [data-testid="stMarkdownContainer"] {
    width: 100% !important;
    padding: 0 !important;
}

[data-testid="stSidebar"] div[role="radiogroup"] > label p {
    margin: 0 !important;
    color: #cad7ea !important;
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
    border-color: rgba(56,189,248,.38) !important;
    background: linear-gradient(135deg, rgba(56,189,248,.14), rgba(15,25,41,.74)) !important;
    box-shadow: 0 12px 28px rgba(0,0,0,.3), 0 0 14px rgba(56,189,248,0.12), inset 0 1px 0 rgba(255,255,255,.05) !important;
}

[data-testid="stSidebar"] div[role="radiogroup"] > label:hover p {
    color: #f8fbff !important;
    text-shadow: 0 0 8px rgba(56,189,248,0.3) !important;
}

/* ── MICROINTERAÇÃO: CLIQUE ATIVO (SENSACÃO TÁTIL) ── */
[data-testid="stSidebar"] div[role="radiogroup"] > label:active {
    transform: translateX(4px) scale(0.985) !important;
    filter: brightness(0.9);
}

/* ── ESTADO SELECIONADO / ATIVO ── */
[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) {
    border-color: rgba(56,189,248,.60) !important;
    background: 
        linear-gradient(135deg, rgba(56,189,248,.22), rgba(167,139,250,.12)), 
        linear-gradient(135deg, rgba(15,25,41,.95), rgba(9,15,28,.80)) !important;
    box-shadow: 0 0 0 1px rgba(56,189,248,.16), 0 0 24px rgba(56,189,248,.22), 0 16px 38px rgba(56,189,248,.14) !important;
}

[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) p {
    color: #ffffff !important;
    font-weight: 800 !important;
    text-shadow: 0 0 12px rgba(56,189,248,0.4) !important;
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
    background: #38bdf8 !important;
    box-shadow: 0 0 14px #38bdf8 !important;
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
    background: rgba(125,143,168,.35) !important;
    transform: translateY(-50%) !important;
    transition: background 0.2s ease, box-shadow 0.2s ease, transform 0.2s ease !important;
}

[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked)::after {
    background: #38bdf8 !important;
    box-shadow: 0 0 14px #38bdf8 !important;
    transform: translateY(-50%) scale(1.2) !important;
}

/* ── CONTAINERS / CARDS ── */
[data-testid="stVerticalBlockBorderWrapper"] {
    position: relative !important;
    overflow: hidden !important;
    background:
        linear-gradient(145deg, rgba(15,25,41,.92), rgba(8,13,24,.82)) !important;
    border: 1px solid rgba(56,189,248,.10) !important;
    border-radius: var(--radius) !important;
    /* Relevo 3D: brilho interno no topo + sombra interna na base (espessura)
       + sombras externas projetadas (a caixa salta da tela) */
    box-shadow:
        inset 0 1px 0 rgba(255,255,255,.05),
        inset 0 -7px 16px rgba(0,0,0,.38),
        0 10px 24px rgba(0,0,0,.42),
        0 2px 6px rgba(0,0,0,.30) !important;
    transition: border-color .2s ease, box-shadow .2s ease, transform .2s ease !important;
}
[data-testid="stVerticalBlockBorderWrapper"]::before {
    content: "";
    position: absolute;
    inset: 0 0 auto;
    height: 1px;
    background: linear-gradient(90deg, transparent, rgba(56,189,248,.32), transparent);
    pointer-events: none;
}
[data-testid="stVerticalBlockBorderWrapper"]:hover {
    border-color: rgba(56,189,248,.50) !important;
    transform: translateY(-3px) !important;
    box-shadow:
        inset 0 1px 0 rgba(255,255,255,.06),
        inset 0 -7px 16px rgba(0,0,0,.34),
        0 0 0 1px rgba(56,189,248,.12),
        0 0 32px rgba(56,189,248,.10),
        0 26px 60px rgba(0,0,0,.55) !important;
}

/* ════════════════════════════════════════════════════════
   FORM INPUTS — PREMIUM, SELETORES data-baseweb DIRETOS
   ════════════════════════════════════════════════════════ */

/* Labels */
.stTextInput label, .stTextArea label, .stSelectbox label,
.stNumberInput label, .stFileUploader label,
.stMultiSelect label, .stSlider label {
    color: #d7e3f4 !important;
    font-size: .82rem !important;
    font-family: var(--font) !important;
    font-weight: 700 !important;
    text-transform: none !important;
    letter-spacing: .01em !important;
    margin-bottom: 5px !important;
}

/* ── base-input: container real do st.text_input ── */
div[data-baseweb="base-input"] {
    background: linear-gradient(160deg, #08131f 0%, #050d18 100%) !important;
    border: 1px solid rgba(56,189,248,.22) !important;
    border-radius: 10px !important;
    box-shadow:
        inset 0 1px 0 rgba(255,255,255,.04),
        inset 0 -1px 0 rgba(0,0,0,.35),
        0 4px 18px rgba(0,0,0,.40) !important;
    transition: border-color .22s ease, box-shadow .22s ease, transform .18s ease !important;
}
div[data-baseweb="base-input"]:hover {
    border-color: rgba(56,189,248,.40) !important;
    transform: translateY(-1px) !important;
}
div[data-baseweb="base-input"]:focus-within {
    border-color: rgba(56,189,248,.72) !important;
    box-shadow:
        0 0 0 3px rgba(56,189,248,.13),
        0 0 28px rgba(56,189,248,.16),
        inset 0 1px 0 rgba(56,189,248,.10) !important;
    transform: translateY(-2px) !important;
}
div[data-baseweb="base-input"] input {
    background: transparent !important;
    color: #bfdbfe !important;
    font-family: var(--font) !important;
    font-size: .90rem !important;
    font-weight: 500 !important;
    caret-color: #38bdf8 !important;
    -webkit-text-fill-color: #bfdbfe !important;
}
div[data-baseweb="base-input"] input::placeholder {
    color: rgba(148,178,208,.50) !important;
    -webkit-text-fill-color: rgba(148,178,208,.50) !important;
    font-style: normal !important;
    font-size: .85rem !important;
}

/* ── textarea ── */
div[data-baseweb="textarea"] {
    background: linear-gradient(160deg, #08131f 0%, #050d18 100%) !important;
    border: 1px solid rgba(56,189,248,.22) !important;
    border-radius: 10px !important;
    box-shadow: inset 0 1px 0 rgba(255,255,255,.04), 0 4px 18px rgba(0,0,0,.40) !important;
    transition: border-color .22s ease, box-shadow .22s ease, transform .18s ease !important;
}
div[data-baseweb="textarea"]:hover {
    border-color: rgba(56,189,248,.40) !important;
    transform: translateY(-1px) !important;
}
div[data-baseweb="textarea"]:focus-within {
    border-color: rgba(56,189,248,.72) !important;
    box-shadow: 0 0 0 3px rgba(56,189,248,.13), 0 0 28px rgba(56,189,248,.16), inset 0 1px 0 rgba(56,189,248,.10) !important;
    transform: translateY(-2px) !important;
}
div[data-baseweb="textarea"] textarea {
    background: transparent !important;
    color: #bfdbfe !important;
    -webkit-text-fill-color: #bfdbfe !important;
    font-family: var(--font) !important;
    font-size: .88rem !important;
    caret-color: #38bdf8 !important;
}

/* ── select: container visual ── */
div[data-baseweb="select"] > div:first-child {
    background: linear-gradient(160deg, #08131f 0%, #050d18 100%) !important;
    border: 1px solid rgba(56,189,248,.22) !important;
    border-radius: 10px !important;
    box-shadow: inset 0 1px 0 rgba(255,255,255,.04), 0 4px 18px rgba(0,0,0,.40) !important;
    transition: border-color .22s ease, box-shadow .22s ease, transform .18s ease !important;
}
div[data-baseweb="select"] > div:first-child:hover {
    border-color: rgba(56,189,248,.40) !important;
    transform: translateY(-1px) !important;
}
div[data-baseweb="select"]:focus-within > div:first-child {
    border-color: rgba(56,189,248,.72) !important;
    box-shadow: 0 0 0 3px rgba(56,189,248,.13), 0 0 28px rgba(56,189,248,.16), inset 0 1px 0 rgba(56,189,248,.10) !important;
    transform: translateY(-2px) !important;
}
div[data-baseweb="select"] span,
div[data-baseweb="select"] div[class*="single"] {
    color: #bfdbfe !important;
    -webkit-text-fill-color: #bfdbfe !important;
    font-family: var(--font) !important;
    font-size: .88rem !important;
    font-weight: 500 !important;
}
div[data-baseweb="select"] svg {
    fill: rgba(56,189,248,.52) !important;
    transition: transform .22s ease, fill .2s ease !important;
}
div[data-baseweb="select"]:focus-within svg {
    fill: #38bdf8 !important;
    transform: rotate(180deg) !important;
}

/* ── Dropdown menu ── */
ul[data-baseweb="menu"] {
    background: linear-gradient(160deg, #07101e, #040a16) !important;
    border: 1px solid rgba(56,189,248,.26) !important;
    border-radius: 12px !important;
    box-shadow: 0 0 0 1px rgba(56,189,248,.06), 0 24px 60px rgba(0,0,0,.80), 0 0 40px rgba(56,189,248,.10) !important;
    padding: 5px !important;
    backdrop-filter: blur(14px) !important;
    overflow: hidden !important;
}
li[role="option"] {
    border-radius: 8px !important;
    color: #8fb0cc !important;
    font-family: var(--font) !important;
    font-size: .87rem !important;
    font-weight: 500 !important;
    transition: background .14s ease, color .14s ease, padding-left .14s ease !important;
    padding: 9px 14px !important;
}
li[role="option"]:hover {
    background: linear-gradient(90deg, rgba(56,189,248,.16), rgba(56,189,248,.06)) !important;
    color: #e0f2fe !important;
    padding-left: 18px !important;
}
li[role="option"][aria-selected="true"] {
    background: linear-gradient(90deg, rgba(56,189,248,.18), rgba(56,189,248,.06)) !important;
    color: #7dd3fc !important;
    font-weight: 700 !important;
    border-left: 2px solid rgba(56,189,248,.75) !important;
}

/* ── FILTER PANEL v2 — TOGGLE CHIP DESIGN ── */
.fpanel {
    position: relative;
    background: linear-gradient(145deg, rgba(8,15,28,.98), rgba(4,9,18,.96));
    border: 1px solid rgba(56,189,248,.14);
    border-radius: 14px;
    overflow: hidden;
    box-shadow: 0 0 0 1px rgba(56,189,248,.04), 0 28px 64px rgba(0,0,0,.52);
    margin-bottom: 0;
}
.fpanel::before {
    content: "";
    position: absolute;
    top: 0; left: 0; right: 0; height: 1px;
    background: linear-gradient(90deg, transparent 3%, rgba(56,189,248,.55) 38%, rgba(167,139,250,.38) 62%, transparent 97%);
}
.fpanel-violet { border-color: rgba(167,139,250,.16); }
.fpanel-violet::before {
    background: linear-gradient(90deg, transparent 3%, rgba(167,139,250,.55) 38%, rgba(56,189,248,.32) 62%, transparent 97%);
}
.fpanel-hdr {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 11px 20px 10px;
    border-bottom: 1px solid rgba(56,189,248,.06);
    background: linear-gradient(90deg, rgba(56,189,248,.055), rgba(167,139,250,.025), transparent);
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
    box-shadow: 0 0 10px var(--accent);
    flex-shrink: 0;
    animation: pulse-dot 2.2s ease-in-out infinite;
}
@keyframes pulse-dot {
    0%, 100% { opacity: 1; box-shadow: 0 0 8px var(--accent); }
    50% { opacity: .55; box-shadow: 0 0 18px var(--accent), 0 0 32px rgba(56,189,248,.28); }
}
.fpanel-meta { display: flex; align-items: center; gap: 8px; }
.fpanel-badge {
    font-family: var(--mono);
    font-size: .62rem;
    font-weight: 700;
    letter-spacing: .07em;
    color: var(--txt-3);
    padding: 3px 10px;
    border: 1px solid rgba(56,189,248,.10);
    border-radius: 999px;
    background: rgba(56,189,248,.04);
    transition: all .2s ease;
}
.fpanel-badge-warn {
    color: #f87171;
    border-color: rgba(248,113,113,.28);
    background: rgba(248,113,113,.06);
    box-shadow: 0 0 12px rgba(248,113,113,.08);
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
    color: #4a6480;
    margin-bottom: 9px;
    padding-bottom: 7px;
    border-bottom: 1px solid rgba(56,189,248,.06);
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
    border-top: 1px solid rgba(56,189,248,.06);
}

/* ── CHIP (wraps each st.checkbox) ── */
.fchip {
    position: relative;
    border-radius: 7px;
    transition: background .15s ease, border-color .15s ease, box-shadow .18s ease, transform .12s ease;
}
/* ON state — chip lit */
.fchip-on {
    background: linear-gradient(135deg, rgba(var(--chip-rgb, 56,189,248), .14), rgba(var(--chip-rgb, 56,189,248), .06));
    background: color-mix(in srgb, var(--chip-color, #38bdf8) 12%, transparent);
    border: 1px solid color-mix(in srgb, var(--chip-color, #38bdf8) 40%, transparent);
    box-shadow: 0 0 0 0 transparent;
}
/* OFF state — chip dimmed/inactive */
.fchip-off {
    background: rgba(10,15,26,.45);
    border: 1px solid rgba(148,163,184,.07);
    opacity: .52;
}
.fchip-off:hover { opacity: .75; border-color: rgba(148,163,184,.18); }
.fchip-on:hover {
    border-color: color-mix(in srgb, var(--chip-color, #38bdf8) 70%, transparent) !important;
    box-shadow: 0 0 12px color-mix(in srgb, var(--chip-color, #38bdf8) 25%, transparent) !important;
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
    border: 1.5px solid color-mix(in srgb, var(--chip-color, #38bdf8) 50%, rgba(148,163,184,.3)) !important;
    background: rgba(4,9,18,.6) !important;
    flex-shrink: 0 !important;
    transition: background .15s ease, border-color .15s ease !important;
}
.fchip-on .stCheckbox > label > span:first-child {
    background: color-mix(in srgb, var(--chip-color, #38bdf8) 85%, white 15%) !important;
    border-color: var(--chip-color, #38bdf8) !important;
    box-shadow: 0 0 6px color-mix(in srgb, var(--chip-color, #38bdf8) 60%, transparent) !important;
}
.fchip .stCheckbox > label > div {
    font-family: var(--mono) !important;
    font-size: .69rem !important;
    font-weight: 700 !important;
    letter-spacing: .06em !important;
    text-transform: uppercase !important;
    color: color-mix(in srgb, var(--chip-color, #38bdf8) 90%, white 10%) !important;
    transition: color .15s ease !important;
    line-height: 1.2 !important;
}
.fchip-off .stCheckbox > label > div {
    color: #3d5068 !important;
}
/* Remove Streamlit's default focus ring clutter */
.fchip .stCheckbox > label:focus-within {
    outline: none !important;
    box-shadow: none !important;
}

/* ── SEARCH + DATE INPUTS inside fpanel-body ── */
.filter-search-wrap .stTextInput > div {
    background: rgba(4,9,18,.75) !important;
    border: 1px solid rgba(56,189,248,.14) !important;
    border-radius: 8px !important;
    height: 40px !important;
    transition: border-color .2s ease, box-shadow .2s ease !important;
}
.filter-search-wrap .stTextInput > div:focus-within {
    border-color: rgba(56,189,248,.52) !important;
    box-shadow: 0 0 0 3px rgba(56,189,248,.09), 0 0 20px rgba(56,189,248,.07) !important;
}
.filter-search-wrap .stTextInput input::placeholder { color: #2d4258 !important; }

/* date inputs in filter row */
.fpanel-body [data-testid="stDateInput"] > div {
    background: rgba(4,9,18,.75) !important;
    border: 1px solid rgba(56,189,248,.14) !important;
    border-radius: 8px !important;
    height: 40px !important;
}
.fpanel-body [data-testid="stDateInput"] > div:focus-within {
    border-color: rgba(56,189,248,.52) !important;
    box-shadow: 0 0 0 3px rgba(56,189,248,.09) !important;
}

/* Reset button inside filter */
.filter-reset-btn .stButton > button {
    background: rgba(248,113,113,.06) !important;
    border: 1px solid rgba(248,113,113,.22) !important;
    border-radius: 8px !important;
    color: rgba(248,113,113,.7) !important;
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
    background: rgba(248,113,113,.14) !important;
    border-color: rgba(248,113,113,.45) !important;
    color: #f87171 !important;
    box-shadow: 0 0 16px rgba(248,113,113,.14) !important;
}

/* ══════════════════════════════════════════════════════════
   FILTER PANEL v3 — MULTISELECT TECH OVERHAUL
   ══════════════════════════════════════════════════════════ */

/* ── Container do painel de filtros ── */
[data-testid="stVerticalBlockBorderWrapper"]:has(.stMultiSelect) {
    background: linear-gradient(145deg, rgba(6,12,24,.98), rgba(3,7,16,.96)) !important;
    border: 1px solid rgba(56,189,248,.18) !important;
    border-radius: 14px !important;
    padding: 4px !important;
    box-shadow:
        0 0 0 1px rgba(56,189,248,.05),
        inset 0 1px 0 rgba(56,189,248,.08),
        0 24px 56px rgba(0,0,0,.55) !important;
    position: relative !important;
    overflow: hidden !important;
}
[data-testid="stVerticalBlockBorderWrapper"]:has(.stMultiSelect)::before {
    content: "" !important;
    position: absolute !important;
    top: 0; left: 0; right: 0; height: 1px !important;
    background: linear-gradient(90deg,
        transparent 2%,
        rgba(56,189,248,.65) 35%,
        rgba(167,139,250,.45) 65%,
        transparent 98%) !important;
    pointer-events: none !important;
}

/* ── Label dos multiselects: estilo monospace neon ── */
.stMultiSelect label {
    font-family: 'Geist Mono', ui-monospace, monospace !important;
    font-size: .62rem !important;
    font-weight: 800 !important;
    letter-spacing: .15em !important;
    text-transform: uppercase !important;
    color: #4a6e8a !important;
    margin-bottom: 5px !important;
    display: flex !important;
    align-items: center !important;
    gap: 5px !important;
}

/* ── Wrapper externo do multiselect ── */
.stMultiSelect > div {
    background: rgba(4,9,20,.82) !important;
    border: 1px solid rgba(56,189,248,.16) !important;
    border-radius: 9px !important;
    min-height: 38px !important;
    box-shadow: inset 0 1px 0 rgba(255,255,255,.03), inset 0 -1px 0 rgba(0,0,0,.3) !important;
    transition: border-color .2s ease, box-shadow .22s ease, background .2s ease !important;
    position: relative !important;
}
.stMultiSelect > div:focus-within {
    border-color: rgba(56,189,248,.55) !important;
    background: rgba(6,14,28,.92) !important;
    box-shadow:
        0 0 0 3px rgba(56,189,248,.10),
        0 0 24px rgba(56,189,248,.12),
        inset 0 1px 0 rgba(56,189,248,.06) !important;
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
   'Geist'/'Geist Mono', o topo das letras fica cortado pelo container. */
.stMultiSelect [data-baseweb="select"] > div:first-child > div {
    line-height: 1.5 !important;
    overflow: visible !important;
    display: flex !important;
    align-items: center !important;
}
.stMultiSelect [data-baseweb="select"] [class*="placeholder"] {
    line-height: 1.5 !important;
    overflow: visible !important;
    color: #4a6e8a !important;
    font-family: 'Geist', sans-serif !important;
    font-size: .8rem !important;
}

/* ── Tags individuais (itens selecionados) ── */
.stMultiSelect [data-baseweb="tag"] {
    background: linear-gradient(135deg,
        rgba(56,189,248,.18) 0%,
        rgba(56,189,248,.08) 100%) !important;
    border: 1px solid rgba(56,189,248,.38) !important;
    border-radius: 6px !important;
    height: 24px !important;
    padding: 0 8px 0 9px !important;
    box-shadow: 0 0 10px rgba(56,189,248,.12), inset 0 1px 0 rgba(56,189,248,.12) !important;
    transition: border-color .15s ease, box-shadow .15s ease, transform .12s ease !important;
    max-width: 160px !important;
    position: relative !important;
}
.stMultiSelect [data-baseweb="tag"]:hover {
    border-color: rgba(56,189,248,.65) !important;
    box-shadow: 0 0 16px rgba(56,189,248,.22) !important;
    transform: translateY(-1px) !important;
}

/* Texto dentro da tag */
.stMultiSelect [data-baseweb="tag"] span:first-child {
    font-family: 'Geist Mono', ui-monospace, monospace !important;
    font-size: .63rem !important;
    font-weight: 700 !important;
    letter-spacing: .07em !important;
    text-transform: uppercase !important;
    color: #7dd3fc !important;
    white-space: nowrap !important;
    overflow: hidden !important;
    text-overflow: ellipsis !important;
}

/* Botão X (remover) da tag */
.stMultiSelect [data-baseweb="tag"] [role="presentation"] {
    color: rgba(56,189,248,.5) !important;
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
    color: #f87171 !important;
}

/* Ícone X via SVG override — make it visible */
.stMultiSelect [data-baseweb="tag"] svg {
    fill: rgba(56,189,248,.55) !important;
    width: 10px !important;
    height: 10px !important;
    transition: fill .15s ease !important;
}
.stMultiSelect [data-baseweb="tag"]:hover svg {
    fill: #f87171 !important;
}

/* Placeholder "Todos / Todas" */
.stMultiSelect [data-baseweb="select"] input::placeholder {
    color: #2e4a62 !important;
    font-family: 'Geist Mono', monospace !important;
    font-size: .75rem !important;
    letter-spacing: .06em !important;
}
.stMultiSelect [data-baseweb="select"] input {
    color: #e2e8f0 !important;
    font-family: 'Geist', sans-serif !important;
    font-size: .8rem !important;
    caret-color: #38bdf8 !important;
    background: transparent !important;
    border: none !important;
    outline: none !important;
    min-width: 40px !important;
}

/* Seta dropdown */
.stMultiSelect [data-baseweb="select"] svg[data-icon="chevron-down"],
.stMultiSelect [data-baseweb="select"] > div:last-child svg {
    fill: rgba(56,189,248,.4) !important;
    transition: fill .2s ease, transform .2s ease !important;
}
.stMultiSelect > div:focus-within svg[data-icon="chevron-down"],
.stMultiSelect > div:focus-within > div:last-child svg {
    fill: rgba(56,189,248,.8) !important;
    transform: rotate(180deg) !important;
}

/* ── Dropdown menu flutuante ── */
[data-baseweb="popover"] [data-baseweb="menu"],
[data-baseweb="menu"] {
    background: linear-gradient(145deg, rgba(7,13,26,.99), rgba(4,8,18,.98)) !important;
    border: 1px solid rgba(56,189,248,.22) !important;
    border-radius: 10px !important;
    box-shadow:
        0 0 0 1px rgba(56,189,248,.06),
        0 24px 60px rgba(0,0,0,.70),
        0 0 40px rgba(56,189,248,.08) !important;
    padding: 4px !important;
    backdrop-filter: blur(12px) !important;
    overflow: hidden !important;
}

/* Opções individuais do dropdown */
[data-baseweb="menu"] li,
[data-baseweb="option"] {
    background: transparent !important;
    border-radius: 6px !important;
    padding: 7px 12px !important;
    font-family: 'Geist Mono', monospace !important;
    font-size: .72rem !important;
    font-weight: 600 !important;
    letter-spacing: .05em !important;
    color: #7d9ab8 !important;
    text-transform: uppercase !important;
    transition: background .14s ease, color .14s ease !important;
    cursor: pointer !important;
    border: 1px solid transparent !important;
    margin-bottom: 2px !important;
}
[data-baseweb="menu"] li:hover,
[data-baseweb="option"]:hover {
    background: rgba(56,189,248,.10) !important;
    border-color: rgba(56,189,248,.22) !important;
    color: #e2f4fd !important;
}
/* Opção já selecionada */
[data-baseweb="option"][aria-selected="true"] {
    background: linear-gradient(135deg, rgba(56,189,248,.18), rgba(56,189,248,.06)) !important;
    border-color: rgba(56,189,248,.30) !important;
    color: #7dd3fc !important;
}

/* ── Botão ↺ reset — estilo neon vermelho ── */
[data-testid="stVerticalBlockBorderWrapper"]:has(.stMultiSelect) .stButton > button {
    background: rgba(248,113,113,.06) !important;
    border: 1px solid rgba(248,113,113,.28) !important;
    border-radius: 8px !important;
    color: rgba(248,113,113,.75) !important;
    font-family: 'Geist Mono', monospace !important;
    font-size: .85rem !important;
    font-weight: 700 !important;
    height: 38px !important;
    width: 100% !important;
    padding: 0 !important;
    transition: all .18s ease !important;
    box-shadow: none !important;
}
[data-testid="stVerticalBlockBorderWrapper"]:has(.stMultiSelect) .stButton > button:hover {
    background: rgba(248,113,113,.14) !important;
    border-color: rgba(248,113,113,.55) !important;
    color: #f87171 !important;
    box-shadow: 0 0 18px rgba(248,113,113,.18) !important;
    transform: none !important;
}

/* ── Date inputs dentro do painel de filtros ── */
[data-testid="stVerticalBlockBorderWrapper"]:has(.stMultiSelect) [data-testid="stDateInput"] > div,
[data-testid="stVerticalBlockBorderWrapper"]:has(.stMultiSelect) [data-baseweb="input"] {
    background: rgba(4,9,20,.82) !important;
    border: 1px solid rgba(56,189,248,.16) !important;
    border-radius: 9px !important;
    min-height: 38px !important;
    box-shadow: inset 0 1px 0 rgba(255,255,255,.03) !important;
    transition: border-color .2s ease, box-shadow .2s ease !important;
}
[data-testid="stVerticalBlockBorderWrapper"]:has(.stMultiSelect) [data-testid="stDateInput"] > div:focus-within,
[data-testid="stVerticalBlockBorderWrapper"]:has(.stMultiSelect) [data-baseweb="input"]:focus-within {
    border-color: rgba(56,189,248,.50) !important;
    box-shadow: 0 0 0 3px rgba(56,189,248,.09), 0 0 18px rgba(56,189,248,.08) !important;
}

/* ── Busca livre abaixo do painel ── */
.stTextInput:has(#f_txt_mgr), .stTextInput:has(#f_txt_cp) {
    margin-top: 6px !important;
}
.stTextInput:has(#f_txt_mgr) > div,
.stTextInput:has(#f_txt_cp) > div {
    background: rgba(4,9,20,.85) !important;
    border: 1px solid rgba(56,189,248,.14) !important;
    border-radius: 9px !important;
    box-shadow: inset 0 1px 0 rgba(255,255,255,.025) !important;
    transition: border-color .2s ease, box-shadow .2s ease !important;
}
.stTextInput:has(#f_txt_mgr) > div:focus-within,
.stTextInput:has(#f_txt_cp) > div:focus-within {
    border-color: rgba(56,189,248,.52) !important;
    box-shadow: 0 0 0 3px rgba(56,189,248,.09), 0 0 20px rgba(56,189,248,.09) !important;
}
.stTextInput:has(#f_txt_mgr) label,
.stTextInput:has(#f_txt_cp) label {
    font-family: 'Geist Mono', monospace !important;
    font-size: .62rem !important;
    font-weight: 800 !important;
    letter-spacing: .15em !important;
    text-transform: uppercase !important;
    color: #4a6e8a !important;
}

/* Keep old filter-panel-wrap styles for any remaining usage */
.filter-panel-wrap {
    position: relative;
    background: linear-gradient(145deg, rgba(10,18,34,.97), rgba(6,11,22,.94));
    border: 1px solid rgba(56,189,248,.16);
    border-radius: 14px;
    padding: 0;
    overflow: hidden;
    box-shadow: 0 0 0 1px rgba(56,189,248,.04), 0 24px 56px rgba(0,0,0,.48);
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
    color: #5a7a9a;
    margin-bottom: 8px;
    padding-bottom: 6px;
    border-bottom: 1px solid rgba(56,189,248,.05);
}
.filter-group-icon {
    width: 18px; height: 18px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    border-radius: 5px;
    font-size: .7rem;
}
.fgi-prio  { background: rgba(251,191,36,.12);  border: 1px solid rgba(251,191,36,.22); }
.fgi-status{ background: rgba(52,211,153,.12);   border: 1px solid rgba(52,211,153,.22); }
.fgi-setor { background: rgba(167,139,250,.12);  border: 1px solid rgba(167,139,250,.22); }
.fgi-busca { background: rgba(56,189,248,.10);   border: 1px solid rgba(56,189,248,.22); }

/* ── BUTTONS ── */
.stButton > button[kind="primary"] {
    position: relative !important;
    overflow: hidden !important;
    background: linear-gradient(135deg, #38bdf8 0%, #0ea5e9 55%, #0284c7 100%) !important;
    border: 1px solid rgba(125,211,252,.45) !important;
    border-radius: 11px !important;
    color: #04121f !important;
    font-family: var(--mono) !important;
    font-weight: 700 !important;
    font-size: .8rem !important;
    letter-spacing: .10em !important;
    text-transform: uppercase !important;
    padding: 4px 0 !important;
    box-shadow: 0 8px 22px rgba(56,189,248,.28), inset 0 1px 0 rgba(255,255,255,.35) !important;
    transition: transform .18s ease, box-shadow .22s ease, filter .2s ease !important;
}
/* brilho sutil que desliza ao passar o mouse */
.stButton > button[kind="primary"]::after {
    content: "" !important;
    position: absolute !important;
    top: 0; left: -140%;
    width: 55%; height: 100%;
    background: linear-gradient(115deg, transparent 0%, rgba(255,255,255,.40) 50%, transparent 100%) !important;
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
    box-shadow: 0 14px 32px rgba(56,189,248,.42), inset 0 1px 0 rgba(255,255,255,.45) !important;
    filter: brightness(1.05) !important;
}
.stButton > button[kind="primary"]:active { transform: translateY(0) !important; }
.stButton > button[kind="primary"] *,
.stButton button[kind="primary"] p,
[data-testid="stButton"] button[kind="primary"] p,
[data-testid="stBaseButton-primary"] p,
[data-testid="stBaseButton-primary"] {
    color: #04070f !important;
    filter: none !important;
}
.stButton > button[kind="secondary"] {
    background: linear-gradient(180deg, rgba(15,25,41,.88), rgba(8,13,24,.82)) !important;
    border: 1px solid rgba(148,163,184,.16) !important;
    border-radius: 8px !important;
    color: #cbd5e1 !important;
    font-family: var(--font) !important;
    font-weight: 650 !important;
    transition: all .2s ease !important;
}
.stButton > button[kind="secondary"]:hover {
    border-color: var(--border-hi) !important;
    color: #ffffff !important;
    box-shadow: 0 0 22px rgba(56,189,248,.12) !important;
}
.stFormSubmitButton > button {
    background: linear-gradient(135deg, #38bdf8 0%, #0ea5e9 55%, #0284c7 100%) !important;
    border: 1px solid rgba(125,211,252,.45) !important;
    border-radius: 11px !important;
    color: #04121f !important;
    font-family: var(--mono) !important;
    font-weight: 700 !important;
    font-size: .8rem !important;
    letter-spacing: .10em !important;
    text-transform: uppercase !important;
    padding: 4px 0 !important;
    box-shadow: 0 8px 22px rgba(56,189,248,.28), inset 0 1px 0 rgba(255,255,255,.35) !important;
    transition: transform .18s ease, box-shadow .22s ease, filter .2s ease !important;
}
.stFormSubmitButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 14px 32px rgba(56,189,248,.42), inset 0 1px 0 rgba(255,255,255,.45) !important;
    filter: brightness(1.05) !important;
}
.stFormSubmitButton > button *,
.stFormSubmitButton button p,
[data-testid="stFormSubmitButton"] button p,
[data-testid="stBaseButton-primary"] p {
    color: #04070f !important;
    filter: none !important;
}

/* ── TABS ── */
.stTabs [data-baseweb="tab-list"] {
    background: rgba(8,13,24,.76) !important;
    gap: 6px !important;
    border: 1px solid rgba(56,189,248,.13) !important;
    border-radius: 10px !important;
    padding: 6px !important;
    box-shadow: inset 0 1px 0 rgba(255,255,255,.035), var(--shadow-soft) !important;
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
    background: rgba(56,189,248,.08) !important;
    border-color: rgba(56,189,248,.16) !important;
}
.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg, rgba(56,189,248,.17), rgba(167,139,250,.07)) !important;
    color: var(--accent) !important;
    border-color: rgba(56,189,248,.35) !important;
    box-shadow: 0 0 18px rgba(56,189,248,.12) !important;
}

/* ── DATAFRAME ── */
[data-testid="stDataFrame"] {
    border-radius: 10px !important;
    overflow: hidden !important;
    border: 1px solid rgba(56,189,248,.14) !important;
    box-shadow: var(--shadow-soft) !important;
    background: rgba(8,13,24,.84) !important;
}

/* ── METRICS ── */
[data-testid="stMetric"] {
    position: relative !important;
    overflow: hidden !important;
    background:
        linear-gradient(145deg, rgba(15,25,41,.92), rgba(8,13,24,.82)) !important;
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
    box-shadow: 0 0 16px rgba(56,189,248,.75);
}
[data-testid="stMetricLabel"] {
    font-size: .68rem !important;
    font-family: var(--mono) !important;
    text-transform: uppercase !important;
    letter-spacing: .1em !important;
    color: #8aa4c2 !important;
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
    border: 1px solid rgba(56,189,248,.14) !important;
    border-left-width: 4px !important;
    font-size: .875rem !important;
    background: rgba(8,13,24,.82) !important;
    box-shadow: var(--shadow-soft) !important;
}

/* ── FILE UPLOADER ── */
[data-testid="stFileUploader"] {
    border: 1px dashed rgba(56,189,248,.26) !important;
    border-radius: 10px !important;
    background:
        linear-gradient(135deg, rgba(56,189,248,.06), rgba(167,139,250,.035)),
        rgba(8,13,24,.72) !important;
    transition: all .2s ease !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: var(--accent) !important;
    background: rgba(56,189,248,.03) !important;
}

/* ── DOWNLOAD BUTTON ── */
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(180deg, rgba(15,25,41,.88), rgba(8,13,24,.82)) !important;
    border: 1px solid rgba(148,163,184,.16) !important;
    border-radius: 8px !important;
    color: #cbd5e1 !important;
    font-family: var(--font) !important;
    font-weight: 650 !important;
    font-size: .8rem !important;
    transition: all .2s ease !important;
}
[data-testid="stDownloadButton"] > button:hover {
    border-color: var(--border-hi) !important;
    color: var(--accent) !important;
    box-shadow: 0 0 22px rgba(56,189,248,.12) !important;
}

/* ── EXPANDERS / DIALOGS / STATUS ── */
[data-testid="stExpander"] {
    border: 1px solid rgba(56,189,248,.14) !important;
    border-radius: 10px !important;
    overflow: hidden !important;
    background: rgba(8,13,24,.78) !important;
    box-shadow: var(--shadow-soft) !important;
}
[data-testid="stExpander"] details {
    background: transparent !important;
}
[data-testid="stExpander"] summary {
    font-weight: 700 !important;
    color: var(--txt-1) !important;
}
[data-testid="stDialog"] > div {
    background:
        linear-gradient(145deg, rgba(15,25,41,.98), rgba(4,7,15,.96)) !important;
    border: 1px solid rgba(56,189,248,.22) !important;
    border-radius: 12px !important;
    box-shadow: 0 28px 80px rgba(0,0,0,.58), 0 0 0 1px rgba(56,189,248,.06) !important;
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
    border-bottom: 1px solid rgba(56,189,248,.12);
    background:
        radial-gradient(circle at 50% 0%, rgba(56,189,248,.22), transparent 42%),
        linear-gradient(180deg, rgba(15,25,41,.88), rgba(8,13,24,.18));
    isolation: isolate;
}
.sidebar-header::before {
    content: "";
    position: absolute;
    inset: 0;
    pointer-events: none;
    background:
        linear-gradient(90deg, transparent, rgba(56,189,248,.11), transparent) 0 0 / 100% 1px no-repeat,
        radial-gradient(circle at 50% 36%, rgba(167,139,250,.12), transparent 34%);
}
.sidebar-header::after {
    content: "";
    position: absolute;
    inset: -40% -60%;
    pointer-events: none;
    z-index: 0;
    background:
        radial-gradient(circle at 35% 20%, rgba(56,189,248,.10), transparent 45%),
        radial-gradient(circle at 70% 45%, rgba(167,139,250,.08), transparent 42%),
        linear-gradient(180deg, transparent, rgba(56,189,248,.06), transparent);
    transform: translate3d(-2%, 0, 0);
    animation: headerAtmosphere 10.5s ease-in-out infinite;
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
    animation: brandFloat 7.5s ease-in-out infinite;
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
    background: rgba(56,189,248,.7);
    box-shadow:
        22px 16px 0 rgba(56,189,248,.40),  44px  7px 0 rgba(167,139,250,.32),
        78px 22px 0 rgba(56,189,248,.28),  108px 11px 0 rgba(56,189,248,.44),
        138px 20px 0 rgba(167,139,250,.28), 168px  5px 0 rgba(56,189,248,.38),
        198px 24px 0 rgba(56,189,248,.22),  34px 44px 0 rgba(56,189,248,.20),
        68px 38px 0 rgba(167,139,250,.24), 104px 48px 0 rgba(56,189,248,.28),
        154px 36px 0 rgba(56,189,248,.20),  62px 58px 0 rgba(167,139,250,.18);
    animation: particleFloat 20s ease-in-out infinite;
}
.header-particles::after {
    content: "";
    position: absolute;
    width: 1px; height: 1px; border-radius: 50%;
    background: rgba(167,139,250,.6);
    box-shadow:
        14px 32px 0 rgba(56,189,248,.28),   54px 14px 0 rgba(167,139,250,.38),
        88px 36px 0 rgba(56,189,248,.22),  124px  8px 0 rgba(56,189,248,.32),
        162px 28px 0 rgba(167,139,250,.28), 196px 18px 0 rgba(56,189,248,.36),
        228px 42px 0 rgba(56,189,248,.18),   40px 54px 0 rgba(167,139,250,.22);
    animation: particleFloat 28s ease-in-out infinite reverse;
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
    border: 1px solid rgba(56,189,248,.22);
    animation: logoPing 3.5s ease-out infinite;
}
.logo-ring::after {
    content: "";
    position: absolute; inset: 6px;
    border-radius: 10px;
    border: 1px solid rgba(56,189,248,.10);
    animation: logoPing 3.5s ease-out infinite .9s;
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
    animation: orbitSpin 9s linear infinite;
    z-index: 2;
}
.logo-orbit::before {
    content: "";
    position: absolute; inset: 0;
    border-radius: 20px;
    border: 1.5px solid transparent;
    border-top-color: rgba(56,189,248,.62);
    border-right-color: rgba(56,189,248,.18);
}
.logo-orbit::after {
    content: "";
    position: absolute;
    top: -3px; left: 50%;
    width: 6px; height: 6px; margin-left: -3px;
    border-radius: 50%;
    background: #38bdf8;
    box-shadow: 0 0 8px 3px rgba(56,189,248,.88), 0 0 18px 6px rgba(56,189,248,.30);
}
/* ── ORBIT 2 — anel satélite roxo (anti-horário) ── */
.logo-orbit-2 {
    position: absolute; inset: -5px;
    border-radius: 16px;
    pointer-events: none;
    animation: orbitSpin 15s linear infinite reverse;
    z-index: 2;
}
.logo-orbit-2::before {
    content: "";
    position: absolute; inset: 0;
    border-radius: 16px;
    border: 1px solid transparent;
    border-bottom-color: rgba(167,139,250,.50);
    border-left-color: rgba(167,139,250,.14);
}
.logo-orbit-2::after {
    content: "";
    position: absolute;
    bottom: -2px; left: 50%;
    width: 4px; height: 4px; margin-left: -2px;
    border-radius: 50%;
    background: #a78bfa;
    box-shadow: 0 0 6px 2px rgba(167,139,250,.88), 0 0 14px 4px rgba(167,139,250,.28);
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
    background: linear-gradient(135deg, rgba(56,189,248,.14), rgba(99,102,241,.10));
    border-radius: 12px;
    border: 1px solid rgba(56,189,248,.25);
    box-shadow: 0 0 20px rgba(56,189,248,.12), inset 0 1px 0 rgba(255,255,255,.06);
    overflow: hidden;
    /* Camada GPU própria: confina o repaint do box-shadow pulsante a este layer,
       sem cascatear para a sidebar durante reruns. */
    transform: translateZ(0);
    will-change: transform;
    animation: markGlow 5s ease-in-out infinite;
}
/* scan line que varre o logo de cima a baixo */
.logo-mark::before {
    content: "";
    position: absolute;
    left: 0; right: 0; height: 40%;
    background: linear-gradient(
        180deg,
        transparent 0%,
        rgba(56,189,248,.16) 40%,
        rgba(56,189,248,.30) 50%,
        rgba(56,189,248,.16) 60%,
        transparent 100%
    );
    animation: logoScan 3.5s ease-in-out infinite;
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
        box-shadow: 0 0 16px rgba(56,189,248,.10), inset 0 1px 0 rgba(255,255,255,.06);
        border-color: rgba(56,189,248,.25);
        background: linear-gradient(135deg, rgba(56,189,248,.14), rgba(99,102,241,.10));
    }
    38% {
        box-shadow: 0 0 32px rgba(56,189,248,.32), 0 0 64px rgba(56,189,248,.10), inset 0 1px 0 rgba(255,255,255,.10);
        border-color: rgba(56,189,248,.48);
        background: linear-gradient(135deg, rgba(56,189,248,.22), rgba(99,102,241,.14));
    }
    68% {
        box-shadow: 0 0 28px rgba(167,139,250,.24), inset 0 1px 0 rgba(255,255,255,.10);
        border-color: rgba(167,139,250,.40);
        background: linear-gradient(135deg, rgba(99,102,241,.16), rgba(167,139,250,.12));
    }
}
.logo-mark svg { width: 60%; height: 60%; overflow: visible; position: relative; z-index: 1; }

/* ── CIRCUIT NODE ICON — pulso de rede tecnológico ── */
.circuit-node {
    animation: nodeGlow 3.5s ease-in-out infinite;
}
@keyframes nodeGlow {
    0%,100% { opacity: .85; }
    50%      { opacity: 1; }
}
.circuit-pulse-ring {
    animation: pulseRingExpand 2.8s ease-out infinite;
    transform-origin: center;
}
.circuit-pulse-ring-2 {
    animation: pulseRingExpand 2.8s ease-out infinite 1.4s;
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
    animation: drawLine 2.4s ease-in-out infinite;
}
.circuit-line-v {
    stroke-dasharray: 16;
    stroke-dashoffset: 16;
    animation: drawLine 2.4s ease-in-out infinite .6s;
}
.circuit-line-d {
    stroke-dasharray: 10;
    stroke-dashoffset: 10;
    animation: drawLine 2.4s ease-in-out infinite 1.2s;
}
@keyframes drawLine {
    0%,100% { stroke-dashoffset: 16; opacity: 0; }
    20%      { opacity: 1; }
    55%      { stroke-dashoffset: 0;  opacity: .9; }
    85%      { stroke-dashoffset: 0;  opacity: .3; }
}
.circuit-dot-1 { animation: dotBlink 2.4s ease-in-out infinite .3s; }
.circuit-dot-2 { animation: dotBlink 2.4s ease-in-out infinite .9s; }
.circuit-dot-3 { animation: dotBlink 2.4s ease-in-out infinite 1.5s; }
.circuit-dot-4 { animation: dotBlink 2.4s ease-in-out infinite 1.8s; }
@keyframes dotBlink {
    0%,100% { opacity: .18; r: 1.2; }
    45%      { opacity: 1;   r: 2.0; filter: drop-shadow(0 0 5px rgba(56,189,248,1)); }
    70%      { opacity: .7;  r: 1.6; }
}
.sidebar-title {
    color: var(--txt-1);
    font-family: var(--font);
    font-weight: 800;
    font-size: clamp(.86rem, 1.6vw, 1.02rem);
    letter-spacing: .14em;
    text-transform: uppercase;
    text-shadow: 0 0 18px rgba(56,189,248,.18);
    text-align: center;
    background: linear-gradient(90deg, #e2e8f0 0%, rgba(56,189,248,.95) 35%, #e2e8f0 70%, rgba(167,139,250,.9) 100%);
    background-size: 220% 100%;
    -webkit-background-clip: text;
    background-clip: text;
    -webkit-text-fill-color: transparent;
    animation: titleSheen 9s ease-in-out infinite;
}
.sidebar-sub {
    color: #6f86a3;
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
    border: 1px solid rgba(56,189,248,.22);
    box-shadow: 0 0 40px rgba(56,189,248,.07), var(--shadow-soft);
    border-radius: 10px;
    background: rgba(4,7,15,.45);
    overflow: hidden;
    color: #8aa4c2;
    font-size: .64rem;
    font-family: var(--mono);
    text-transform: uppercase;
    letter-spacing: .08em;
    transition: transform .18s ease, border-color .18s ease, box-shadow .18s ease;
}
.status-strip:hover {
    transform: translateY(-1px);
    border-color: rgba(56,189,248,.35);
    box-shadow: 0 0 46px rgba(56,189,248,.10), 0 18px 44px rgba(0,0,0,.38);
}
.status-strip::after {
    content: "";
    position: absolute;
    inset: 0;
    pointer-events: none;
    background: linear-gradient(
        90deg,
        transparent 0%,
        rgba(56,189,248,.10) 35%,
        rgba(167,139,250,.08) 50%,
        rgba(56,189,248,.10) 65%,
        transparent 100%
    );
    transform: translateX(-110%);
    opacity: .55;
    animation: statusShimmer 6.5s ease-in-out infinite;
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
    box-shadow: 0 0 14px rgba(52,211,153,.8);
    animation: onlinePulse 1.2s ease-in-out infinite;
}

.status-strip strong {
    color: #d7f9ff;
    font-weight: 700;
}
.status-strip em {
    color: var(--accent);
    font-style: normal;
    font-weight: 700;
    font-variant-numeric: tabular-nums;
    animation: onlinePulse 1.2s ease-in-out infinite, uptimeFlicker 9s ease-in-out infinite;
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
    color: #7f95b2 !important;
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
    box-shadow: 0 0 14px rgba(56,189,248,.75);
    animation: navGuidePulse 2.8s ease-in-out infinite;
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
    background: linear-gradient(180deg, transparent, rgba(56,189,248,.28), rgba(167,139,250,.18), transparent);
    box-shadow: 0 0 18px rgba(56,189,248,.18);
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
    background:
        radial-gradient(circle at 0% 50%, rgba(56,189,248,.18), transparent 30%),
        linear-gradient(110deg, transparent 0%, rgba(56,189,248,.13) 42%, rgba(167,139,250,.11) 50%, transparent 62%);
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
    background: linear-gradient(90deg, rgba(56,189,248,.95), rgba(167,139,250,.80), transparent);
    box-shadow: 0 0 10px rgba(56,189,248,.50);
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
    animation: navActiveBreath 4.8s ease-in-out infinite !important;
}
[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked)::before {
    width: 4px !important;
    animation: navRailGlow 2.2s ease-in-out infinite !important;
}
[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked)::after {
    animation: navDotOrbit 2.6s ease-in-out infinite !important;
}
@keyframes navGuidePulse {
    0%,100% { opacity: .45; transform: scale(.92); }
    50% { opacity: 1; transform: scale(1.18); }
}
@keyframes navRailGlow {
    0%,100% { box-shadow: 0 0 10px rgba(56,189,248,.70), 0 0 28px rgba(56,189,248,.20); }
    50% { box-shadow: 0 0 16px rgba(56,189,248,1), 0 0 46px rgba(167,139,250,.24); }
}
@keyframes navDotOrbit {
    0%,100% { transform: translateY(-50%) scale(1.12); box-shadow: 0 0 10px rgba(56,189,248,.75); }
    50% { transform: translateY(-50%) scale(1.42); box-shadow: 0 0 18px rgba(56,189,248,1), 0 0 34px rgba(167,139,250,.28); }
}
@keyframes navActiveBreath {
    0%,100% { box-shadow: 0 0 0 1px rgba(56,189,248,.16), 0 0 24px rgba(56,189,248,.20), 0 16px 38px rgba(56,189,248,.12); }
    50% { box-shadow: 0 0 0 1px rgba(56,189,248,.24), 0 0 34px rgba(56,189,248,.30), 0 18px 46px rgba(167,139,250,.12); }
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
    border: 1px solid rgba(56,189,248,.14);
    border-radius: 10px;
    background:
        linear-gradient(135deg, rgba(56,189,248,.10), rgba(167,139,250,.055)),
        rgba(8,13,24,.82);
    box-shadow: inset 0 1px 0 rgba(255,255,255,.035);
    /* Camada GPU própria — isola o repaint do pulse de box-shadow/borda. */
    transform: translateZ(0);
    will-change: transform;
    animation: navContextPulse 5s ease-in-out infinite;
}
@keyframes navContextPulse {
    0%,100% { border-color: rgba(56,189,248,.14); box-shadow: inset 0 1px 0 rgba(255,255,255,.035); }
    50%      { border-color: rgba(56,189,248,.30); box-shadow: inset 0 1px 0 rgba(255,255,255,.035), 0 0 20px rgba(56,189,248,.08); }
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
    color: #7d8fa8;
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
    border-color: rgba(56,189,248,.20);
    box-shadow: inset 0 1px 0 rgba(255,255,255,.04), 0 12px 28px rgba(0,0,0,.20);
}
.nav-context::before {
    content: "";
    position: absolute;
    inset: 0;
    pointer-events: none;
    background:
        radial-gradient(circle at 10% 18%, rgba(56,189,248,.14), transparent 34%),
        linear-gradient(105deg, transparent, rgba(56,189,248,.10), rgba(167,139,250,.08), transparent);
    transform: translateX(-64%);
    animation: navContextScan 5.2s ease-in-out infinite;
}
.nav-context::after {
    content: "";
    position: absolute;
    left: 14px;
    right: 14px;
    top: 0;
    height: 1px;
    background: linear-gradient(90deg, transparent, rgba(56,189,248,.82), rgba(167,139,250,.52), transparent);
    box-shadow: 0 0 16px rgba(56,189,248,.36);
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
    background: rgba(56,189,248,.88);
    box-shadow: 0 0 10px rgba(56,189,248,.65);
    animation: signalStep 1.8s ease-in-out infinite;
}
.nav-context-signal span:nth-child(2) {
    background: rgba(125,211,252,.78);
    animation-delay: .18s;
}
.nav-context-signal span:nth-child(3) {
    background: rgba(167,139,250,.82);
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
    color: #6f86a3;
    border: 1px solid rgba(56,189,248,.12);
    border-radius: 10px;
    background: rgba(8,13,24,.92);
    box-shadow: 0 12px 30px rgba(0,0,0,.22), inset 0 1px 0 rgba(255,255,255,.03);
    animation: footerPulse 9s ease-in-out infinite;
}
@keyframes footerPulse {
    0%,100% { border-color: rgba(56,189,248,.12); box-shadow: 0 12px 30px rgba(0,0,0,.22), inset 0 1px 0 rgba(255,255,255,.03); }
    50%      { border-color: rgba(56,189,248,.26); box-shadow: 0 12px 30px rgba(0,0,0,.22), 0 0 22px rgba(56,189,248,.07), inset 0 1px 0 rgba(255,255,255,.03); }
}
.sidebar-footer a { color: var(--accent) !important; font-weight: 700; }
.sidebar-footer small {
    display: block;
    margin-top: 3px;
    color: #53667f;
    font-size: .64rem;
}

/* Page hero */
.page-hero {
    position: relative;
    overflow: hidden;
    padding: 22px 26px 22px 28px;
    margin: 0 0 24px;
    border: 1px solid rgba(56,189,248,.09);
    border-radius: 10px;
    background: #080f1c;
    box-shadow: none;
}
.page-hero::before {
    content: "";
    position: absolute;
    left: 0; top: 0; bottom: 0;
    width: 2px;
    pointer-events: none;
    background: #38bdf8;
    box-shadow: 0 0 10px rgba(56,189,248,.5);
}
.page-hero::after {
    position: absolute;
    right: 22px;
    top: 18px;
    color: rgba(125,143,168,.16);
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
    color: rgba(125,211,252,.75); margin-bottom: 9px;
}
.page-hero-eyebrow::before {
    content: "";
    width: 7px; height: 7px;
    border-radius: 999px;
    background: var(--green);
    box-shadow: 0 0 14px rgba(52,211,153,.8);
}
.page-hero-title {
    position: relative;
    z-index: 1;
    font-size: clamp(1.15rem, 1.6vw, 1.5rem); font-weight: 600;
    color: #f1f5f9; line-height: 1.2;
    margin: 0 0 5px;
    letter-spacing: -.005em;
}
.page-hero-desc {
    position: relative;
    z-index: 1;
    max-width: 680px;
    font-size: .9rem;
    line-height: 1.6;
    color: rgba(154,173,198,.82);
    margin: 0;
}

/* Section label */
.sec-label {
    display: flex; align-items: center; gap: 9px;
    font-size: .72rem; font-weight: 600;
    font-family: var(--mono);
    text-transform: uppercase; letter-spacing: .05em;
    color: #93a7c2; padding-bottom: 12px;
    border-bottom: 1px solid rgba(56,189,248,.07);
    margin: 6px 0 18px;
}
.sec-dot {
    width: 6px; height: 6px; border-radius: 50%;
    background: var(--accent);
    box-shadow: 0 0 14px var(--accent), 0 0 28px rgba(56,189,248,.4);
    flex-shrink: 0;
}

/* KPI cards */
.kpi-wrap { display: grid; grid-template-columns: repeat(4,1fr); gap: 14px; margin-bottom: 24px; }
.kpi-card {
    background:
        linear-gradient(145deg, rgba(15,25,41,.92), rgba(8,13,24,.80)) !important;
    border: 1px solid rgba(148,163,184,.14) !important;
    border-radius: var(--radius);
    padding: 18px 20px;
    position: relative; overflow: hidden;
    box-shadow: var(--shadow-soft);
    transition: border-color .2s, transform .2s, box-shadow .2s;
}
.kpi-card:hover {
    border-color: var(--border-hi) !important;
    transform: translateY(-2px);
    box-shadow: 0 18px 44px rgba(0,0,0,.45), 0 0 30px rgba(56,189,248,.15), 0 0 0 1px rgba(56,189,248,.12);
}
.kpi-card::before {
    content: ''; position: absolute;
    top: 0; left: 0; bottom: 0; width: 3px;
    background: var(--c, var(--accent));
    box-shadow: 0 0 22px var(--c, var(--accent)), 0 0 50px rgba(56,189,248,.15);
}
.kpi-card::after {
    content: "";
    position: absolute;
    inset: 0;
    background:
        linear-gradient(
            135deg,
            rgba(255,255,255,0.03) 0%,
            rgba(255,255,255,0.01) 30%,
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
    letter-spacing: .1em; color: #8aa4c2; font-weight: 700;
}

/* Status pills */
.pill {
    display: inline-flex; align-items: center; gap: 5px;
    padding: 4px 12px; border-radius: 999px;
    font-size: var(--text-xs); font-weight: 500;
    font-family: var(--font);
    letter-spacing: .01em; white-space: nowrap;
}
.pill-gray  { background: rgba(148,163,184,.14); color: #b6c2d2; }
.pill-amber { background: rgba(251,191,36,.14);  color: #fbbf24; }
.pill-green { background: rgba(52,211,153,.14);  color: #34d399; }
.pill-red   { background: rgba(248,113,113,.14); color: #f87171;
    animation: pillBreath 2.8s ease-in-out infinite; }
@keyframes pillBreath {
    0%, 100% { box-shadow: 0 0 0 0 rgba(248,113,113,0); }
    50%      { box-shadow: 0 0 0 3px rgba(248,113,113,.10), 0 0 12px rgba(248,113,113,.20); }
}
.pill-blue  { background: rgba(56,189,248,.14);  color: #7dd3fc; }

/* Ticket confirm card */
.confirm-card {
    position: relative;
    overflow: hidden;
    background:
        linear-gradient(135deg, rgba(56,189,248,.10), rgba(167,139,250,.045)),
        rgba(8,13,24,.82);
    border: 1px solid rgba(56,189,248,.22);
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
    box-shadow: 0 0 16px rgba(52,211,153,.75);
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
    border: 1px solid rgba(56,189,248,.14);
    border-radius: 10px;
    background: linear-gradient(145deg, rgba(15,25,41,.80), rgba(8,13,24,.68));
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
    background: rgba(56,189,248,.12);
    border: 1px solid rgba(56,189,248,.22);
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
    color: #8aa4c2 !important;
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
    color: #d7e3f4 !important;
    font-weight: 700 !important;
}

/* Helper/caption text less prominent */
[data-testid="stCaptionContainer"] p,
[data-testid="stCaptionContainer"],
.stCaption,
.st-emotion-cache-1v0mbdj, 
.st-emotion-cache-1wmy9hl {
    color: rgba(138,164,194,.56) !important;
    font-size: .72rem !important;
}

/* Placeholders more subtle */
input::placeholder,
textarea::placeholder {
    color: rgba(138,164,194,.46) !important;
    opacity: .75 !important;
}

/* File uploader helper line */
[data-testid="stFileUploader"] small,
[data-testid="stFileUploader"] [data-testid="stMarkdownContainer"] p {
    color: rgba(138,164,194,.55) !important;
    font-size: .72rem !important;
}

/* Dataframes: more dashboard, less Excel */
div[data-testid="stDataFrame"] {
    border-radius: 12px !important;
    overflow: hidden !important;
    border: 1px solid rgba(56,189,248,.14) !important;
    background: rgba(8,13,24,.55) !important;
    box-shadow: 0 18px 44px rgba(0,0,0,.28) !important;
}
div[data-testid="stDataFrame"] [data-baseweb="data-table"] {
    border-radius: 12px !important;
}
div[data-testid="stDataFrame"] [role="columnheader"] {
    background: linear-gradient(180deg, rgba(15,25,41,.92), rgba(8,13,24,.70)) !important;
    border-bottom: 1px solid rgba(56,189,248,.14) !important;
}
div[data-testid="stDataFrame"] [role="columnheader"] * {
    font-family: var(--mono) !important;
    font-size: .62rem !important;
    letter-spacing: .14em !important;
    text-transform: uppercase !important;
    color: rgba(138,164,194,.78) !important;
}
div[data-testid="stDataFrame"] [role="row"] {
    border-bottom: 1px solid rgba(148,163,184,.08) !important;
}
div[data-testid="stDataFrame"] [role="gridcell"] {
    font-size: .78rem !important;
    color: rgba(226,232,240,.92) !important;
}
div[data-testid="stDataFrame"] [role="row"]:nth-child(even) [role="gridcell"] {
    background: rgba(255,255,255,.012) !important;
}
div[data-testid="stDataFrame"] [role="row"]:hover [role="gridcell"] {
    background: rgba(56,189,248,.05) !important;
}

/* Tabs: more product, less default */
div[data-testid="stTabs"] [role="tablist"] {
    gap: 8px !important;
    padding: 6px 6px 8px !important;
    border-radius: 12px !important;
    background: rgba(8,13,24,.55) !important;
    border: 1px solid rgba(56,189,248,.12) !important;
}
div[data-testid="stTabs"] button[role="tab"] {
    border-radius: 10px !important;
    padding: 10px 12px !important;
    background: rgba(15,25,41,.55) !important;
    border: 1px solid rgba(148,163,184,.12) !important;
    transition: transform .18s ease, border-color .18s ease, background .18s ease, box-shadow .18s ease !important;
}
div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
    border-color: rgba(56,189,248,.35) !important;
    background: linear-gradient(135deg, rgba(56,189,248,.10), rgba(167,139,250,.055)) !important;
    box-shadow: 0 0 24px rgba(56,189,248,.10) !important;
}
div[data-testid="stTabs"] button[role="tab"]:hover {
    transform: translateY(-1px) !important;
    border-color: rgba(56,189,248,.22) !important;
}
.ux-note {
    display: flex;
    align-items: flex-start;
    gap: 10px;
    padding: 10px 11px;
    border: 1px solid rgba(56,189,248,.16);
    border-radius: 10px;
    background:
        linear-gradient(135deg, rgba(56,189,248,.08), rgba(167,139,250,.035)),
        rgba(8,13,24,.72);
    color: #9aadc6;
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
    box-shadow: 0 0 10px rgba(56,189,248,.55);
}
.form-action-note {
    min-height: 40px;
    display: flex;
    align-items: center;
    padding: 8px 10px;
    border: 1px solid rgba(148,163,184,.12);
    border-radius: 9px;
    background: rgba(8,13,24,.62);
    color: #8aa4c2;
    font-size: var(--text-sm);
    line-height: 1.35;
}
.empty-state {
    padding: 28px 22px;
    text-align: center;
    border: 1px dashed rgba(56,189,248,.16);
    border-radius: 18px;
    background:
        linear-gradient(135deg, rgba(56,189,248,.07), rgba(167,139,250,.035)),
        rgba(8,13,24,.66);
}
.empty-state strong {
    display: block;
    color: var(--txt-1);
    font-size: var(--text-md);
    margin-bottom: 5px;
}
.empty-state span {
    color: #8aa4c2 !important;
    font-size: var(--text-sm);
}
.search-panel {
    margin-bottom: 18px;
    padding: 14px 16px;
    border: 1px solid rgba(56,189,248,.14);
    border-radius: 10px;
    background: rgba(8,13,24,.70);
}
.queue-card {
    position: relative;
    overflow: hidden;
    border: 1px solid rgba(148,163,184,.13);
    border-radius: 10px;
    padding: 10px 12px;
    background: linear-gradient(145deg, rgba(15,25,41,.82), rgba(8,13,24,.72));
    box-shadow: 0 14px 32px rgba(0,0,0,.40);
}
.queue-card::before {
    content: "";
    position: absolute;
    left: 0; top: 0; bottom: 0;
    width: 3px;
    background: var(--queue-accent, var(--accent));
    box-shadow: 0 0 12px var(--queue-accent, var(--accent));
}
.queue-card + .queue-card { margin-top: 10px; }
.queue-meta {
    color: #8aa4c2;
    font-size: .74rem;
    font-family: var(--mono);
    letter-spacing: .04em;
    text-transform: uppercase;
}
.queue-critical { --queue-accent: #f87171; box-shadow: -3px 0 20px rgba(248,113,113,.22); }
.queue-high { --queue-accent: #fbbf24; }
.queue-medium { --queue-accent: #38bdf8; }
.queue-low { --queue-accent: #34d399; }
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
    color: #8aa4c2;
    font-size: .78rem;
    margin-top: 2px;
}
.work-card-desc {
    color: #cbd5e1;
    font-size: .84rem;
    line-height: 1.45;
    margin: 9px 0 12px;
}
/* linha de metadados compacta dentro do card (substitui a fileira de chips) */
.work-card-meta {
    display: flex; flex-wrap: wrap; gap: 6px 16px; align-items: center;
    margin-top: 10px; padding-top: 10px;
    border-top: 1px solid rgba(148,163,184,.09);
}
.work-card-meta span {
    font-size: .74rem; color: #8194ab; display: inline-flex; gap: 5px; align-items: center;
}
.work-card-meta strong { color: #cbd5e1; font-weight: 700; }
/* linha enxuta de arquivo (histórico de concluídos) */
.arc-row { display: flex; align-items: center; gap: 8px 14px; flex-wrap: wrap; }
.arc-row .arc-id {
    font-family: 'Geist Mono', ui-monospace, monospace; font-size: .60rem; font-weight: 700;
    letter-spacing: .10em; text-transform: uppercase; color: #5fa8d8;
}
.arc-row .arc-user { font-weight: 700; color: #e2e8f0; font-size: .85rem; }
.arc-row .arc-meta { color: #8194ab; font-size: .76rem; }
.arc-row .arc-date { color: #6f8197; font-size: .72rem; margin-left: auto; white-space: nowrap; }
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
    border: 1px solid rgba(148,163,184,.12);
    border-radius: 9px;
    background: rgba(8,13,24,.64);
    color: #9aadc6;
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
    border-bottom: 1px solid rgba(148,163,184,.14);
}
.history-panel-title {
    display: flex;
    align-items: center;
    gap: 10px;
    color: #f8fafc;
    font-size: 1.16rem;
    font-weight: 850;
    letter-spacing: 0;
}
.history-title-dot {
    width: 13px;
    height: 13px;
    border-radius: 999px;
    background: #34d399;
    box-shadow: 0 0 18px rgba(52,211,153,.75);
}
.history-count {
    display: inline-flex;
    align-items: center;
    min-height: 30px;
    padding: 5px 10px;
    border: 1px solid rgba(56,189,248,.18);
    border-radius: 999px;
    background: rgba(56,189,248,.08);
    color: #bae6fd;
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
    color: #f8fafc;
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
    border: 1px solid rgba(148,163,184,.16);
    white-space: nowrap;
}
.history-prio-critical {
    color: #fecaca;
    background: rgba(248,113,113,.14);
    border-color: rgba(248,113,113,.34);
}
.history-prio-medium {
    color: #fef3c7;
    background: rgba(251,191,36,.13);
    border-color: rgba(251,191,36,.30);
}
.history-prio-low {
    color: #bbf7d0;
    background: rgba(52,211,153,.12);
    border-color: rgba(52,211,153,.28);
}
.history-user {
    color: #f8fafc;
    font-size: 1rem;
    font-weight: 850;
    line-height: 1.2;
}
.history-meta {
    margin-top: 5px;
    color: #a7bed8;
    font-size: .80rem;
    font-weight: 650;
}
.history-desc {
    margin-top: 7px;
    color: #dbeafe;
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
    background: rgba(52,211,153,.12);
    border: 1px solid rgba(52,211,153,.28);
    color: #bbf7d0;
    font-weight: 850;
    font-size: .82rem;
}
.history-date {
    margin-top: 8px;
    color: #cbd5e1;
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
    border: 1px solid rgba(56,189,248,.12);
    border-radius: 10px;
    background: linear-gradient(145deg, rgba(12,24,40,.94), rgba(6,11,21,.84));
    box-shadow: var(--shadow-soft), inset 0 1px 0 rgba(255,255,255,.035);
}
.ops-card::before {
    content: "";
    position: absolute;
    inset: 0 0 auto;
    height: 2px;
    background: var(--tone, #38bdf8);
    box-shadow: 0 0 18px var(--tone, #38bdf8);
}
.ops-label {
    color: #8aa4c2;
    font-family: var(--mono);
    font-size: .66rem;
    font-weight: 800;
    letter-spacing: .1em;
    text-transform: uppercase;
}
.ops-value {
    margin-top: 8px;
    color: #f8fbff;
    font-family: var(--mono);
    font-size: 1.65rem;
    font-weight: 800;
    line-height: 1;
}
.ops-copy {
    margin-top: 8px;
    color: #7d8fa8;
    font-size: .78rem;
    line-height: 1.35;
}
.insight-panel {
    border: 1px solid rgba(56,189,248,.16);
    border-radius: 10px;
    padding: 15px 16px;
    background:
        linear-gradient(135deg, rgba(56,189,248,.08), rgba(167,139,250,.04)),
        rgba(8,13,24,.72);
    color: #b7c7dd;
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
    border: 1px solid rgba(148,163,184,.12);
    background: rgba(8,13,24,.62);
}
.kanban-lane strong {
    display: block;
    color: #e2e8f0;
    font-size: .88rem;
    margin-bottom: 7px;
}
.kanban-lane span {
    color: #7d8fa8 !important;
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
    border: 1px solid rgba(148,163,184,.14);
    background: rgba(8,13,24,.62);
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
        rgba(56,189,248,.00) 0%,
        rgba(56,189,248,.10) 38%,
        rgba(167,139,250,.10) 55%,
        rgba(56,189,248,.00) 100%
    );
    animation: skeletonShimmer 1.15s ease-in-out infinite;
    pointer-events: none;
}
.skeleton-line {
    height: 12px;
    border-radius: 999px;
    background: rgba(148,163,184,.12);
    margin: 16px 14px;
}
.skeleton-line.sm { height: 9px; width: 72%; }
.skeleton-line.md { height: 10px; width: 86%; }
.skeleton-line.lg { height: 12px; width: 100%; }
.skeleton-label {
    color: rgba(138,164,194,.70);
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
    border: 1px solid rgba(148,163,184,.13);
    background: linear-gradient(145deg, rgba(15,25,41,.74), rgba(8,13,24,.64));
    box-shadow: 0 14px 32px rgba(0,0,0,.26);
    overflow: hidden;
    position: relative;
    padding: 10px 12px;
    min-height: 98px;
}
.skeleton-queue::before {
    content: "";
    position: absolute;
    inset: 0;
    background: radial-gradient(circle at 20% 0%, rgba(56,189,248,.10), transparent 45%);
    pointer-events: none;
}
.skeleton-queue .skeleton-block {
    height: 14px;
    border-radius: 999px;
    background: rgba(148,163,184,.12);
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
            box-shadow: inset -1px 0 0 rgba(255,255,255,.03), 12px 0 44px rgba(0,0,0,.26) !important;
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
            box-shadow: 0 0 20px rgba(56,189,248,.3) !important;
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

</style>
"""

# ═══════════════════════════════════════════════════════
# CSS GLOBAL — DESIGN SYSTEM PROFISSIONAL
# ═══════════════════════════════════════════════════════
# Fontes via <link> com preconnect: não bloqueiam as regras CSS inline
# (que aplicam imediatamente). Antes era um @import no topo do <style>,
# que travava todo o design até a fonte baixar — causava o FOUC.
st.markdown(
    """
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link rel="stylesheet"
          href="https://fonts.googleapis.com/css2?family=Geist:wght@300;400;500;600;700;800&family=Geist+Mono:wght@400;600&display=swap">
    """,
    unsafe_allow_html=True,
)
# CSS GLOBAL — injetado UMA ÚNICA VEZ por sessão direto no <head> do documento.
# PERFORMANCE: antes este bloco (~98 KB minificados) era reenviado pelo
# st.markdown a CADA rerun (todo clique/navegação), pois o Streamlit re-emite
# todos os elementos da página a cada interação. Agora o CSS é escrito no <head>
# pelo navegador, onde PERSISTE entre os reruns (o <head> fica fora da árvore
# React gerenciada pelo Streamlit). Resultado: o payload trafega só no primeiro
# carregamento da sessão; as navegações seguintes não recarregam o stylesheet.
# Em F5/refresh a sessão reinicia e o <head> recarrega → reinjeta normalmente.
def _injetar_css_global_uma_vez():
    # ANTI-FOUC: o CSS é renderizado de forma SÍNCRONA pelo próprio st.markdown,
    # como parte do HTML servido. Assim o estilo aplica no mesmo frame em que o
    # conteúdo aparece — sem a janela de "texto sem estilo" (FOUC/jitter) que
    # ocorria quando o CSS era empurrado para o <head> por um iframe assíncrono
    # (components.html), que só carregava DEPOIS do conteúdo já visível.
    # O payload é estático: o Streamlit usa o cache de mensagens (ForwardMsg)
    # para não retransmitir os ~98 KB a cada rerun — só trafega o hash.
    _css_html = _minify_css(_get_global_css())  # já cacheado (custo pago 1x)
    st.markdown(_css_html, unsafe_allow_html=True)

_injetar_css_global_uma_vez()

# ── SCROLLBAR INJECTION para iframes de dataframe (contexto isolado) ──
st.markdown("""
<style>
* { scrollbar-width: thin !important; scrollbar-color: rgba(56,189,248,.55) transparent !important; }
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
        * { scrollbar-width: thin; scrollbar-color: rgba(56,189,248,.55) transparent; }
        ::-webkit-scrollbar { width: 5px !important; height: 5px !important; }
        ::-webkit-scrollbar-track { background: rgba(6,11,21,.40) !important; border-radius: 99px !important; }
        ::-webkit-scrollbar-thumb {
            background: linear-gradient(180deg, #38bdf8 0%, rgba(56,189,248,.60) 100%) !important;
            border-radius: 99px !important;
            box-shadow: 0 0 6px rgba(56,189,248,.45) !important;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: #7dd3fc !important;
            box-shadow: 0 0 12px rgba(56,189,248,.70) !important;
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

# ── ECONOMIA DE CPU: pausa as animações infinitas (orbits, pulsos, gradientes)
# quando a aba/janela não está visível, e as desliga para quem prefere menos
# movimento. São ~55 animações em loop — pausá-las quando ociosas deixa a UI
# mais leve e responsiva ao voltar o foco.
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

# ── CSS de componentes (header/wizard/cards) injetado UMA vez, fora dos
# @st.fragment, para nao "piscar" texto cru a cada rerun de fragment ──
st.markdown("""
<style>
.wz{display:flex;align-items:center;padding:6px 0 20px;user-select:none}
.wz-node{display:flex;flex-direction:column;align-items:center;gap:5px;flex:0 0 auto}
.wz-circle{width:30px;height:30px;border-radius:50%;display:flex;align-items:center;
  justify-content:center;font-size:12px;font-weight:800;transition:all .25s}
.wz-c-done{background:#10b981;color:#fff;box-shadow:0 0 10px rgba(16,185,129,.45)}
.wz-c-act{background:#0d1e35;border:2px solid #38bdf8;color:#38bdf8;box-shadow:0 0 12px rgba(56,189,248,.3)}
.wz-c-idle{background:#0a1520;border:2px solid #132030;color:#1e3040}
.wz-lbl{font-size:10px;font-weight:700;letter-spacing:.02em;white-space:nowrap}
.wz-lbl-done{color:#34d399}
.wz-lbl-act{color:#cbd5e1}
.wz-lbl-idle{color:#1e3040}
.wz-line{flex:1;height:2px;margin:0 8px;margin-bottom:15px;border-radius:2px}
.wz-line-done{background:#10b981}
.wz-line-idle{background:#132030}
.hz{background:linear-gradient(135deg,#0a1828,#060c18);border:1px solid #132030;
  border-radius:12px;padding:18px 22px;position:relative;overflow:hidden;margin-bottom:14px;
  box-shadow:inset 0 1px 0 rgba(255,255,255,.05),inset 0 -6px 14px rgba(0,0,0,.40),0 10px 22px rgba(0,0,0,.40),0 2px 5px rgba(0,0,0,.30)}
.hz::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;
  background:linear-gradient(90deg,#38bdf8,#a78bfa);box-shadow:0 0 12px rgba(56,189,248,.4)}
.hz-eye{font-size:9px;font-weight:800;letter-spacing:.15em;text-transform:uppercase;color:#2a4a60;margin-bottom:5px}
.hz-title{font-size:1.1rem;font-weight:800;color:#e2e8f0;margin-bottom:3px}
.hz-desc{font-size:.78rem;color:#4a6a80;line-height:1.5}
.hz-crumb{font-size:10px;color:#3a5a70;margin-bottom:5px}
.hz-crumb b{color:#5a7a90}
.prio-chip{display:inline-flex;align-items:center;gap:7px;padding:6px 14px;
  border-radius:8px;font-size:11px;font-weight:700;margin:6px 0 2px}
.prio-dot{width:7px;height:7px;border-radius:50%;flex-shrink:0}
.char-c{font-size:9px;text-align:right;margin-top:3px}
.ok-card{background:linear-gradient(145deg,#071a10,#040d18);
  border:1.5px solid #1a4a2a;border-radius:14px;padding:28px 24px;text-align:center;
  box-shadow:inset 0 1px 0 rgba(255,255,255,.05),inset 0 -7px 16px rgba(0,0,0,.40),0 12px 26px rgba(0,0,0,.42),0 0 40px rgba(52,211,153,.12);margin-bottom:12px}
.ok-card::before{content:'';display:block;height:3px;
  background:linear-gradient(90deg,#10b981,#34d399);border-radius:12px 12px 0 0;margin:-28px -24px 24px}
.ok-proto{font-size:.72rem;font-family:monospace;background:#0a2015;
  border:1px solid #1a4a2a;color:#6ee7b7;padding:4px 14px;border-radius:7px;
  display:inline-block;margin:6px 0 12px;letter-spacing:.08em}
.ok-steps{display:grid;grid-template-columns:repeat(3,1fr);gap:7px;margin-top:14px}
.ok-s{background:linear-gradient(160deg,#0a2014,#050f0a);border:1px solid #1a3a20;border-radius:8px;padding:9px 7px;box-shadow:inset 0 1px 0 rgba(255,255,255,.05),inset 0 -3px 7px rgba(0,0,0,.40),0 4px 10px rgba(0,0,0,.35)}
.ok-sn{font-size:8px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:#34d399}
.ok-st{font-size:10px;font-weight:700;color:#3a7a50;margin-top:2px}
.wz-c-act{background:#0a1020;border:2px solid #38bdf8;color:#38bdf8;box-shadow:0 0 12px rgba(56,189,248,.3)}
.hzp{background:linear-gradient(135deg,#0a1020,#040810);border:1px solid #0a1020;
  border-radius:12px;padding:18px 22px;position:relative;overflow:hidden;margin-bottom:14px;
  box-shadow:inset 0 1px 0 rgba(255,255,255,.05),inset 0 -6px 14px rgba(0,0,0,.40),0 10px 22px rgba(0,0,0,.40),0 2px 5px rgba(0,0,0,.30)}
.hzp::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;
  background:linear-gradient(90deg,#38bdf8,#0369a1);box-shadow:0 0 12px rgba(56,189,248,.4)}
.hzp-eye{font-size:9px;font-weight:800;letter-spacing:.15em;text-transform:uppercase;color:#1a2a3a;margin-bottom:5px}
.hzp-title{font-size:1.1rem;font-weight:800;color:#e2e8f0;margin-bottom:3px}
.hzp-desc{font-size:.78rem;color:#3d5068;line-height:1.5}
.hzp-crumb{font-size:10px;color:#27384d;margin-bottom:5px}
.hzp-crumb b{color:#4a7a9a}
.prio-chip2{display:inline-flex;align-items:center;gap:7px;padding:6px 14px;
  border-radius:8px;font-size:11px;font-weight:700;margin:6px 0 2px}
.prio-dot2{width:7px;height:7px;border-radius:50%;flex-shrink:0}
.ok-cp{background:linear-gradient(145deg,#0a1020,#040810);border:1.5px solid #102844;
  border-radius:14px;padding:28px 24px;text-align:center;
  box-shadow:0 0 40px rgba(56,189,248,.1);margin-bottom:12px}
.ok-cp::before{content:'';display:block;height:3px;
  background:linear-gradient(90deg,#0369a1,#38bdf8);border-radius:12px 12px 0 0;margin:-28px -24px 24px}
.ok-cp-proto{font-size:.72rem;font-family:monospace;background:#0a1020;
  border:1px solid #102844;color:#7dd3fc;padding:4px 14px;border-radius:7px;
  display:inline-block;margin:6px 0 12px;letter-spacing:.08em}
.ctx-panel{background:linear-gradient(160deg,rgba(12,24,40,.92),rgba(6,11,21,.85));
          border:1px solid rgba(56,189,248,.14);border-radius:14px;padding:20px 20px 18px;
          box-shadow:inset 0 1px 0 rgba(255,255,255,.05),inset 0 -7px 16px rgba(0,0,0,.34),0 12px 28px rgba(0,0,0,.40),0 2px 6px rgba(0,0,0,.30);position:relative;overflow:hidden}
.ctx-panel::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;
          background:linear-gradient(90deg,#38bdf8,#0284c7);box-shadow:0 0 12px rgba(56,189,248,.4)}
.ctx-h{font-family:'Geist Mono',monospace;font-size:.62rem;font-weight:800;letter-spacing:.16em;
          text-transform:uppercase;color:#38bdf8;margin-bottom:16px}
.ctx-step{display:flex;gap:12px;align-items:flex-start;margin-bottom:14px}
.ctx-n{flex:0 0 24px;width:24px;height:24px;border-radius:50%;display:flex;align-items:center;
          justify-content:center;font-size:11px;font-weight:800;color:#38bdf8;
          background:rgba(56,189,248,.10);border:1px solid rgba(56,189,248,.3)}
.ctx-step strong{display:block;font-size:.84rem;color:#e2e8f0;font-weight:700;margin-bottom:2px}
.ctx-step span{font-size:.74rem;color:#7d8fa8;line-height:1.45}
.ctx-divider{height:1px;background:linear-gradient(90deg,transparent,rgba(56,189,248,.25),transparent);margin:4px 0 14px}
.ctx-tip{font-size:.76rem;color:#bae6fd;background:rgba(56,189,248,.06);
          border:1px solid rgba(56,189,248,.13);border-radius:9px;padding:10px 12px;line-height:1.5}
.cat3d:hover{
          transform:translateY(-3px)!important;
          border-color:rgba(56,189,248,.55)!important;
          box-shadow:inset 0 1px 0 rgba(255,255,255,.08),
                     inset 0 -5px 11px rgba(0,0,0,.40),
                     0 14px 26px rgba(0,0,0,.50),
                     0 0 20px rgba(56,189,248,.20)!important;
        }
</style>
""", unsafe_allow_html=True)

# --- DIRETÓRIO DE ANEXOS ---
if not os.path.exists("anexos"):
    os.makedirs("anexos")

DATABASE = "chamados.csv"
DATABASE_COMPRAS = "solicitacoes_compras.csv"
BACKUP_DIR = "backups"

COLUNAS = [
    "Data", "Protocolo", "Usuário", "Setor", "Categoria", "Impacto", "Prioridade",
    "Descrição", "Status", "Resposta do Técnico", "Técnico Responsável",
    "Atualizado em", "Anexo"
]
COLUNAS_COMPRAS = [
    "Data", "Protocolo", "Usuário", "Setor", "Categoria", "Item", "Quantidade",
    "Justificativa", "Urgência", "Status", "Observação TI", "Responsável",
    "Atualizado em"
]

# ── SQLite — substitui os dois CSVs por um único arquivo transacional ────────
DATABASE_SQLITE = "servicedesk.db"
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


def token_padrao_ativo():
    return obter_token_admin() == "admin123"


def garantir_colunas(df, colunas):
    df_ok = df.copy()
    for coluna in colunas:
        if coluna not in df_ok.columns:
            df_ok[coluna] = ""
    extras = [coluna for coluna in df_ok.columns if coluna not in colunas]
    return df_ok[colunas + extras]


# ── SANITIZAÇÃO DE TEXTO ANTES DE PERSISTIR ─────────────────────────────────
# Defende contra a corrupção de colunas que já ocorreu: descrições com aspas,
# quebras de linha e caracteres de controle "vazavam" para os campos seguintes,
# deslocando Status/Técnico/Data. Roda em toda gravação (choke point único).
_SAN_CTRL_RE  = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_SAN_QUOTE_RE = re.compile(r'"{3,}')

def _sanitizar_celula(v):
    if v is None:
        return ""
    txt = str(v)
    txt = txt.replace("\r\n", "\n").replace("\r", "\n")  # normaliza quebras
    txt = _SAN_CTRL_RE.sub("", txt)                       # remove caracteres de controle
    txt = _SAN_QUOTE_RE.sub("", txt)                      # remove runs de aspas (artefato CSV)
    txt = re.sub(r"\n{3,}", "\n\n", txt)                  # no máximo 2 quebras seguidas
    return txt.strip()


def salvar_dataframe_seguro(df, caminho, colunas):
    """Persiste o DataFrame no SQLite — fonte única de verdade.

    (O snapshot automático em CSV foi removido: o banco SQLite é o único
    armazenamento agora. Isso evita o crescimento descontrolado de backups/.)
    """
    df_ok = garantir_colunas(df, colunas)

    # Sanitiza todas as células de texto para evitar corrupção de colunas
    for _col in df_ok.columns:
        df_ok[_col] = df_ok[_col].map(_sanitizar_celula)

    # Persiste no SQLite (replace atômico da tabela)
    tabela = TABLE_CHAMADOS if caminho == DATABASE else TABLE_COMPRAS
    _conn  = _conectar()
    df_ok.to_sql(tabela, _conn, if_exists="replace", index=False)
    # Recria índices removidos pelo replace
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

# ═══════════════════════════════════════════════════════
# MICRO-LOADING STATE (enterprise UX)
# Mostra skeleton shimmer após ações que atualizam o app.
# ═══════════════════════════════════════════════════════
def _ui_skeleton_html(context_label: str = "Atualizando...") -> str:
    # Renderiza skeleton em HTML para aproveitar o CSS já injetado no app.
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
    # Guardamos o "até quando" em timestamp para desenhar por um intervalo curto.
    st.session_state["ui_skeleton_until"] = time.time() + float(duration_s)
    st.session_state["ui_skeleton_label"] = context_label

def maybe_show_ui_skeleton() -> None:
    # Limpa o estado sem renderizar skeleton (evita loop de refresh)
    st.session_state.pop("ui_skeleton_until", None)
    st.session_state.pop("ui_skeleton_label", None)


# ═══════════════════════════════════════════════════════
# CONFIGURAÇÃO DE NOTIFICAÇÕES — edite aqui
# ═══════════════════════════════════════════════════════
# IP e porta do servidor Streamlit na sua rede local
APP_HOST = "192.168.1.222"
APP_PORT = 8501

# Número do WhatsApp que receberá as notificações (com DDI, sem + ou espaços)
# Exemplo: "5548999998888"  → Brasil (55) + DDD (48) + número
# Pode ser sobrescrito em .streamlit/secrets.toml com a chave WHATSAPP_DESTINO.
WHATSAPP_DESTINO = "554830293440"  # PABX TI (48 3029-3440)


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

    # Monta o corpo: título em negrito + detalhes; URL em linha própria
    # para o WhatsApp converter automaticamente em link clicável.
    linhas = [f"*{titulo}*", ""]
    linhas += corpo.strip().splitlines()
    if url:
        linhas += ["", url]
    mensagem = "\n".join(linhas)

    def _enviar():
        # --- 1) CallMeBot via HTTP (preferencial) ---
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
                return  # enviado com sucesso
            except Exception:
                pass  # se falhar, tenta o fallback abaixo

        # --- 2) pywhatkit / WhatsApp Web (fallback) ---
        try:
            import pywhatkit as pwk
            try:
                # Envia imediatamente (evita o bug de minuto+1 == 60)
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
            pass  # pywhatkit ausente ou WhatsApp Web indisponível

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
            return  # backup de hoje já feito
        import sqlite3 as _sq3
        _src = _sq3.connect(DATABASE_SQLITE)
        _dst = _sq3.connect(destino)
        with _dst:
            _src.backup(_dst)
        _dst.close()
        _src.close()
        # poda: mantém só as cópias mais recentes
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

    # Journaling padrão (DELETE) — robusto à sincronização da pasta.
    # WAL foi REMOVIDO: deixava arquivos -wal/-shm que a sincronização
    # corrompia, resultando em 'database disk image is malformed'.
    try:
        _conn.execute("PRAGMA journal_mode=DELETE")
    except Exception:
        pass

    # Tabela de chamados
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

    # Tabela de compras
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

    # Índices para queries frequentes de Status e Prioridade
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
        # "Urgente" foi descontinuado: passa a ser tratado como "Crítica"
        if "Prioridade" in _df.columns:
            _df["Prioridade"] = _df["Prioridade"].replace({"Urgente": "Crítica"})
        return _df
    except Exception:
        return pd.DataFrame(columns=colunas)


_inicializar_banco()       # cria as tabelas SQLite + WAL (roda 1x por sessão)

@st.cache_resource(show_spinner=False)
def _backup_diario_db_once(_dia: str):
    """Garante que o backup diário rode no máximo 1x por dia por sessão do
    servidor, evitando checagens de disco a cada rerun (cada clique reexecuta
    o script inteiro no Streamlit). A chave _dia invalida o cache à meia-noite."""
    _backup_diario_db()

_backup_diario_db_once(datetime.now().strftime("%Y%m%d"))  # cópia diária (últimas 14)

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


def ordenar_por_prioridade_e_data(df, coluna_prioridade):
    if df.empty:
        return df
    df_ord = df.copy()
    df_ord["_peso"] = df_ord[coluna_prioridade].apply(peso_prioridade)
    df_ord["_data_ord"] = pd.to_datetime(df_ord["Data"], format="%d/%m/%Y %H:%M", errors="coerce")
    return df_ord.sort_values(["_peso", "_data_ord"], ascending=[False, True]).drop(columns=["_peso", "_data_ord"])


def gerar_pdf(df_geral):
    try:
        from fpdf import FPDF  # import lazy: só carrega ao exportar PDF
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
        from fpdf import FPDF  # import lazy: só carrega ao exportar PDF
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


# ═══════════════════════════════════════════════════════
# SIDEBAR — NAVEGAÇÃO COM QUERY PARAMS (SPA FEEL)
# ═══════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
        <div class="sidebar-header">
            <div class="header-particles" aria-hidden="true"></div>
            <div class="brand-lockup">
                <div class="logo-ring" aria-hidden="true">
                    <div class="logo-orbit"></div>
                    <div class="logo-orbit-2"></div>
                    <div class="logo-mark">
                        <!-- Circuit node — pulso de rede tecnológico -->
                        <svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg" style="overflow:visible;">
                            <defs>
                                <linearGradient id="cG" x1="0" y1="0" x2="1" y2="1">
                                    <stop offset="0%" stop-color="#38bdf8">
                                        <animate attributeName="stop-color" values="#38bdf8;#a78bfa;#38bdf8" dur="4s" repeatCount="indefinite"/>
                                    </stop>
                                    <stop offset="100%" stop-color="#a78bfa">
                                        <animate attributeName="stop-color" values="#a78bfa;#38bdf8;#a78bfa" dur="4s" repeatCount="indefinite"/>
                                    </stop>
                                </linearGradient>
                                <filter id="glow">
                                    <feGaussianBlur stdDeviation="0.8" result="blur"/>
                                    <feMerge><feMergeNode in="blur"/><feMergeNode in="SourceGraphic"/></feMerge>
                                </filter>
                            </defs>
                            <!-- Pulse rings — emanam do nó central -->
                            <circle class="circuit-pulse-ring"   cx="10" cy="10" r="5.5" stroke="url(#cG)" stroke-width="0.9" fill="none"/>
                            <circle class="circuit-pulse-ring-2" cx="10" cy="10" r="5.5" stroke="url(#cG)" stroke-width="0.7" fill="none" opacity=".6"/>
                            <!-- Circuit lines — traçam conexões -->
                            <line class="circuit-line-h" x1="0"  y1="10" x2="4.5"  y2="10" stroke="url(#cG)" stroke-width="1" stroke-linecap="round"/>
                            <line class="circuit-line-h" x1="15.5" y1="10" x2="20" y2="10" stroke="url(#cG)" stroke-width="1" stroke-linecap="round"/>
                            <line class="circuit-line-v" x1="10" y1="0"  x2="10" y2="4.5"  stroke="url(#cG)" stroke-width="1" stroke-linecap="round"/>
                            <line class="circuit-line-v" x1="10" y1="15.5" x2="10" y2="20" stroke="url(#cG)" stroke-width="1" stroke-linecap="round"/>
                            <line class="circuit-line-d" x1="3"  y1="3"  x2="6.8" y2="6.8" stroke="#38bdf8" stroke-width="0.8" stroke-linecap="round" opacity=".7"/>
                            <line class="circuit-line-d" x1="17" y1="3"  x2="13.2" y2="6.8" stroke="#a78bfa" stroke-width="0.8" stroke-linecap="round" opacity=".7"/>
                            <line class="circuit-line-d" x1="3"  y1="17" x2="6.8" y2="13.2" stroke="#a78bfa" stroke-width="0.8" stroke-linecap="round" opacity=".7"/>
                            <line class="circuit-line-d" x1="17" y1="17" x2="13.2" y2="13.2" stroke="#38bdf8" stroke-width="0.8" stroke-linecap="round" opacity=".7"/>
                            <!-- Dots at line ends -->
                            <circle class="circuit-dot-1" cx="0.5"  cy="10"   r="1.2" fill="#38bdf8" filter="url(#glow)"/>
                            <circle class="circuit-dot-2" cx="19.5" cy="10"   r="1.2" fill="#38bdf8" filter="url(#glow)"/>
                            <circle class="circuit-dot-3" cx="10"   cy="0.5"  r="1.2" fill="#a78bfa" filter="url(#glow)"/>
                            <circle class="circuit-dot-4" cx="10"   cy="19.5" r="1.2" fill="#a78bfa" filter="url(#glow)"/>
                            <!-- Nó central — núcleo do sistema -->
                            <circle class="circuit-node" cx="10" cy="10" r="3" fill="url(#cG)" filter="url(#glow)" opacity=".95"/>
                            <circle cx="10" cy="10" r="1.4" fill="white" opacity=".9">
                                <animate attributeName="opacity" values=".5;1;.5" dur="2.8s" repeatCount="indefinite"/>
                                <animate attributeName="r"       values="1.2;1.6;1.2" dur="2.8s" repeatCount="indefinite"/>
                            </circle>
                        </svg>
                    </div>
                </div>
                <div>
                    <div class="sidebar-title">IT Service Desk</div>
                    <div class="sidebar-sub">v3.8</div>
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    components.html(f"""
    <!DOCTYPE html><html><head><meta charset="utf-8">
    <style>
      html,body{{margin:0;padding:0;background:transparent;}}
      .strip{{
        display:flex;align-items:center;gap:8px;
        margin-top:2px;padding:8px 10px;
        border:1px solid rgba(56,189,248,.22);
        border-radius:10px;background:rgba(4,7,15,.45);
        color:#8aa4c2;font-size:11px;
        font-family:'Geist Mono',ui-monospace,SFMono-Regular,monospace;
        text-transform:uppercase;letter-spacing:.08em;
        box-shadow:0 0 40px rgba(56,189,248,.07);overflow:hidden;
      }}
      .strip .dot{{width:7px;height:7px;border-radius:999px;background:#34d399;
        box-shadow:0 0 14px rgba(52,211,153,.8);display:inline-block;
        animation:onp 1.2s ease-in-out infinite;}}
      .strip strong{{color:#d7f9ff;font-weight:700;}}
      .strip .up{{color:#38bdf8;font-weight:700;font-variant-numeric:tabular-nums;margin-left:auto;}}
      @keyframes onp{{0%,100%{{opacity:1}}50%{{opacity:.45}}}}
    </style></head>
    <body>
      <div class="strip">
        <span class="dot"></span><strong>Sistema online</strong>
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
            "label": "🚨 Registrar Chamado",
            "titulo": "Entrada de Chamados",
            "descricao": "Abertura guiada com prioridade, categoria e evidências."
        },
        "solicitar_compra": {
            "label": "🛒 Solicitar Compra TI",
            "titulo": "Compra de TI",
            "descricao": "Solicitações padronizadas para equipamentos e serviços."
        },
        "historico": {
            "label": "📊 Meu Histórico",
            "titulo": "Histórico Operacional",
            "descricao": "Acompanhe chamados e requisições por status."
        },
        "gestor": {
            "label": "🔐 Painel do Gestor (TI)",
            "titulo": "Console Gerencial",
            "descricao": "Fila ativa, prioridades e decisões da equipe de TI."
        },
        "relatorios": {
            "label": "📥 Relatórios & Exportação",
            "titulo": "Central de Dados",
            "descricao": "KPIs, backups e exportações para análise."
        },
    }
    lista_paginas = list(paginas_nav.keys())
    rotas_legadas = {dados["label"]: chave for chave, dados in paginas_nav.items()}
    
    # 1. Inicializa o estado da página lendo a URL apenas na primeira vez
    if "pagina_atual" not in st.session_state:
        param_pagina = st.query_params.get("page", "registrar_chamado")
        st.session_state.pagina_atual = (
            param_pagina if param_pagina in lista_paginas
            else rotas_legadas.get(param_pagina, "registrar_chamado")
        )
    elif st.session_state.pagina_atual not in lista_paginas:
        st.session_state.pagina_atual = rotas_legadas.get(st.session_state.pagina_atual, "registrar_chamado")

    # 2. Callback: Quando o usuário clicar, atualiza a URL imediatamente
    def atualizar_url():
        st.query_params["page"] = st.session_state.pagina_atual

    # 3. Widget de rádio amarrado ao session_state via "key"
    pagina = st.radio(
        "NAVEGAÇÃO PRINCIPAL",
        lista_paginas,
        key="pagina_atual", # Isso conecta o valor do botão diretamente ao session_state
        on_change=atualizar_url, # Aciona a função acima no exato momento do clique
        format_func=lambda chave: paginas_nav[chave]["label"],
        label_visibility="visible"
    )

    nav_ativa = paginas_nav[pagina]
    st.markdown(f"""
        <div class="nav-context">
            <div class="nav-context-signal" aria-hidden="true"><span></span><span></span><span></span></div>
            <div class="nav-context-label">Módulo ativo</div>
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

# MICRO-LOADING: skeleton shimmer por transição curta
maybe_show_ui_skeleton()

# ═══════════════════════════════════════════════════════
# 1. ABA: REGISTRAR CHAMADO
# ═══════════════════════════════════════════════════════
if pagina == "registrar_chamado":

    # Página isolada em fragment: cliques reroda só esta página, não a app inteira.
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
            ("Baixa",   "#38bdf8", "rgba(56,189,248,.12)",  72, "🔵", "Posso aguardar"),
            ("Média",   "#fbbf24", "rgba(251,191,36,.12)",  24, "🟡", "Afeta meu trabalho"),
            ("Alta",    "#f97316", "rgba(249,115,22,.12)",   4, "🟠", "Estou parado"),
            ("Crítica", "#ef4444", "rgba(239,68,68,.12)",    2, "🔴", "Empresa parada"),
        ]

        st.markdown("""
<style>
    label,.stRadio label{outline:none!important;box-shadow:none!important}
    .wz{display:flex;align-items:center;padding:6px 0 20px;user-select:none}
    .wz-node{display:flex;flex-direction:column;align-items:center;gap:5px;flex:0 0 auto}
    .wz-circle{width:30px;height:30px;border-radius:50%;display:flex;align-items:center;
      justify-content:center;font-size:12px;font-weight:800;transition:all .25s}
    .wz-c-done{background:#10b981;color:#fff;box-shadow:0 0 10px rgba(16,185,129,.45)}
    .wz-c-act{background:#0d1e35;border:2px solid #38bdf8;color:#38bdf8;box-shadow:0 0 12px rgba(56,189,248,.3)}
    .wz-c-idle{background:#0a1520;border:2px solid #132030;color:#1e3040}
    .wz-lbl{font-size:10px;font-weight:700;letter-spacing:.02em;white-space:nowrap}
    .wz-lbl-done{color:#34d399}.wz-lbl-act{color:#cbd5e1}.wz-lbl-idle{color:#1e3040}
    .wz-line{flex:1;height:2px;margin:0 8px;margin-bottom:15px;border-radius:2px}
    .wz-line-done{background:#10b981}.wz-line-idle{background:#132030}
    .hz{background:linear-gradient(135deg,#0a1828,#060c18);border:1px solid #132030;
      border-radius:12px;padding:18px 22px;position:relative;overflow:hidden;margin-bottom:14px;
      box-shadow:inset 0 1px 0 rgba(255,255,255,.05),inset 0 -6px 14px rgba(0,0,0,.40),0 10px 22px rgba(0,0,0,.40),0 2px 5px rgba(0,0,0,.30)}
    .hz::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;
      background:linear-gradient(90deg,#38bdf8,#a78bfa);box-shadow:0 0 12px rgba(56,189,248,.4)}
    .hz-eye{font-size:9px;font-weight:800;letter-spacing:.15em;text-transform:uppercase;color:#2a4a60;margin-bottom:5px}
    .hz-title{font-size:1.1rem;font-weight:800;color:#e2e8f0;margin-bottom:3px}
    .hz-desc{font-size:.78rem;color:#4a6a80;line-height:1.5}
    .hz-crumb{font-size:10px;color:#3a5a70;margin-bottom:5px}
    .hz-crumb b{color:#5a7a90}
    .prio-chip{display:inline-flex;align-items:center;gap:7px;padding:6px 14px;
      border-radius:8px;font-size:11px;font-weight:700;margin:6px 0 2px}
    .prio-dot{width:7px;height:7px;border-radius:50%;flex-shrink:0}
    .char-c{font-size:9px;text-align:right;margin-top:3px}
    .ok-card{background:linear-gradient(145deg,#071a10,#040d18);
      border:1.5px solid #1a4a2a;border-radius:14px;padding:28px 24px;text-align:center;
      box-shadow:inset 0 1px 0 rgba(255,255,255,.05),inset 0 -7px 16px rgba(0,0,0,.40),0 12px 26px rgba(0,0,0,.42),0 0 40px rgba(52,211,153,.12);margin-bottom:12px}
    .ok-card::before{content:'';display:block;height:3px;
      background:linear-gradient(90deg,#10b981,#34d399);border-radius:12px 12px 0 0;margin:-28px -24px 24px}
    .ok-proto{font-size:.72rem;font-family:monospace;background:#0a2015;
      border:1px solid #1a4a2a;color:#6ee7b7;padding:4px 14px;border-radius:7px;
      display:inline-block;margin:6px 0 12px;letter-spacing:.08em}
    .ok-steps{display:grid;grid-template-columns:repeat(3,1fr);gap:7px;margin-top:14px}
    .ok-s{background:linear-gradient(160deg,#0a2014,#050f0a);border:1px solid #1a3a20;border-radius:8px;padding:9px 7px;box-shadow:inset 0 1px 0 rgba(255,255,255,.05),inset 0 -3px 7px rgba(0,0,0,.40),0 4px 10px rgba(0,0,0,.35)}
    .ok-sn{font-size:8px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:#34d399}
    .ok-st{font-size:10px;font-weight:700;color:#3a7a50;margin-top:2px}
    /* CAT CARD — tudo no botão, sem markdown separado */
    div[data-testid="stButton"] button.cat-btn {
      background:#0b1524; border:1.5px solid #132030; border-radius:11px;
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
      color:#d7e3f4!important;
    }
    div[data-baseweb="base-input"],div[data-baseweb="input"]{
      background:linear-gradient(160deg,#08131f 0%,#050d18 100%)!important;
      border:1px solid rgba(56,189,248,.22)!important;
      border-radius:10px!important;
      box-shadow:inset 0 1px 0 rgba(255,255,255,.04),inset 0 -1px 0 rgba(0,0,0,.35),0 4px 18px rgba(0,0,0,.40)!important;
      transition:border-color .22s ease,box-shadow .22s ease,transform .18s ease!important;
    }
    div[data-baseweb="base-input"]:focus-within,div[data-baseweb="input"]:focus-within{
      border-color:rgba(56,189,248,.72)!important;
      box-shadow:0 0 0 3px rgba(56,189,248,.13),0 0 28px rgba(56,189,248,.16),inset 0 1px 0 rgba(56,189,248,.10)!important;
      transform:translateY(-2px)!important;
    }
    div[data-baseweb="base-input"]:hover:not(:focus-within),div[data-baseweb="input"]:hover:not(:focus-within){
      border-color:rgba(56,189,248,.36)!important;transform:translateY(-1px)!important;
    }
    div[data-baseweb="base-input"] input,div[data-baseweb="input"] input{
      background:transparent!important;color:#bfdbfe!important;
      font-family:'Geist',sans-serif!important;font-size:.90rem!important;
      font-weight:500!important;caret-color:#38bdf8!important;padding:0 14px!important;
    }
    div[data-baseweb="base-input"] input::placeholder,div[data-baseweb="input"] input::placeholder{
      color:rgba(148,178,208,.50)!important;font-style:normal!important;font-size:.85rem!important;
    }
    div[data-baseweb="select"]>div:first-child{
      background:linear-gradient(160deg,#08131f 0%,#050d18 100%)!important;
      border:1px solid rgba(56,189,248,.22)!important;border-radius:10px!important;
      box-shadow:inset 0 1px 0 rgba(255,255,255,.04),0 4px 18px rgba(0,0,0,.40)!important;
      transition:border-color .22s ease,box-shadow .22s ease,transform .18s ease!important;
    }
    div[data-baseweb="select"]>div:first-child:hover{
      border-color:rgba(56,189,248,.38)!important;transform:translateY(-1px)!important;
    }
    div[data-baseweb="select"]:focus-within>div:first-child{
      border-color:rgba(56,189,248,.72)!important;
      box-shadow:0 0 0 3px rgba(56,189,248,.13),0 0 28px rgba(56,189,248,.16),inset 0 1px 0 rgba(56,189,248,.10)!important;
      transform:translateY(-2px)!important;
    }
    div[data-baseweb="select"] svg{
      fill:rgba(56,189,248,.52)!important;transition:transform .22s ease,fill .2s ease!important;
    }
    div[data-baseweb="select"]:focus-within svg{fill:#38bdf8!important;transform:rotate(180deg)!important;}
    ul[data-baseweb="menu"],div[data-baseweb="menu"]{
      background:linear-gradient(160deg,#07101e,#040a16)!important;
      border:1px solid rgba(56,189,248,.26)!important;border-radius:12px!important;
      box-shadow:0 0 0 1px rgba(56,189,248,.06),0 24px 60px rgba(0,0,0,.80),0 0 40px rgba(56,189,248,.10)!important;
      padding:5px!important;
    }
    li[role="option"]{
      border-radius:8px!important;color:#8fb0cc!important;
      font-family:'Geist',sans-serif!important;font-size:.87rem!important;font-weight:500!important;
      transition:background .14s ease,color .14s ease,padding-left .14s ease!important;padding:9px 14px!important;
    }
    li[role="option"]:hover{
      background:linear-gradient(90deg,rgba(56,189,248,.16),rgba(56,189,248,.06))!important;
      color:#e0f2fe!important;padding-left:18px!important;
    }
    li[role="option"][aria-selected="true"]{
      color:#7dd3fc!important;font-weight:700!important;border-left:2px solid rgba(56,189,248,.75)!important;
      background:linear-gradient(90deg,rgba(56,189,248,.18),rgba(56,189,248,.06))!important;
    }
    </style>
    """, unsafe_allow_html=True)

        # ══════════════════════════════════════════════════════════════════
        # LIVE FEED — Chamados Ativos (sem autenticação, leitura pública)
        # ══════════════════════════════════════════════════════════════════
        st.markdown("""
<style>
    /* ── LIVE FEED PANEL ── */
    .lf-wrap {
        position: relative;
        margin-bottom: 18px;
        border: 1px solid rgba(56,189,248,.14);
        border-radius: 12px;
        background: linear-gradient(145deg, rgba(8,14,26,.96), rgba(4,8,18,.92));
        overflow: hidden;
        box-shadow: 0 0 0 1px rgba(56,189,248,.04), 0 14px 36px rgba(0,0,0,.38);
    }
    /* top scanner line */
    .lf-wrap::before {
        content:"";
        position:absolute; top:0; left:0; right:0; height:1px;
        background: linear-gradient(90deg, transparent, rgba(56,189,248,.55), rgba(167,139,250,.35), transparent);
        animation: lfScanline 3s ease-in-out infinite;
    }
    @keyframes lfScanline {
        0%,100%{opacity:.55} 50%{opacity:1}
    }

    .lf-head {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 10px 16px 9px;
        border-bottom: 1px solid rgba(56,189,248,.08);
        background: rgba(4,8,18,.55);
        backdrop-filter: blur(4px);
    }
    .lf-head-left  { display:flex; align-items:center; gap:10px; }
    .lf-head-right { display:flex; align-items:center; gap:8px; }

    .lf-live-dot {
        position: relative;
        width: 7px; height: 7px; border-radius: 50%;
        background: #34d399;
        box-shadow: 0 0 8px rgba(52,211,153,.9);
        flex-shrink: 0;
    }
    .lf-live-dot::after {
        content:"";
        position:absolute; inset:-4px; border-radius:50%;
        background: rgba(52,211,153,.25);
        animation: lfPulse 1.8s ease-out infinite;
    }
    @keyframes lfPulse {
        0%{transform:scale(.6);opacity:1} 100%{transform:scale(2.2);opacity:0}
    }

    .lf-label {
        font-family:'Geist Mono',monospace;
        font-size:.60rem; font-weight:800; letter-spacing:.18em;
        text-transform:uppercase; color:rgba(148,163,184,.60);
    }
    .lf-count-badge {
        display:inline-flex; align-items:center;
        padding:2px 9px; border-radius:999px;
        background:rgba(56,189,248,.08);
        border:1px solid rgba(56,189,248,.20);
        font-family:'Geist Mono',monospace;
        font-size:.62rem; font-weight:800; color:#38bdf8;
        letter-spacing:.06em;
    }
    .lf-live-tag {
        font-family:'Geist Mono',monospace;
        font-size:.55rem; font-weight:800; letter-spacing:.16em;
        color:rgba(52,211,153,.65); text-transform:uppercase;
    }

    /* ── ROWS ── */
    .lf-body {
        max-height: 236px;
        overflow-y: auto;
        overflow-x: hidden;
        scrollbar-width: thin;
        scrollbar-color: rgba(56,189,248,.25) transparent;
    }
    .lf-body::-webkit-scrollbar { width:3px; }
    .lf-body::-webkit-scrollbar-thumb {
        background:rgba(56,189,248,.28);
        border-radius:99px;
    }

    .lf-row {
        display: grid;
        grid-template-columns: 4px 112px 1fr 80px 74px 76px;
        align-items: center;
        gap: 0 14px;
        padding: 8px 16px 8px 12px;
        border-bottom: 1px solid rgba(148,163,184,.06);
        transition: background .15s ease;
        min-height: 44px;
    }
    .lf-row:last-child { border-bottom: none; }
    .lf-row:hover { background: rgba(56,189,248,.04); }

    /* priority stripe */
    .lf-stripe { border-radius: 2px; height: 100%; min-height: 28px; width: 4px; }
    .lf-stripe-critica { background:#f87171; box-shadow:0 0 8px rgba(248,113,113,.65); }
    .lf-stripe-alta    { background:#fb923c; box-shadow:0 0 8px rgba(251,146,60,.55); }
    .lf-stripe-media   { background:#fbbf24; box-shadow:0 0 6px rgba(251,191,36,.40); }
    .lf-stripe-baixa   { background:#38bdf8; box-shadow:0 0 6px rgba(56,189,248,.35); }

    .lf-user {
        display:flex; flex-direction:column; gap:1px;
        overflow:hidden; min-width:0;
    }
    .lf-user-name {
        font-size:.74rem; font-weight:700; color:#e2e8f0;
        white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
    }
    .lf-protocol {
        font-family:'Geist Mono',monospace;
        font-size:.58rem; font-weight:700; letter-spacing:.03em;
        color:rgba(56,189,248,.65); white-space:nowrap; overflow:hidden;
        text-overflow:ellipsis;
    }
    .lf-desc {
        font-size:.76rem; font-weight:600; color:#cbd5e1;
        white-space:normal; overflow-wrap:anywhere; line-height:1.45;
    }
    .lf-meta {
        font-size:.68rem; color:rgba(125,143,168,.65);
        white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
    }
    .lf-time {
        font-family:'Geist Mono',monospace;
        font-size:.64rem; color:rgba(148,163,184,.55);
        text-align:right; white-space:nowrap;
    }
    .lf-status {
        display:inline-flex; align-items:center; gap:4px;
        padding:3px 8px; border-radius:999px;
        font-size:.60rem; font-weight:800; letter-spacing:.04em;
        white-space:nowrap;
    }
    .lf-status-aberto {
        background:rgba(148,163,184,.10); color:#94a3b8;
        border:1px solid rgba(148,163,184,.20);
    }
    .lf-status-em {
        background:rgba(167,139,250,.10); color:#a78bfa;
        border:1px solid rgba(167,139,250,.25);
    }

    /* ── EMPTY STATE ── */
    .lf-empty {
        display:flex; align-items:center; justify-content:center; gap:10px;
        padding:18px;
        color:rgba(125,143,168,.45);
        font-family:'Geist Mono',monospace;
        font-size:.68rem; letter-spacing:.06em;
    }
    .lf-empty::before {
        content:""; width:6px; height:6px; border-radius:50%;
        background:rgba(125,143,168,.30);
        box-shadow:0 0 6px rgba(125,143,168,.15);
    }

    /* ── FOOTER BAR ── */
    .lf-foot {
        display:flex; align-items:center; justify-content:space-between;
        padding:6px 16px;
        border-top:1px solid rgba(56,189,248,.06);
        background:rgba(4,8,18,.35);
    }
    .lf-foot-txt {
        font-family:'Geist Mono',monospace;
        font-size:.55rem; color:rgba(56,189,248,.30);
        letter-spacing:.08em; text-transform:uppercase;
    }
    .lf-critical-alert {
        display:inline-flex; align-items:center; gap:5px;
        font-family:'Geist Mono',monospace;
        font-size:.56rem; font-weight:800; color:#f87171;
        letter-spacing:.08em; animation: lfAlertBlink 1.2s ease-in-out infinite;
    }
    @keyframes lfAlertBlink {
        0%,100%{opacity:.9} 50%{opacity:.4}
    }
    </style>
    """, unsafe_allow_html=True)

        # ── dados ativos (sem autenticação) ──────────────────────────────
        _df_ativos = carregar_dados_cacheados(DATABASE, COLUNAS)
        _df_ativos = _df_ativos[_df_ativos["Status"].isin(["Aberto", "Em Atendimento"])].copy()

        # ordenar: prioridade desc, data asc
        _PRIO_PESO = {"Crítica": 5, "Urgente": 5, "Alta": 4, "Média": 3, "Baixa": 1}
        if not _df_ativos.empty:
            _df_ativos["_pw"] = _df_ativos["Prioridade"].map(_PRIO_PESO).fillna(0)
            _df_ativos["_dt"] = pd.to_datetime(_df_ativos["Data"], format="%d/%m/%Y %H:%M", errors="coerce")
            _df_ativos = _df_ativos.sort_values(["_pw", "_dt"], ascending=[False, True])

        _total_ativos = len(_df_ativos)
        _criticos     = int((_df_ativos["Prioridade"].isin(["Crítica", "Urgente", "Alta"])).sum()) if not _df_ativos.empty else 0

        # ── build rows HTML ───────────────────────────────────────────────
        def _lf_stripe_cls(prio):
            p = str(prio).lower()
            if "cr" in p or "urg" in p: return "lf-stripe-critica"
            if "alt" in p:               return "lf-stripe-alta"
            if "m" in p:                 return "lf-stripe-media"
            return "lf-stripe-baixa"

        def _lf_status_cls(status):
            return "lf-status-em" if status == "Em Atendimento" else "lf-status-aberto"

        def _lf_tempo(data_txt):
            try:
                dt = datetime.strptime(str(data_txt), "%d/%m/%Y %H:%M")
                delta = datetime.now() - dt
                m = int(delta.total_seconds() / 60)
                if m < 60:  return f"{m}m"
                if m < 1440: return f"{m//60}h{m%60:02d}m"
                return f"{m//1440}d"
            except Exception:
                return "—"

        _rows_html = ""
        for _, _r in _df_ativos.head(12).iterrows():
            _proto  = html.escape(str(_r.get("Protocolo","—"))[-16:])
            _user   = html.escape(str(_r.get("Usuário","—")).strip()[:22] or "—")
            _cat    = html.escape(str(_r.get("Categoria","—")))
            _desc   = html.escape(str(_r.get("Descrição","")).strip() or "—")
            _setor  = html.escape(str(_r.get("Setor","—")))
            _prio   = html.escape(str(_r.get("Prioridade","—")))
            _status = str(_r.get("Status","—"))
            _tempo  = _lf_tempo(_r.get("Data",""))
            _stripe = _lf_stripe_cls(str(_r.get("Prioridade","")))
            _stcls  = _lf_status_cls(_status)
            _st_lbl = "EM ATEND." if _status == "Em Atendimento" else "ABERTO"
            # Row em linha única — sem quebras que viram linhas em branco no parser MD
            _rows_html += (
                f'<div class="lf-row">'
                f'<div class="lf-stripe {_stripe}"></div>'
                f'<div class="lf-user"><div class="lf-user-name">{_user}</div><div class="lf-protocol">{_proto}</div></div>'
                f'<div><div class="lf-desc">{_cat} &mdash; {_desc}</div>'
                f'<div class="lf-meta">{_setor}</div></div>'
                f'<div class="lf-meta" style="text-align:right">{_prio}</div>'
                f'<div class="lf-time">{_tempo}</div>'
                f'<div><span class="lf-status {_stcls}">{_st_lbl}</span></div>'
                f'</div>'
            )

        if not _rows_html:
            _rows_html = '<div class="lf-empty">nenhum chamado ativo no momento</div>'

        # Montar HTML via concatenação — evita linhas em branco de variáveis vazias
        # que quebram o parser Markdown do Streamlit e exibem o HTML como texto puro
        _critico_span = (
            f'<span class="lf-critical-alert">&#9873; {_criticos} CR&Iacute;TICO{"S" if _criticos > 1 else ""}</span>'
            if _criticos > 0 else ""
        )
        _mais_span = (
            f'<span class="lf-foot-txt">+{_total_ativos - 12} n&atilde;o exibidos &middot; Meu Hist&oacute;rico</span>'
            if _total_ativos > 12 else ""
        )
        _lf_html = (
            '<div class="lf-wrap">'
            '<div class="lf-head">'
            '<div class="lf-head-left">'
            '<div class="lf-live-dot"></div>'
            '<span class="lf-label">Chamados Ativos</span>'
            f'<span class="lf-count-badge">{_total_ativos}</span>'
            '</div>'
            f'<div class="lf-head-right">{_critico_span}<span class="lf-live-tag">&#9679; LIVE</span></div>'
            '</div>'
            f'<div class="lf-body">{_rows_html}</div>'
            '<div class="lf-foot">'
            '<span class="lf-foot-txt">solicitante &middot; categoria &middot; setor &middot; prioridade &middot; tempo &middot; status</span>'
            f'{_mais_span}'
            '</div>'
            '</div>'
        )
        st.markdown(_lf_html, unsafe_allow_html=True)

        # ── SUCESSO ──────────────────────────────────────────────────────
        if "ultimo_chamado_registrado" in st.session_state:
            proto_ok = st.session_state.pop("ultimo_chamado_registrado")
            st.markdown(f"""
        <div class="ok-card">
          <div style="font-size:2.4rem;margin-bottom:8px">✅</div>
          <div style="font-size:1.1rem;font-weight:800;color:#34d399;margin-bottom:4px">Chamado aberto!</div>
          <div class="ok-proto">{proto_ok}</div>
          <div style="font-size:.78rem;color:#2a5040;margin-bottom:16px;line-height:1.6">Ticket registrado — a equipe de TI já foi notificada.</div>
          <div class="ok-steps">
            <div class="ok-s"><div class="ok-sn">01</div><div class="ok-st">Na fila ✓</div></div>
            <div class="ok-s"><div class="ok-sn">02</div><div class="ok-st">Triagem</div></div>
            <div class="ok-s"><div class="ok-sn">03</div><div class="ok-st">Técnico</div></div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        # ── WIZARD BAR ──────────────────────────────────────────────────
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

        # ════ PASSO 1 ═══════════════════════════════════════════════════
        if wz == 1:
            st.markdown("""<style>
            .ctx-panel{background:linear-gradient(160deg,rgba(12,24,40,.92),rgba(6,11,21,.85));
              border:1px solid rgba(56,189,248,.14);border-radius:14px;padding:20px 20px 18px;
              box-shadow:inset 0 1px 0 rgba(255,255,255,.05),inset 0 -7px 16px rgba(0,0,0,.34),0 12px 28px rgba(0,0,0,.40),0 2px 6px rgba(0,0,0,.30);position:relative;overflow:hidden}
            .ctx-panel::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;
              background:linear-gradient(90deg,#38bdf8,#0284c7);box-shadow:0 0 12px rgba(56,189,248,.4)}
            .ctx-h{font-family:'Geist Mono',monospace;font-size:.62rem;font-weight:800;letter-spacing:.16em;
              text-transform:uppercase;color:#38bdf8;margin-bottom:16px}
            .ctx-step{display:flex;gap:12px;align-items:flex-start;margin-bottom:14px}
            .ctx-n{flex:0 0 24px;width:24px;height:24px;border-radius:50%;display:flex;align-items:center;
              justify-content:center;font-size:11px;font-weight:800;color:#38bdf8;
              background:rgba(56,189,248,.10);border:1px solid rgba(56,189,248,.3)}
            .ctx-step strong{display:block;font-size:.84rem;color:#e2e8f0;font-weight:700;margin-bottom:2px}
            .ctx-step span{font-size:.74rem;color:#7d8fa8;line-height:1.45}
            .ctx-divider{height:1px;background:linear-gradient(90deg,transparent,rgba(56,189,248,.25),transparent);margin:4px 0 14px}
            .ctx-tip{font-size:.76rem;color:#bae6fd;background:rgba(56,189,248,.06);
              border:1px solid rgba(56,189,248,.13);border-radius:9px;padding:10px 12px;line-height:1.5}
            </style>""", unsafe_allow_html=True)
            col_form, col_ctx = st.columns([1.5, 1], gap="large")
            with col_form:
                st.markdown("""<div class="hz">
          <div class="hz-eye">Service Desk · Ticket Intake</div>
          <div class="hz-title">Abertura de Chamado</div>
          <div class="hz-desc">Comece pela sua identificação — leva menos de 1 minuto.</div>
        </div>""", unsafe_allow_html=True)
                with st.container(border=True):
                    nome_tmp = st.text_input("👤 Nome Completo",
                        value=st.session_state.get("ch_nome",""), placeholder="Ex: Hiago Cavalheiro")
                    SETORES = ["GERENCIA","SDR","TI","FINANCEIRO","OPERACIONAL","VENDAS","PÓS-VENDAS","RH","SUPERVISÃO","MARKETING","OUTRO"]
                    idx_s = SETORES.index(st.session_state.get("ch_setor","OPERACIONAL")) if st.session_state.get("ch_setor","OPERACIONAL") in SETORES else 4
                    setor_tmp = st.selectbox("🏢 Setor", SETORES, index=idx_s)
                if st.button("Continuar →", use_container_width=True, type="primary", key="ch_n1"):
                    if nome_tmp.strip():
                        st.session_state.ch_nome = nome_tmp.strip()
                        st.session_state.ch_setor = setor_tmp
                        st.session_state.ch_step = 2
                        st.rerun(scope="fragment")
                    else:
                        st.warning("⚠️ Informe seu nome para continuar.")
            with col_ctx:
                st.markdown("""<div class="ctx-panel">
          <div class="ctx-h">Como funciona</div>
          <div class="ctx-step"><span class="ctx-n">1</span><div><strong>Identifique-se</strong><span>Seu nome e setor para localizarmos você.</span></div></div>
          <div class="ctx-step"><span class="ctx-n">2</span><div><strong>Categoria</strong><span>Escolha o tipo de problema em cards.</span></div></div>
          <div class="ctx-step"><span class="ctx-n">3</span><div><strong>Detalhes &amp; urgência</strong><span>Descreva o problema e defina a prioridade.</span></div></div>
          <div class="ctx-divider"></div>
          <div class="ctx-tip">⚡ Quanto mais detalhes na descrição, mais rápido o time resolve.</div>
        </div>""", unsafe_allow_html=True)

        # ════ PASSO 2 ═══════════════════════════════════════════════════
        elif wz == 2:
            nome_s2 = st.session_state.get("ch_nome","")
            st.markdown(f"""<div class="hz">
          <div class="hz-eye">Service Desk · Categoria</div>
          <div class="hz-title">Olá, {nome_s2}! Qual o tipo de problema?</div>
          <div class="hz-desc">Clique no card que melhor descreve o incidente.</div>
        </div>""", unsafe_allow_html=True)

            # Elevação 3D ao passar o mouse nos cards de categoria
            st.markdown("""<style>
            .cat3d:hover{
              transform:translateY(-3px)!important;
              border-color:rgba(56,189,248,.55)!important;
              box-shadow:inset 0 1px 0 rgba(255,255,255,.08),
                         inset 0 -5px 11px rgba(0,0,0,.40),
                         0 14px 26px rgba(0,0,0,.50),
                         0 0 20px rgba(56,189,248,.20)!important;
            }
            </style>""", unsafe_allow_html=True)

            cat_sel = st.session_state.get("ch_cat_idx", None)

            # Callback de seleção: roda ANTES do rerun do fragment, então o
            # card já aparece selecionado em um único ciclo (sem st.rerun extra
            # nem refresh do app inteiro = sem "piscada" na página).
            def _sel_cat_ch(_idx, _ico, _label):
                st.session_state.ch_cat_idx = _idx
                st.session_state.ch_cat = f"{_ico} {_label}"

            for row in [CATS[i:i+3] for i in range(0, len(CATS), 3)]:
                cols = st.columns(3, gap="small")
                for col, (ico, label, hint) in zip(cols, row):
                    i_abs = CATS.index((ico, label, hint))
                    with col:
                        is_sel = cat_sel == i_abs
                        bc  = "#38bdf8" if is_sel else "#243648"
                        # Fundo em gradiente (claro no topo → escuro embaixo) = volume
                        bg  = ("linear-gradient(160deg,#103056 0%,#08182f 100%)" if is_sel
                               else "linear-gradient(160deg,#101f33 0%,#0a1320 100%)")
                        # Pilha de sombras 3D: brilho interno no topo + sombra interna
                        # embaixo (dá espessura) + sombra externa projetada (eleva a caixa)
                        _sh3d = ("inset 0 1px 0 rgba(255,255,255,.07),"
                                 "inset 0 -5px 11px rgba(0,0,0,.45),"
                                 "0 8px 18px rgba(0,0,0,.45),"
                                 "0 2px 5px rgba(0,0,0,.30)")
                        glw = (f"box-shadow:{_sh3d},0 0 22px rgba(56,189,248,.30);transform:translateY(-2px);"
                               if is_sel else f"box-shadow:{_sh3d};")
                        ico_col   = "#7dd3fc" if is_sel else "#94a3b8"
                        lbl_col   = "#7dd3fc" if is_sel else "#94a3b8"
                        hint_col  = "#7dd3fc99" if is_sel else "#64748b"   # ← visível agora
                        chk = f'<div style="position:absolute;top:8px;right:9px;width:16px;height:16px;border-radius:50%;background:#38bdf8;color:#fff;font-size:9px;font-weight:900;display:flex;align-items:center;justify-content:center;line-height:1">✓</div>' if is_sel else ""
                        st.markdown(f"""
<div class="cat3d" style="background:{bg};border:1.5px solid {bc};border-radius:12px;
padding:20px 10px 8px;text-align:center;position:relative;{glw}
transition:transform .18s ease,box-shadow .18s ease;margin-bottom:4px">
{chk}
<div style="font-size:2rem;margin-bottom:7px;line-height:1">{ico}</div>
<div style="font-size:12px;font-weight:800;color:{lbl_col};margin-bottom:4px;letter-spacing:.02em">{label}</div>
<div style="font-size:10px;color:{hint_col};line-height:1.4">{hint}</div>
</div>""", unsafe_allow_html=True)
                        st.button(
                            "✓ Selecionado" if is_sel else "Selecionar",
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
                st.caption("👆 Selecione uma categoria para continuar.")

        # ════ PASSO 3 ═══════════════════════════════════════════════════
        elif wz == 3:
            cat_ch   = st.session_state.get("ch_cat","—")
            nome_ch  = st.session_state.get("ch_nome","—")
            setor_ch = st.session_state.get("ch_setor","—")
            st.markdown(f"""<div class="hz">
          <div class="hz-crumb">👤 <b>{nome_ch}</b> &nbsp;·&nbsp; 🏢 {setor_ch} &nbsp;·&nbsp; 🗂 {cat_ch}</div>
          <div class="hz-title">Quase lá! Defina urgência e descreva o problema.</div>
        </div>""", unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown("**⚡ Urgência**")
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
                    f'<div class="prio-dot" style="background:{pd_[1]};box-shadow:0 0 7px {pd_[1]}"></div>'
                    f'<span style="color:{pd_[1]}">Prioridade: <strong>{prio_name}</strong></span></div>',
                    unsafe_allow_html=True)
                st.markdown("---")
                desc = st.text_area("📝 Descrição do problema",
                    value=st.session_state.get("ch_desc",""),
                    placeholder="Descreva o que acontece, quando começou e qualquer mensagem de erro.\nEx: 'O Outlook não abre desde as 8h — erro 0x8004010F'",
                    height=120, key="ch_da")
                nc = len(desc)
                st.markdown(
                    f'<div class="char-c" style="color:{"#34d399" if nc>=30 else "#fbbf24" if nc>0 else "#334155"}">'
                    f'{nc} caracteres {"✓" if nc>=30 else "— recomendado 30+"}</div>',
                    unsafe_allow_html=True)


                st.markdown("---")
                st.markdown("📎 **Evidências** *(opcional)* — prints agilizam o diagnóstico")
                arquivo_anexado = st.file_uploader("", type=["png","jpg","jpeg","pdf"],
                    label_visibility="collapsed", key="ch_f")
            st.markdown(
                f'<div style="background:#07101a;border:1px solid #132030;border-radius:9px;'
                f'padding:10px 15px;margin:8px 0 14px;font-size:11px;color:#4a6a80">'
                f'👤 <strong style="color:#cbd5e1">{nome_ch}</strong> &nbsp;·&nbsp; 🏢 {setor_ch}'
                f' &nbsp;·&nbsp; 🗂 {cat_ch} &nbsp;·&nbsp; '
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
                    st.warning("⚠️ Descreva o problema antes de enviar.")
                else:
                    with st.status("Abrindo chamado...", expanded=True) as status:
                        nome_arquivo_salvo = "Nenhum"
                        if arquivo_anexado:
                            status.write("⬆️ Enviando anexo...")
                            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                            nome_arquivo_salvo = f"{ts}_{arquivo_anexado.name}"
                            with open(os.path.join("anexos", nome_arquivo_salvo), "wb") as f:
                                f.write(arquivo_anexado.getbuffer())
                        status.write("🗄️ Registrando...")
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
                        status.write("🔔 Notificando TI...")
                        _corpo = f"Protocolo: {protocolo}\nUsuário: {nome_ch}\nSetor: {setor_ch}\nPrioridade: {prio_name}"
                        notificar_windows("🚨 NOVO CHAMADO TI!", _corpo)
                        notificar_whatsapp("🚨 NOVO CHAMADO TI!", _corpo, url=_url_pagina("gestor"))
                        status.update(label=f"✅ {protocolo} aberto!", state="complete", expanded=False)
                    for k in ["ch_step","ch_nome","ch_setor","ch_cat","ch_cat_idx","ch_prio","ch_desc"]:
                        st.session_state.pop(k, None)
                    st.session_state.ultimo_chamado_registrado = protocolo
                    st.rerun(scope="fragment")


    # ═══════════════════════════════════════════════════════
    # 2. ABA: SOLICITAR COMPRA TI
    # ═══════════════════════════════════════════════════════
    _render_registrar()
elif pagina == "solicitar_compra":

    # Página isolada em fragment: cliques reroda só esta página, não a app inteira.
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
            "Sem pressa": ("#38bdf8","rgba(56,189,248,.12)",120),
            "Moderada":   ("#38bdf8","rgba(56,189,248,.12)", 48),
            "Alta":       ("#f97316","rgba(249,115,22,.12)",  12),
            "Crítica":    ("#ef4444","rgba(239,68,68,.12)",    4),
        }

        st.markdown("""
<style>
    label,.stRadio label{outline:none!important;box-shadow:none!important}
    .wz{display:flex;align-items:center;padding:6px 0 20px;user-select:none}
    .wz-node{display:flex;flex-direction:column;align-items:center;gap:5px;flex:0 0 auto}
    .wz-circle{width:30px;height:30px;border-radius:50%;display:flex;align-items:center;
      justify-content:center;font-size:12px;font-weight:800;transition:all .25s}
    .wz-c-done{background:#10b981;color:#fff;box-shadow:0 0 10px rgba(16,185,129,.45)}
    .wz-c-act{background:#0a1020;border:2px solid #38bdf8;color:#38bdf8;box-shadow:0 0 12px rgba(56,189,248,.3)}
    .wz-c-idle{background:#0a1520;border:2px solid #132030;color:#1e3040}
    .wz-lbl{font-size:10px;font-weight:700;letter-spacing:.02em;white-space:nowrap}
    .wz-lbl-done{color:#34d399}.wz-lbl-act{color:#cbd5e1}.wz-lbl-idle{color:#1e3040}
    .wz-line{flex:1;height:2px;margin:0 8px;margin-bottom:15px;border-radius:2px}
    .wz-line-done{background:#10b981}.wz-line-idle{background:#132030}
    .hzp{background:linear-gradient(135deg,#0a1020,#040810);border:1px solid #0a1020;
      border-radius:12px;padding:18px 22px;position:relative;overflow:hidden;margin-bottom:14px;
      box-shadow:inset 0 1px 0 rgba(255,255,255,.05),inset 0 -6px 14px rgba(0,0,0,.40),0 10px 22px rgba(0,0,0,.40),0 2px 5px rgba(0,0,0,.30)}
    .hzp::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;
      background:linear-gradient(90deg,#38bdf8,#0369a1);box-shadow:0 0 12px rgba(56,189,248,.4)}
    .hzp-eye{font-size:9px;font-weight:800;letter-spacing:.15em;text-transform:uppercase;color:#1a2a3a;margin-bottom:5px}
    .hzp-title{font-size:1.1rem;font-weight:800;color:#e2e8f0;margin-bottom:3px}
    .hzp-desc{font-size:.78rem;color:#3d5068;line-height:1.5}
    .hzp-crumb{font-size:10px;color:#27384d;margin-bottom:5px}
    .hzp-crumb b{color:#4a7a9a}
    .prio-chip2{display:inline-flex;align-items:center;gap:7px;padding:6px 14px;
      border-radius:8px;font-size:11px;font-weight:700;margin:6px 0 2px}
    .prio-dot2{width:7px;height:7px;border-radius:50%;flex-shrink:0}
    .ok-cp{background:linear-gradient(145deg,#0a1020,#040810);border:1.5px solid #102844;
      border-radius:14px;padding:28px 24px;text-align:center;
      box-shadow:0 0 40px rgba(56,189,248,.1);margin-bottom:12px}
    .ok-cp::before{content:'';display:block;height:3px;
      background:linear-gradient(90deg,#0369a1,#38bdf8);border-radius:12px 12px 0 0;margin:-28px -24px 24px}
    .ok-cp-proto{font-size:.72rem;font-family:monospace;background:#0a1020;
      border:1px solid #102844;color:#7dd3fc;padding:4px 14px;border-radius:7px;
      display:inline-block;margin:6px 0 12px;letter-spacing:.08em}

    /* ══ INPUT FIELD OVERHAUL — violet theme ══ */
    .stTextInput label p,
    .stSelectbox label p,
    .stNumberInput label p,
    .stTextArea label p {
      font-family:var(--font)!important;
      font-size:.82rem!important;font-weight:700!important;
      letter-spacing:.01em!important;text-transform:none!important;
      color:#d7e3f4!important;
    }
    div[data-baseweb="base-input"],div[data-baseweb="input"]{
      background:linear-gradient(160deg,#080d18 0%,#04070f 100%)!important;
      border:1px solid rgba(56,189,248,.22)!important;
      border-radius:10px!important;
      box-shadow:inset 0 1px 0 rgba(255,255,255,.04),inset 0 -1px 0 rgba(0,0,0,.35),0 4px 18px rgba(0,0,0,.40)!important;
      transition:border-color .22s ease,box-shadow .22s ease,transform .18s ease!important;
    }
    div[data-baseweb="base-input"]:focus-within,div[data-baseweb="input"]:focus-within{
      border-color:rgba(56,189,248,.72)!important;
      box-shadow:0 0 0 3px rgba(56,189,248,.13),0 0 28px rgba(56,189,248,.16),inset 0 1px 0 rgba(56,189,248,.10)!important;
      transform:translateY(-2px)!important;
    }
    div[data-baseweb="base-input"]:hover:not(:focus-within),div[data-baseweb="input"]:hover:not(:focus-within){
      border-color:rgba(56,189,248,.36)!important;transform:translateY(-1px)!important;
    }
    div[data-baseweb="base-input"] input,div[data-baseweb="input"] input{
      background:transparent!important;color:#e0f2fe!important;
      font-family:'Geist',sans-serif!important;font-size:.90rem!important;
      font-weight:500!important;caret-color:#38bdf8!important;padding:0 14px!important;
    }
    div[data-baseweb="base-input"] input::placeholder,div[data-baseweb="input"] input::placeholder{
      color:rgba(148,178,208,.50)!important;font-style:normal!important;font-size:.85rem!important;
    }
    div[data-baseweb="select"]>div:first-child{
      background:linear-gradient(160deg,#080d18 0%,#04070f 100%)!important;
      border:1px solid rgba(56,189,248,.22)!important;border-radius:10px!important;
      box-shadow:inset 0 1px 0 rgba(255,255,255,.04),0 4px 18px rgba(0,0,0,.40)!important;
      transition:border-color .22s ease,box-shadow .22s ease,transform .18s ease!important;
    }
    div[data-baseweb="select"]>div:first-child:hover{
      border-color:rgba(56,189,248,.38)!important;transform:translateY(-1px)!important;
    }
    div[data-baseweb="select"]:focus-within>div:first-child{
      border-color:rgba(56,189,248,.72)!important;
      box-shadow:0 0 0 3px rgba(56,189,248,.13),0 0 28px rgba(56,189,248,.16),inset 0 1px 0 rgba(56,189,248,.10)!important;
      transform:translateY(-2px)!important;
    }
    div[data-baseweb="select"] svg{
      fill:rgba(56,189,248,.52)!important;transition:transform .22s ease,fill .2s ease!important;
    }
    div[data-baseweb="select"]:focus-within svg{fill:#38bdf8!important;transform:rotate(180deg)!important;}
    ul[data-baseweb="menu"],div[data-baseweb="menu"]{
      background:linear-gradient(160deg,#080d18,#04070f)!important;
      border:1px solid rgba(56,189,248,.26)!important;border-radius:12px!important;
      box-shadow:0 0 0 1px rgba(56,189,248,.06),0 24px 60px rgba(0,0,0,.80),0 0 40px rgba(56,189,248,.10)!important;
      padding:5px!important;
    }
    li[role="option"]{
      border-radius:8px!important;color:#7dd3fc!important;
      font-family:'Geist',sans-serif!important;font-size:.87rem!important;font-weight:500!important;
      transition:background .14s ease,color .14s ease,padding-left .14s ease!important;padding:9px 14px!important;
    }
    li[role="option"]:hover{
      background:linear-gradient(90deg,rgba(56,189,248,.16),rgba(56,189,248,.06))!important;
      color:#e0f2fe!important;padding-left:18px!important;
    }
    li[role="option"][aria-selected="true"]{
      color:#7dd3fc!important;font-weight:700!important;border-left:2px solid rgba(56,189,248,.75)!important;
      background:linear-gradient(90deg,rgba(56,189,248,.18),rgba(56,189,248,.06))!important;
    }
        /* CTA primário em violeta — coerente com o tema desta página */
    .stButton > button[kind="primary"]{
      background:linear-gradient(135deg,#38bdf8 0%,#0ea5e9 55%,#0284c7 100%)!important;
      border:1px solid rgba(56,189,248,.5)!important;color:#04121f!important;
      box-shadow:0 8px 22px rgba(56,189,248,.30),inset 0 1px 0 rgba(255,255,255,.30)!important;
    }
    .stButton > button[kind="primary"]:hover{
      box-shadow:0 14px 32px rgba(56,189,248,.45),inset 0 1px 0 rgba(255,255,255,.42)!important;
    }
    .stButton > button[kind="primary"] *,.stButton button[kind="primary"] p{color:#04121f!important}
</style>
    """, unsafe_allow_html=True)

        # ── SUCESSO ──────────────────────────────────────────────────────
        if "ultima_compra_registrada" in st.session_state:
            proto_c = st.session_state.pop("ultima_compra_registrada")
            st.markdown(f"""<div class="ok-cp">
          <div style="font-size:2.4rem;margin-bottom:8px">🛒</div>
          <div style="font-size:1.1rem;font-weight:800;color:#38bdf8;margin-bottom:4px">Solicitação registrada!</div>
          <div class="ok-cp-proto">{proto_c}</div>
          <div style="font-size:.78rem;color:#27384d;margin-top:4px;line-height:1.6">
            Sua requisição foi enviada para análise da TI.</div>
        </div>""", unsafe_allow_html=True)

        # ── WIZARD BAR ──────────────────────────────────────────────────
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

        # ════ PASSO 1 ═══════════════════════════════════════════════════
        if wz_c == 1:
            st.markdown("""<style>
            .ctx-panel{background:linear-gradient(160deg,rgba(12,24,40,.92),rgba(6,11,21,.85));
              border:1px solid rgba(56,189,248,.14);border-radius:14px;padding:20px 20px 18px;
              box-shadow:inset 0 1px 0 rgba(255,255,255,.05),inset 0 -7px 16px rgba(0,0,0,.34),0 12px 28px rgba(0,0,0,.40),0 2px 6px rgba(0,0,0,.30);position:relative;overflow:hidden}
            .ctx-panel::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;
              background:linear-gradient(90deg,#38bdf8,#0284c7);box-shadow:0 0 12px rgba(56,189,248,.4)}
            .ctx-h{font-family:'Geist Mono',monospace;font-size:.62rem;font-weight:800;letter-spacing:.16em;
              text-transform:uppercase;color:#38bdf8;margin-bottom:16px}
            .ctx-step{display:flex;gap:12px;align-items:flex-start;margin-bottom:14px}
            .ctx-n{flex:0 0 24px;width:24px;height:24px;border-radius:50%;display:flex;align-items:center;
              justify-content:center;font-size:11px;font-weight:800;color:#38bdf8;
              background:rgba(56,189,248,.10);border:1px solid rgba(56,189,248,.3)}
            .ctx-step strong{display:block;font-size:.84rem;color:#e2e8f0;font-weight:700;margin-bottom:2px}
            .ctx-step span{font-size:.74rem;color:#7d8fa8;line-height:1.45}
            .ctx-divider{height:1px;background:linear-gradient(90deg,transparent,rgba(56,189,248,.25),transparent);margin:4px 0 14px}
            .ctx-tip{font-size:.76rem;color:#bae6fd;background:rgba(56,189,248,.06);
              border:1px solid rgba(56,189,248,.13);border-radius:9px;padding:10px 12px;line-height:1.5}
            </style>""", unsafe_allow_html=True)
            col_form, col_ctx = st.columns([1.5, 1], gap="large")
            with col_form:
                st.markdown("""<div class="hzp">
          <div class="hzp-eye">IT Procurement · Requisição</div>
          <div class="hzp-title">Solicitação de Compras</div>
          <div class="hzp-desc">Identifique-se e indique a urgência do item.</div>
        </div>""", unsafe_allow_html=True)
                with st.container(border=True):
                    nome_ct = st.text_input("👤 Nome Completo",
                        value=st.session_state.get("cp_nome",""), placeholder="Quem receberá o item?")
                    cF1, cF2 = st.columns(2)
                    with cF1:
                        SETORES_C = ["GERENCIA","SDR","TI","FINANCEIRO","OPERACIONAL","VENDAS","RH","DIRETORIA","SUPERVISÃO","MARKETING","OUTRO"]
                        idx_sc = SETORES_C.index(st.session_state.get("cp_setor","OPERACIONAL")) if st.session_state.get("cp_setor","OPERACIONAL") in SETORES_C else 4
                        setor_ct = st.selectbox("🏢 Setor", SETORES_C, index=idx_sc)
                    with cF2:
                        urg_t = st.selectbox("⚡ Urgência", list(URG_DATA.keys()),
                            index=list(URG_DATA.keys()).index(st.session_state.get("cp_urg","Moderada")) if st.session_state.get("cp_urg","Moderada") in URG_DATA else 1)
                    uc,ubg,uh = URG_DATA[urg_t]
                    st.markdown(
                        f'<div class="prio-chip2" style="background:{ubg};border:1px solid {uc}44">'
                        f'<div class="prio-dot2" style="background:{uc};box-shadow:0 0 7px {uc}"></div>'
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
                        st.warning("⚠️ Informe seu nome para continuar.")
            with col_ctx:
                st.markdown("""<div class="ctx-panel">
          <div class="ctx-h">Como funciona</div>
          <div class="ctx-step"><span class="ctx-n">1</span><div><strong>Solicitante</strong><span>Quem vai receber o item e a urgência.</span></div></div>
          <div class="ctx-step"><span class="ctx-n">2</span><div><strong>O que precisa</strong><span>Categoria e item desejado.</span></div></div>
          <div class="ctx-step"><span class="ctx-n">3</span><div><strong>Justificativa</strong><span>Motivo da compra e link do fornecedor.</span></div></div>
          <div class="ctx-divider"></div>
          <div class="ctx-tip">🔗 Inclua o link do produto na justificativa para acelerar a aprovação.</div>
        </div>""", unsafe_allow_html=True)

        # ════ PASSO 2 — CARDS CLICÁVEIS ═════════════════════════════════
        elif wz_c == 2:
            nome_c2 = st.session_state.get("cp_nome","")
            st.markdown(f"""<div class="hzp">
          <div class="hzp-title">O que você precisa, {nome_c2}?</div>
          <div class="hzp-desc">Clique na categoria do item solicitado.</div>
        </div>""", unsafe_allow_html=True)

            # Elevação 3D ao passar o mouse nos cards de categoria
            st.markdown("""<style>
            .cat3d:hover{
              transform:translateY(-3px)!important;
              border-color:rgba(56,189,248,.55)!important;
              box-shadow:inset 0 1px 0 rgba(255,255,255,.08),
                         inset 0 -5px 11px rgba(0,0,0,.40),
                         0 14px 26px rgba(0,0,0,.50),
                         0 0 20px rgba(56,189,248,.20)!important;
            }
            </style>""", unsafe_allow_html=True)

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
                        bc = "#38bdf8" if is_s else "#243648"
                        bg = ("linear-gradient(160deg,#103056 0%,#08182f 100%)" if is_s
                              else "linear-gradient(160deg,#101f33 0%,#0a1320 100%)")
                        _sh3d = ("inset 0 1px 0 rgba(255,255,255,.07),"
                                 "inset 0 -5px 11px rgba(0,0,0,.45),"
                                 "0 8px 18px rgba(0,0,0,.45),"
                                 "0 2px 5px rgba(0,0,0,.30)")
                        glw = (f"box-shadow:{_sh3d},0 0 22px rgba(56,189,248,.30);transform:translateY(-2px);"
                               if is_s else f"box-shadow:{_sh3d};")
                        lbl_col  = "#7dd3fc" if is_s else "#94a3b8"
                        hint_col = "#7dd3fcaa" if is_s else "#64748b"  # ← visível
                        chk = f'<div style="position:absolute;top:8px;right:9px;width:16px;height:16px;border-radius:50%;background:#38bdf8;color:#fff;font-size:9px;font-weight:900;display:flex;align-items:center;justify-content:center">✓</div>' if is_s else ""
                        st.markdown(f"""
<div class="cat3d" style="background:{bg};border:1.5px solid {bc};border-radius:12px;
padding:20px 10px 8px;text-align:center;position:relative;{glw}
transition:transform .18s ease,box-shadow .18s ease;margin-bottom:4px">
{chk}
<div style="font-size:2rem;margin-bottom:7px;line-height:1">{ico}</div>
<div style="font-size:12px;font-weight:800;color:{lbl_col};margin-bottom:4px">{label}</div>
<div style="font-size:10px;color:{hint_col};line-height:1.4">{hint}</div>
</div>""", unsafe_allow_html=True)
                        st.button("✓" if is_s else "Selecionar",
                                  key=f"ccat_{i_c}", use_container_width=True,
                                  type="primary" if is_s else "secondary",
                                  on_click=_sel_cat_cp, args=(i_c, label))

            st.write("")
            if cat_sel_c is not None:
                with st.container(border=True):
                    col_item, col_qtd = st.columns([3, 1])
                    with col_item:
                        label_sel = CATS_C[cat_sel_c][1]
                        item_s = st.selectbox("🔍 Item específico", ITEMS_C[label_sel], key="cp_is")
                    with col_qtd:
                        qtd_s = st.number_input("Qtd", min_value=1,
                            value=st.session_state.get("cp_qtd",1), key="cp_qs")
                    spec_s = st.text_input("📝 Modelo / Especificação *(opcional)*",
                        value=st.session_state.get("cp_spec",""),
                        placeholder="Ex: Dell Latitude 3420, Logitech G203...", key="cp_ss")
            else:
                st.caption("👆 Selecione uma categoria acima para ver os itens.")
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

        # ════ PASSO 3 ═══════════════════════════════════════════════════
        elif wz_c == 3:
            nome_c3  = st.session_state.get("cp_nome","—")
            setor_c3 = st.session_state.get("cp_setor","—")
            urg_c3   = st.session_state.get("cp_urg","Moderada")
            cat_c3   = st.session_state.get("cp_cat","—")
            item_c3  = st.session_state.get("cp_item","—")
            qtd_c3   = st.session_state.get("cp_qtd",1)
            uc3,ubg3,uh3 = URG_DATA.get(urg_c3, ("#38bdf8","rgba(56,189,248,.12)",48))
            st.markdown(f"""<div class="hzp">
          <div class="hzp-crumb">👤 <b>{nome_c3}</b> · 🏢 {setor_c3} · 📦 {item_c3} ×{qtd_c3} · <b style="color:{uc3}">{urg_c3}</b></div>
          <div class="hzp-title">Justifique a necessidade da compra.</div>
        </div>""", unsafe_allow_html=True)
            with st.container(border=True):
                forn = st.text_input("🔗 Fornecedor / link *(opcional)*",
                    value=st.session_state.get("cp_forn",""),
                    placeholder="Cole um link ou nome do fornecedor", key="cp_fs")
                st.markdown("---")
                just = st.text_area("📋 Justificativa *",
                    value=st.session_state.get("cp_just",""),
                    placeholder="Por que esse item é necessário? Qual o impacto se não for adquirido?",
                    height=120, key="cp_js")
                nj = len(just)
                st.markdown(
                    f'<div style="font-size:9px;text-align:right;margin-top:3px;'
                    f'color:{"#34d399" if nj>=20 else "#fbbf24" if nj>0 else "#334155"}">'
                    f'{nj} chars {"✓" if nj>=20 else "— mínimo 20"}</div>',
                    unsafe_allow_html=True)
            st.markdown(
                f'<div style="background:#04070f;border:1px solid #0a1020;border-radius:9px;'
                f'padding:10px 15px;margin:8px 0 14px;font-size:11px;color:#3d5068">'
                f'👤 <strong style="color:#cbd5e1">{nome_c3}</strong> · 🏢 {setor_c3}'
                f' · 📦 {item_c3} ×{qtd_c3} · <strong style="color:{uc3}">{urg_c3}</strong></div>',
                unsafe_allow_html=True)
            cols_c3 = st.columns([1, 2, 1])
            with cols_c3[0]:
                if st.button("← Voltar", key="cp_b3"):
                    st.session_state.cp_just = just; st.session_state.cp_forn = forn
                    st.session_state.cp_step = 2; st.rerun(scope="fragment")
            enviar_c = st.button("Enviar Solicitação →", use_container_width=True, type="primary", key="cp_env")
            if enviar_c:
                if not just.strip():
                    st.warning("⚠️ Adicione uma justificativa antes de enviar.")
                else:
                    with st.status("Registrando...", expanded=True) as sc:
                        sc.write("🗄️ Gravando no banco...")
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
                        sc.write("🔔 Notificando TI...")
                        _corpo_c = f"Protocolo: {protocolo_compra}\nSolicitante: {nome_c3}\nItem: {item_c3} (x{qtd_c3})\nUrgência: {urg_c3}"
                        notificar_windows("🛒 NOVA SOLICITAÇÃO DE COMPRA", _corpo_c)
                        notificar_whatsapp("🛒 NOVA SOLICITAÇÃO DE COMPRA", _corpo_c, url=_url_pagina("gestor"))
                        sc.update(label=f"✅ {protocolo_compra} registrada!", state="complete", expanded=False)
                    for k in ["cp_step","cp_nome","cp_setor","cp_urg","cp_cat","cp_cat_idx","cp_item","cp_qtd","cp_spec","cp_forn","cp_just"]:
                        st.session_state.pop(k, None)
                    st.session_state.ultima_compra_registrada = protocolo_compra
                    st.rerun(scope="fragment")


    # ═══════════════════════════════════════════════════════
    # 3. ABA: MEU HISTÓRICO
    # ═══════════════════════════════════════════════════════
    _render_solicitar()
elif pagina == "historico":

    # Página isolada em fragment: a busca reroda só este trecho, não a app inteira.
    @st.fragment
    def _render_historico():

        st.markdown("""
        <div class="page-hero hero-compact animate-fade-in">
            <div class="page-hero-eyebrow">Self-Service • Rastreamento</div>
            <div class="page-hero-title">Meu Histórico</div>
            <p class="page-hero-desc">Acompanhe seus chamados e compras em um só lugar.</p>
        </div>
    """, unsafe_allow_html=True)

        # ── REDESIGN DOS CARDS DO HISTÓRICO (escopo isolado .hist-scope) ──
        st.markdown("""
<style>
        .hist-scope .queue-card{
            position:relative; padding:15px 18px 15px 22px; border-radius:14px;
            background:linear-gradient(145deg,rgba(11,17,29,.94),rgba(6,11,20,.94));
            border:1px solid rgba(56,189,248,.10); overflow:hidden;
            transition:transform .18s ease, border-color .2s ease, box-shadow .2s ease;
        }
        .hist-scope .queue-card::before{
            content:""; position:absolute; left:0; top:0; bottom:0; width:4px;
            background:var(--hs-accent,#38bdf8);
            box-shadow:0 0 14px -2px var(--hs-accent,#38bdf8);
        }
        .hist-scope .queue-card:hover{
            transform:translateY(-2px); border-color:rgba(56,189,248,.30);
            box-shadow:0 14px 32px rgba(0,0,0,.42);
        }
        .hist-scope .hs-open{--hs-accent:#f87171;}
        .hist-scope .hs-prog{--hs-accent:#fbbf24;}
        .hist-scope .hs-done{--hs-accent:#34d399;}
        .hist-scope .work-card-head{
            display:flex; justify-content:space-between; align-items:flex-start;
            gap:14px; flex-wrap:wrap;
        }
        .hist-scope .work-card-id{
            font-family:'Geist Mono',monospace; font-size:.60rem; font-weight:700;
            letter-spacing:.14em; text-transform:uppercase; color:#5fa8d8; margin-bottom:4px;
        }
        .hist-scope .work-card-title{
            font-size:1.04rem; font-weight:800; color:#f1f5f9; line-height:1.18;
            letter-spacing:-.01em;
        }
        .hist-scope .work-card-subtitle{
            font-size:.73rem; color:#7c8ba1; margin-top:4px; font-weight:500;
        }
        .hist-scope .work-badge-row{
            display:flex; gap:6px; flex-wrap:wrap; align-items:center; justify-content:flex-end;
        }
        .hist-scope .work-card-desc{
            margin-top:12px; padding-top:11px; border-top:1px solid rgba(148,163,184,.09);
            font-size:.82rem; color:#aebbcd; line-height:1.52;
        }
        /* ── Cabeçalho de perfil da pessoa ── */
        .hist-profile{
            display:flex; align-items:center; gap:16px; flex-wrap:wrap;
            margin:14px 0 18px; padding:18px 20px; border-radius:16px;
            background:linear-gradient(135deg,rgba(56,189,248,.10),rgba(167,139,250,.08));
            border:1px solid rgba(56,189,248,.18);
        }
        .hist-avatar{
            width:56px; height:56px; border-radius:14px; flex:0 0 auto;
            display:flex; align-items:center; justify-content:center;
            font-size:1.3rem; font-weight:800; color:#04121f;
            background:linear-gradient(135deg,#38bdf8,#7c3aed);
            box-shadow:0 8px 20px rgba(56,189,248,.30);
        }
        .hist-profile-info{flex:1 1 240px; min-width:0;}
        .hist-profile-eyebrow{font-family:'Geist Mono',monospace; font-size:.58rem;
            font-weight:700; letter-spacing:.16em; text-transform:uppercase; color:#5fa8d8;}
        .hist-profile-name{font-size:1.45rem; font-weight:800; color:#f1f5f9;
            line-height:1.15; margin:2px 0 4px; letter-spacing:-.01em;}
        .hist-profile-sub{font-size:.84rem; color:#aebbcd;}
        .hist-profile-stats{display:flex; gap:18px; flex:0 0 auto;}
        .hps{text-align:center;}
        .hps-v{font-size:1.5rem; font-weight:800; color:#e2e8f0; line-height:1;}
        .hps-l{font-size:.58rem; font-weight:700; letter-spacing:.10em;
            text-transform:uppercase; color:#7c8ba1; margin-top:4px;}
        /* ── Métricas ricas dentro do card ── */
        .hist-scope .hist-metrics{display:grid; gap:8px; margin-top:12px;
            grid-template-columns:repeat(auto-fit,minmax(118px,1fr));}
        .hist-scope .hm{display:flex; flex-direction:column; gap:2px; padding:8px 10px;
            border-radius:9px; background:rgba(4,8,18,.5); border:1px solid rgba(148,163,184,.09);}
        .hist-scope .hm-wide{grid-column:1/-1;}
        .hist-scope .hm-l{font-size:.55rem; font-weight:700; letter-spacing:.09em;
            text-transform:uppercase; color:#5f7088;}
        .hist-scope .hm-v{font-size:.82rem; font-weight:600; color:#cbd5e1;}
        /* ── Caixa de busca de colaborador (minimal) ── */
        .hist-search-head{display:flex; align-items:center; gap:8px; margin-bottom:7px;}
        .hist-search-ico{width:24px; height:24px; border-radius:7px; flex:0 0 auto;
            display:flex; align-items:center; justify-content:center; font-size:.78rem;
            background:linear-gradient(135deg,rgba(56,189,248,.20),rgba(124,58,237,.18));
            border:1px solid rgba(56,189,248,.26);}
        .hist-search-ttl{font-size:.74rem; font-weight:600; color:#cdd9e8; line-height:1;
            letter-spacing:.02em;}
        .hist-search-hint{font-size:.72rem; color:#7c8ba1; margin-top:1px;}
        .hist-hit{display:inline-flex; align-items:center; gap:6px; font-size:.74rem;
            color:#9fb4cc; background:rgba(56,189,248,.08); border:1px solid rgba(56,189,248,.18);
            padding:4px 10px; border-radius:999px;}
        .hist-hit strong{color:#e2e8f0; font-weight:700;}

        /* ═══ CAIXA DE BUSCA HIGH-TECH (escopo via placeholder único) ═══ */
        [data-testid="stVerticalBlockBorderWrapper"]:has(input[placeholder^="Ex.: jo"]){
            position:relative !important;
            border:1px solid rgba(56,189,248,.20) !important;
            border-radius:13px !important;
            padding:12px 14px !important;
            background:
                radial-gradient(130% 150% at 0% 0%, rgba(56,189,248,.12), transparent 55%),
                radial-gradient(130% 150% at 100% 0%, rgba(124,58,237,.14), transparent 55%),
                linear-gradient(145deg, rgba(10,17,29,.97), rgba(6,11,20,.97)) !important;
            box-shadow:0 20px 55px -20px rgba(0,0,0,.75), inset 0 1px 0 rgba(255,255,255,.05) !important;
            overflow:hidden !important;
            transition:border-color .25s ease, box-shadow .25s ease !important;
        }
        /* linha de varredura no topo da caixa */
        [data-testid="stVerticalBlockBorderWrapper"]:has(input[placeholder^="Ex.: jo"])::before{
            content:"" !important; position:absolute !important; left:0; right:0; top:0; height:2px;
            background:linear-gradient(90deg, transparent, #38bdf8 40%, #7c3aed 60%, transparent);
            background-size:220% 100%; animation:hsScan 3.4s linear infinite; opacity:.95; z-index:2;
        }
        @keyframes hsScan{0%{background-position:220% 0}100%{background-position:-220% 0}}
        [data-testid="stVerticalBlockBorderWrapper"]:has(input[placeholder^="Ex.: jo"]):focus-within{
            border-color:rgba(56,189,248,.55) !important;
            box-shadow:0 0 0 1px rgba(56,189,248,.28), 0 26px 64px -22px rgba(56,189,248,.40) !important;
        }
        /* ícone com pulso de radar */
        .hist-search-ico{position:relative;}
        .hist-search-ico::after{
            content:""; position:absolute; inset:-3px; border-radius:13px;
            border:1px solid rgba(56,189,248,.40); animation:hsPulse 2.4s ease-out infinite;
        }
        @keyframes hsPulse{0%{transform:scale(.92);opacity:.7}100%{transform:scale(1.55);opacity:0}}
        /* campo de texto da busca */
        div[data-baseweb="base-input"]:has(input[placeholder^="Ex.: jo"]),
        div[data-baseweb="input"]:has(input[placeholder^="Ex.: jo"]){
            border-radius:10px !important; min-height:42px !important;
            background:linear-gradient(160deg,#0a1626,#060f1c) !important;
            border:1px solid rgba(56,189,248,.24) !important;
        }
        input[placeholder^="Ex.: jo"]{ font-size:.92rem !important; }
        input[placeholder^="Ex.: jo"]::placeholder{
            font-family:'Geist Mono',monospace !important; letter-spacing:.05em !important;
            color:rgba(120,160,200,.45) !important;
        }

        /* ═══ ESTADO VAZIO ESTILO RADAR ═══ */
        .hist-empty{
            position:relative; margin-top:10px; padding:26px 22px 22px; border-radius:14px;
            text-align:center; overflow:hidden;
            background:
                radial-gradient(80% 120% at 50% -8%, rgba(56,189,248,.12), transparent 60%),
                linear-gradient(160deg, rgba(9,15,26,.92), rgba(5,10,18,.94));
            border:1px solid rgba(56,189,248,.16);
            box-shadow:inset 0 1px 0 rgba(255,255,255,.04), 0 24px 60px -28px rgba(0,0,0,.8);
        }
        .hist-empty::before{
            content:""; position:absolute; inset:0; opacity:.55;
            background-image:linear-gradient(rgba(56,189,248,.06) 1px, transparent 1px),
                             linear-gradient(90deg, rgba(56,189,248,.06) 1px, transparent 1px);
            background-size:36px 36px;
            -webkit-mask-image:radial-gradient(circle at 50% 38%, #000 0%, transparent 68%);
            mask-image:radial-gradient(circle at 50% 38%, #000 0%, transparent 68%);
        }
        .hist-empty-radar{position:relative; width:64px; height:64px; margin:0 auto 12px;}
        .hist-empty-radar::before, .hist-empty-radar::after{
            content:""; position:absolute; inset:0; border-radius:50%;
            border:1px solid rgba(56,189,248,.40);
        }
        .hist-empty-radar::before{animation:hsRing 2.8s ease-out infinite;}
        .hist-empty-radar::after{animation:hsRing 2.8s ease-out infinite 1.4s;}
        @keyframes hsRing{0%{transform:scale(.5);opacity:.85}100%{transform:scale(1.18);opacity:0}}
        .hist-empty-core{position:absolute; inset:16px; border-radius:50%;
            background:radial-gradient(circle at 50% 40%, rgba(56,189,248,.30), rgba(124,58,237,.16) 70%, transparent);
            display:flex; align-items:center; justify-content:center; font-size:1.3rem;
            filter:drop-shadow(0 0 12px rgba(56,189,248,.55));}
        .hist-empty-ttl{position:relative; font-size:.98rem; font-weight:700; color:#eaf2fc; letter-spacing:-.01em;}
        .hist-empty-sub{position:relative; font-size:.82rem; color:#90a6bf; margin-top:5px;}
        .hist-empty-pills{position:relative; margin-top:16px; display:flex; gap:7px;
            justify-content:center; flex-wrap:wrap;}

        /* ── Hero e cabeçalho de seção compactos (só nesta página) ── */
        .hero-compact{padding:15px 22px 16px !important; margin-bottom:14px !important;}
        .hero-compact .page-hero-eyebrow{margin-bottom:6px;}
        .hero-compact .page-hero-desc{font-size:.84rem; line-height:1.45;}
        .hist-scope-sec{margin:2px 0 12px !important; padding-bottom:9px !important;}
        </style>
    """, unsafe_allow_html=True)

        # ── BUSCA INTELIGENTE POR TEXTO ───────────────────────────────────
        st.markdown('<div class="sec-label hist-scope-sec animate-fade-in"><div class="sec-dot"></div>CONSULTA DE REGISTROS</div>', unsafe_allow_html=True)

        import unicodedata as _ud
        def _norm_nome(s):
            """Normaliza o nome: remove acentos, caixa e espaços extras."""
            s = str(s or "").strip().lower()
            s = _ud.normalize("NFD", s)
            s = "".join(c for c in s if _ud.category(c) != "Mn")
            return " ".join(s.split())

        with st.container(border=True):
            st.markdown("""
            <div class="hist-search-head">
                <span class="hist-search-ico">🔎</span>
                <span class="hist-search-ttl">Buscar pelo seu nome</span>
            </div>
            """, unsafe_allow_html=True)
            _q = st.text_input(
                "Buscar pelo seu nome",
                key="hist_query",
                placeholder="Ex.: joão  ·  silva  ·  maria souza",
                label_visibility="collapsed",
            )

        # Tokens da busca: cada pedaço digitado precisa aparecer no nome.
        _qtokens = _norm_nome(_q).split()
        nome_busca = _q.strip()

        # ── ESTADO VAZIO (sem busca digitada) ────────────────────────────
        if not _qtokens:
            st.markdown("""
            <div class="hist-empty animate-fade-in">
                <div class="hist-empty-radar">
                    <div class="hist-empty-core">🔍</div>
                </div>
                <div class="hist-empty-ttl">Pronto para consultar</div>
                <div class="hist-empty-sub">Digite seu nome na busca acima para ver seus chamados e solicitações.</div>
                <div class="hist-empty-pills">
                    <span class="pill pill-red">🔴 Chamados abertos</span>
                    <span class="pill pill-amber">🟡 Em atendimento</span>
                    <span class="pill pill-green">✅ Concluídos</span>
                    <span class="pill pill-gray">🛒 Compras</span>
                </div>
            </div>
        """, unsafe_allow_html=True)

        else:
            # ════════════════════════════════════════════════════════════════
            # DINÂMICA: perfil da pessoa + linha do tempo unificada
            # (chamados + compras juntos, do mais recente ao mais antigo)
            # ════════════════════════════════════════════════════════════════
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

            # Busca por texto: o registro entra se TODOS os trechos digitados
            # aparecem no nome normalizado (ex.: "joao" acha "joão silva").
            def _bate(_h):
                return all(_t in _h for _t in _qtokens)
            df_user = df_base[df_base['Usuário'].map(_norm_nome).map(_bate)] if not df_base.empty else df_base.iloc[0:0]
            df_user_c = df_compras[df_compras['Usuário'].map(_norm_nome).map(_bate)] if not df_compras.empty else df_compras.iloc[0:0]

            # nome canônico para exibição (variação mais frequente)
            _variantes = []
            if not df_user.empty:
                _variantes += list(df_user['Usuário'].astype(str))
            if not df_user_c.empty:
                _variantes += list(df_user_c['Usuário'].astype(str))
            _nome_disp = (max(set(_variantes), key=_variantes.count).strip() if _variantes else nome_busca.strip())

            # métricas pessoais
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
                _resumo = "Tudo em dia — nenhum chamado em aberto. 🎉"
            else:
                _resumo = "Nenhum registro encontrado para este nome."
            _iniciais = ("".join(p[0] for p in _nome_disp.split()[:2]).upper() or "?")

            # ── Count-up: anima os números de 0 até o valor (CSS @property) ──
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
                    <div class="hps"><div class="hps-v" style="color:#34d399">{_v_tx}</div><div class="hps-l">Resolvidos</div></div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            if tot_ch == 0 and tot_cp == 0:
                st.markdown("""
                <div class="empty-state animate-fade-in" style="padding:32px 20px;margin-top:10px;">
                    <div style="font-size:2.4rem;margin-bottom:10px;">🔎</div>
                    <strong>Nenhum registro encontrado</strong>
                    <span>Não há chamados ou compras vinculados a este nome.</span>
                </div>
                """, unsafe_allow_html=True)
                return

            # ── FILTRO: meus chamados por status ─────────────────────────────
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

            # eventos: quando há filtro de status ativo, mostra SÓ os chamados
            # daquele status (sem misturar compras). Sem filtro, mostra tudo junto.
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
                st.caption(f"⚠️ A busca encontrou {len(_distintos)} nomes diferentes. Digite mais do nome (ex.: nome + sobrenome) para refinar.")

            if _status_filtro and not _eventos:
                st.markdown(f"""
                <div class="empty-state animate-fade-in" style="padding:26px 20px;margin-top:6px;">
                    <div style="font-size:2rem;margin-bottom:8px;">🔍</div>
                    <strong>Nenhum chamado com status "{escapar_html(_status_filtro)}"</strong>
                    <span>Tente outro status no filtro acima.</span>
                </div>
                """, unsafe_allow_html=True)

            for _tipo, _data_dt, idx, row in _eventos[:40]:
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
                    tecnico = html.escape(str(row.get('Técnico Responsável', 'Não atribuído')))
                    resposta = str(row.get('Resposta do Técnico', '—'))
                    _hs = "hs-open" if status_atual == "Aberto" else ("hs-prog" if status_atual == "Em Atendimento" else "hs-done")
                    _ico, _tipo_lbl = "🎫", "Chamado"
                    status_pill = ('<span class="pill pill-red">🔴 Aberto</span>' if status_atual == "Aberto"
                                   else '<span class="pill pill-amber">🟡 Em Atendimento</span>' if status_atual == "Em Atendimento"
                                   else '<span class="pill pill-green">✅ Concluído</span>')
                    prox = ("Aguardando triagem/atribuição da TI." if status_atual == "Aberto"
                            else "Técnico trabalhando — aguarde retorno." if status_atual == "Em Atendimento"
                            else "Resolvido. Nenhuma ação necessária.")
                    _meta_extra = f'⚡ {escapar_html(prio)} • 🗂 {escapar_html(categoria)}'
                    titulo = desc_preview or "(sem descrição)"
                else:
                    urg = str(row.get('Urgência', '—'))
                    item = str(row.get('Item', '—')).strip()
                    desc_preview = item[:160] + '...' if len(item) > 160 else item
                    resposta = None
                    tecnico = None
                    _hs = ("hs-done" if any(s in status_atual for s in ["Aprovado", "Concluído"])
                           else "hs-open" if "Recusado" in status_atual else "hs-prog")
                    _ico, _tipo_lbl = "🛒", "Compra"
                    status_pill = ('<span class="pill pill-gray">⏳ Pendente</span>' if "Pendente" in status_atual
                                   else '<span class="pill pill-amber">🔍 Em Análise</span>' if "Análise" in status_atual
                                   else '<span class="pill pill-blue">✅ Aprovado</span>' if "Aprovado" in status_atual
                                   else '<span class="pill pill-red">❌ Recusado</span>' if "Recusado" in status_atual
                                   else '<span class="pill pill-green">🟢 Concluído</span>')
                    prox = ("Aguardando análise da TI." if ("Pendente" in status_atual or "Análise" in status_atual)
                            else "Aprovada — acompanhe a entrega." if "Aprovado" in status_atual
                            else "Recusada — verifique a justificativa." if "Recusado" in status_atual
                            else "Concluída.")
                    _meta_extra = f'📦 {escapar_html(str(row.get("Quantidade", "—")))} • 🏷️ {escapar_html(str(row.get("Categoria", "—")))}'
                    titulo = desc_preview or "(sem item)"

                # cor da barra = mesma cor do status usado nos filtros (🔴/🟡/✅)
                _hs_color = {"hs-open": "#f87171", "hs-prog": "#fbbf24", "hs-done": "#34d399"}.get(_hs, "#38bdf8")

                with st.container(border=True):
                    st.markdown(f"""
                    <div class="hist-scope">
                    <div class="queue-card {_hs}" style="--hs-accent:{_hs_color};">
                        <div class="work-card-head">
                            <div>
                                <div class="work-card-id">{_ico} {_tipo_lbl} {escapar_html(protocolo)}</div>
                                <div class="work-card-title">{escapar_html(titulo)}</div>
                                <div class="work-card-subtitle">{escapar_html(_nome_disp)} • {escapar_html(data_str)}</div>
                            </div>
                            <div class="work-badge-row">
                                {status_pill}
                            </div>
                        </div>
                        <div class="hist-metrics">
                            <div class="hm"><span class="hm-l">Aberto há</span><span class="hm-v">{aberto_ha}</span></div>
                            <div class="hm"><span class="hm-l">No status há</span><span class="hm-v">{no_status_ha}</span></div>
                            <div class="hm hm-wide"><span class="hm-l">Próxima ação</span><span class="hm-v">{escapar_html(prox)}</span></div>
                        </div>
                        <div class="work-card-desc">{_meta_extra}</div>
                    </div>
                    </div>
                    """, unsafe_allow_html=True)

                    with st.expander("Ver detalhes", expanded=False):
                        if _tipo == "chamado":
                            _resp_txt = html.escape(resposta) if resposta and resposta not in ['—', 'nan', 'Aguardando'] else 'Aguardando resposta do técnico.'
                            st.markdown(f'<div class="meta-chip"><strong>🔧 Técnico:</strong> &nbsp;{tecnico}</div>', unsafe_allow_html=True)
                            st.write("")
                            st.markdown(f'<div class="meta-chip" style="height:auto;align-items:flex-start;padding:10px 12px;"><div><strong>💬 Resposta:</strong><br><span style="color:#94a3b8;font-size:.78rem;">{_resp_txt}</span></div></div>', unsafe_allow_html=True)
                            if desc_texto:
                                st.write("")
                                st.markdown(f'<div class="ux-note" style="margin-top:0;"><strong>📝 Descrição completa:</strong> {html.escape(desc_texto)}</div>', unsafe_allow_html=True)
                            _anexo = str(row.get('Anexo', ''))
                            if _anexo and _anexo not in ["Nenhum", "nan", ""]:
                                st.caption(f"📎 Anexo: {html.escape(_anexo)}")
                        else:
                            _just = html.escape(str(row.get('Justificativa', '—')))
                            _just_txt = _just if _just and _just != "—" else "Não informada."
                            st.markdown(f'<div class="meta-chip"><strong>📦 Quantidade:</strong> &nbsp;{html.escape(str(row.get("Quantidade", "—")))}</div>', unsafe_allow_html=True)
                            st.write("")
                            st.markdown(f'<div class="meta-chip" style="height:auto;align-items:flex-start;padding:10px 12px;"><div><strong>📝 Justificativa:</strong><br><span style="color:#94a3b8;font-size:.78rem;">{_just_txt}</span></div></div>', unsafe_allow_html=True)
            return

    # ═══════════════════════════════════════════════════════
    # 4. ABA: PAINEL DO GESTOR (ADMINISTRATIVO)
    # ═══════════════════════════════════════════════════════
    _render_historico()
elif pagina == "gestor":

    # Página isolada em fragment: interações reroda só este trecho, não a app inteira.
    @st.fragment
    def _render_gestor():
        global df_base, df_compras

        # ── CONSTANTES DE SEGURANÇA ────────────────────────────────────────
        _MAX_TENTATIVAS    = 5      # tentativas antes de bloquear
        _LOCKOUT_SEGUNDOS  = 300    # 5 minutos de bloqueio
        _SESSION_TTL       = 3600   # 1 hora de sessão autenticada

        # ── ESTADO INICIAL DE AUTENTICAÇÃO ────────────────────────────────
        if "gestor_autenticado"        not in st.session_state:
            st.session_state.gestor_autenticado   = False
        if "gestor_auth_ts"            not in st.session_state:
            st.session_state.gestor_auth_ts       = 0.0
        if "gestor_tentativas"         not in st.session_state:
            st.session_state.gestor_tentativas    = 0
        if "gestor_lockout_ate"        not in st.session_state:
            st.session_state.gestor_lockout_ate   = 0.0

        # ── EXPIRAÇÃO AUTOMÁTICA DA SESSÃO ────────────────────────────────
        _agora_ts = time.time()
        if st.session_state.gestor_autenticado:
            if (_agora_ts - st.session_state.gestor_auth_ts) > _SESSION_TTL:
                st.session_state.gestor_autenticado = False
                st.session_state.gestor_auth_ts     = 0.0

        # ── TELA DE AUTENTICAÇÃO ──────────────────────────────────────────
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
            background: rgba(3,7,16,.97);
            border: 1px solid rgba(56,189,248,.22);
            border-radius: 12px 12px 0 0;
            font-family: 'Geist Mono', ui-monospace, monospace;
            box-shadow: 0 0 50px rgba(56,189,248,.10), 0 28px 70px rgba(0,0,0,.70);
        }
        .tg-card.locked {
            border-color: rgba(248,113,113,.35) !important;
        }
        .tg-card::before {
            content:''; position:absolute; top:0; left:0; right:0; height:1px;
            background: linear-gradient(90deg,transparent,rgba(56,189,248,.7),transparent);
            background-size: 200% 1px; animation: _lockout_scan 3s linear infinite;
        }
        .tg-card.locked::before {
            background: linear-gradient(90deg,transparent,rgba(248,113,113,.8),transparent);
            background-size: 200% 1px;
        }
        .tg-bar {
            display:flex; align-items:center; gap:7px; padding:9px 16px;
            background:rgba(56,189,248,.04); border-bottom:1px solid rgba(56,189,248,.08);
        }
        .tg-bar.locked { background:rgba(248,113,113,.04); border-bottom-color:rgba(248,113,113,.08); }
        .tg-dot { width:8px; height:8px; border-radius:50%; }
        .tg-bar-title {
            flex:1; text-align:center;
            font-size:.50rem; letter-spacing:.22em; text-transform:uppercase;
            color:rgba(56,189,248,.30); margin:0;
        }
        .tg-bar.locked .tg-bar-title { color:rgba(248,113,113,.30); }
        .tg-body { padding:22px 24px 18px; }
        .tg-icon-ring {
            width:52px; height:52px; border-radius:13px;
            margin:0 auto 14px;
            display:flex; align-items:center; justify-content:center;
            background:linear-gradient(135deg,rgba(56,189,248,.15),rgba(167,139,250,.08));
            border:1px solid rgba(56,189,248,.22);
            font-size:1.5rem;
        }
        .tg-icon-ring.locked {
            background:linear-gradient(135deg,rgba(248,113,113,.15),rgba(251,191,36,.06));
            border-color:rgba(248,113,113,.30);
        }
        .tg-badge {
            display:inline-flex; align-items:center; gap:6px;
            font-size:.50rem; letter-spacing:.16em; text-transform:uppercase;
            color:rgba(56,189,248,.45); margin-bottom:8px;
        }
        .tg-badge.locked { color:rgba(248,113,113,.55); }
        .tg-badge-dot { width:5px; height:5px; border-radius:50%; background:#38bdf8; flex-shrink:0; animation:_blink_g 1.5s ease-in-out infinite; }
        .tg-badge-dot.locked { background:#f87171; }
        .tg-title { font-size:.96rem; font-weight:800; color:#dce8f5; margin-bottom:4px; text-align:center; }
        .tg-sub   { font-size:.62rem; color:rgba(148,163,184,.42); text-align:center; }
        .tg-sep { height:1px; margin:16px 0 14px; background:linear-gradient(90deg,transparent,rgba(56,189,248,.20),transparent); }
        .tg-sep.locked { background:linear-gradient(90deg,transparent,rgba(248,113,113,.18),transparent); }
        .tg-attempts { display:flex; align-items:center; gap:6px; margin-bottom:10px; justify-content:center; }
        .tg-att-dot { width:8px; height:8px; border-radius:50%; }
        .tg-att-dot.used  { background:#f87171; box-shadow:0 0 8px rgba(248,113,113,.6); }
        .tg-att-dot.avail { background:rgba(56,189,248,.25); border:1px solid rgba(56,189,248,.18); }
        .tg-lockout-bar {
            padding:11px 13px; border:1px solid rgba(248,113,113,.28); border-radius:8px;
            background:rgba(248,113,113,.06);
        }
        .tg-lockout-title { font-size:.62rem; font-weight:800; color:#f87171; letter-spacing:.08em; text-transform:uppercase; margin-bottom:3px; }
        .tg-lockout-msg   { font-size:.60rem; color:rgba(248,113,113,.65); }
        .tg-warn-token {
            padding:9px 12px; margin-bottom:10px;
            border:1px solid rgba(251,191,36,.24); border-radius:7px;
            background:rgba(251,191,36,.05);
            font-size:.60rem; color:rgba(251,191,36,.75); line-height:1.45;
        }
        .tg-warn-token strong { color:#fbbf24; }
        /* Form integrado abaixo do card */
        div[data-testid="stForm"][data-form-key="flg"] {
            background:rgba(3,7,16,.97) !important;
            border:1px solid rgba(56,189,248,.22) !important;
            border-top:none !important;
            border-radius:0 0 12px 12px !important;
            padding:4px 24px 18px !important;
            margin:0 !important;
            box-shadow:0 28px 70px rgba(0,0,0,.70) !important;
        }
        div[data-testid="stForm"][data-form-key="flg"] input[type="password"] {
            background:rgba(56,189,248,.04) !important;
            border:1px solid rgba(56,189,248,.18) !important;
            border-radius:7px !important; color:#7dd3fc !important;
            font-family:'Geist Mono',monospace !important;
            font-size:.78rem !important; letter-spacing:.06em !important;
            height:40px !important; padding:0 14px !important;
        }
        div[data-testid="stForm"][data-form-key="flg"] input[type="password"]:focus {
            border-color:rgba(56,189,248,.55) !important;
            box-shadow:0 0 0 3px rgba(56,189,248,.10) !important;
        }
        div[data-testid="stForm"][data-form-key="flg"] .stFormSubmitButton button {
            width:100% !important; height:34px !important;
            background:transparent !important;
            border:1px solid rgba(56,189,248,.18) !important;
            border-radius:6px !important; color:rgba(56,189,248,.55) !important;
            font-family:'Geist Mono',monospace !important;
            font-size:.56rem !important; font-weight:700 !important;
            letter-spacing:.18em !important; text-transform:uppercase !important;
            padding:0 !important; margin-top:10px !important;
            transition:all .18s !important;
        }
        div[data-testid="stForm"][data-form-key="flg"] .stFormSubmitButton button:hover {
            background:rgba(56,189,248,.08) !important;
            border-color:rgba(56,189,248,.42) !important;
            color:#38bdf8 !important;
            box-shadow:0 0 14px rgba(56,189,248,.14) !important;
        }
        .tg-footer-g {
            display:flex; justify-content:space-between; align-items:center;
            padding:10px 0 0;
            font-family:'Geist Mono',monospace; font-size:.46rem; letter-spacing:.10em;
        }
        </style>
        """, unsafe_allow_html=True)

            # Dots de tentativas
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
              <div class="tg-lockout-title">🔒 Acesso Bloqueado</div>
              <div class="tg-lockout-msg">Aguarde <strong style="color:#f87171;">{_seg_restantes}s</strong> para tentar novamente.</div>
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

            # Colunas para centralizar o card
            _ga, _gb, _gc = st.columns([1, 2.4, 1])
            with _gb:
                st.markdown(f"""
            <div class="tg-card {_locked_cls}">
              <div class="tg-bar {_locked_cls}">
                <div class="tg-dot" style="background:#f87171;opacity:.65;"></div>
                <div class="tg-dot" style="background:#fbbf24;opacity:.65;"></div>
                <div class="tg-dot" style="background:#34d399;opacity:.65;"></div>
                <span class="tg-bar-title">console · admin · gestor</span>
              </div>
              <div class="tg-body">
                <div class="tg-icon-ring {_locked_cls}">{"🔒" if _em_lockout else "🛡️"}</div>
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
                        _btn_submit = st.form_submit_button("→  autenticar", use_container_width=True)
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
                                    st.markdown(f"""<div style="font-family:'Geist Mono',monospace;font-size:.58rem;
                                    color:rgba(248,113,113,.9);text-align:center;padding:4px 0 0;">
                                    🔒 &nbsp;Bloqueado por {_LOCKOUT_SEGUNDOS//60} min.</div>""", unsafe_allow_html=True)
                                else:
                                    _rest = _MAX_TENTATIVAS - st.session_state.gestor_tentativas
                                    st.markdown(f"""<div style="font-family:'Geist Mono',monospace;font-size:.58rem;
                                    color:rgba(248,113,113,.85);text-align:center;padding:4px 0 0;">
                                    ✗ &nbsp;Token inválido — {_rest} tentativa(s) restante(s)</div>""", unsafe_allow_html=True)

                st.markdown(f"""
            <div class="tg-footer-g">
              <span style="color:rgba(56,189,248,.18);">IT ServiceDesk Pro · v2.1</span>
              <span style="color:rgba(52,211,153,.35);display:flex;align-items:center;gap:5px;">
                <span style="width:4px;height:4px;border-radius:50%;background:#34d399;
                  box-shadow:0 0 4px #34d399;display:inline-block;"></span>
                hmac · sha256
              </span>
            </div>
            """, unsafe_allow_html=True)

            st.stop()


        # ── CONTEÚDO LIBERADO (SÓ EXECUTA SE O GESTOR ESTIVER AUTENTICADO) ──
        _sessao_restante = max(0, int(_SESSION_TTL - (_agora_ts - st.session_state.gestor_auth_ts)))
        _sessao_min      = _sessao_restante // 60
        _agora_str       = datetime.now().strftime("%d/%m/%Y  %H:%M")

        # ── CSS EXCLUSIVO DA PÁGINA GESTOR ──
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
        background: linear-gradient(135deg, rgba(14,26,46,.98), rgba(8,15,30,.95));
        border: 1px solid rgba(56,189,248,.18);
        border-radius: 14px;
        overflow: hidden;
        box-shadow: 0 2px 40px rgba(0,0,0,.45), 0 0 0 1px rgba(56,189,248,.05);
    }
    .gestor-hero::before {
        content: "";
        position: absolute;
        top: 0; left: 0; right: 0; height: 2px;
        background: linear-gradient(90deg, #38bdf8, #a78bfa, #38bdf8);
        background-size: 200% 100%;
        animation: heroShimmer 4s linear infinite;
    }
    @keyframes heroShimmer { 0%{background-position:0} 100%{background-position:200%} }
    .gestor-hero::after {
        content: "";
        position: absolute;
        right: 0; top: 0; bottom: 0;
        width: 35%;
        background: radial-gradient(ellipse at right center, rgba(56,189,248,.06), transparent 70%);
        pointer-events: none;
    }
    .gestor-hero-left { position: relative; z-index: 1; }
    .gestor-hero-eyebrow {
        font-family: 'Geist Mono', monospace;
        font-size: .58rem; font-weight: 800;
        letter-spacing: .18em; text-transform: uppercase;
        color: rgba(56,189,248,.60);
        display: flex; align-items: center; gap: 8px;
        margin-bottom: 6px;
    }
    .gestor-hero-eyebrow::before {
        content: "";
        width: 6px; height: 6px; border-radius: 50%;
        background: #34d399;
        box-shadow: 0 0 10px rgba(52,211,153,.8);
        animation: gepulse 1.4s ease-in-out infinite;
    }
    @keyframes gepulse { 0%,100%{opacity:.9} 50%{opacity:.5} }
    .gestor-hero-title {
        font-size: 1.32rem; font-weight: 800;
        color: #f0f6ff; letter-spacing: -.01em;
        margin: 0 0 4px;
        text-shadow: 0 0 30px rgba(56,189,248,.14);
    }
    .gestor-hero-desc {
        font-size: .76rem; color: rgba(148,163,184,.62);
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
        background: rgba(52,211,153,.06);
        border: 1px solid rgba(52,211,153,.18);
        border-radius: 999px;
        font-family: 'Geist Mono', monospace;
        font-size: .58rem; font-weight: 700;
        letter-spacing: .08em; color: rgba(52,211,153,.75);
    }
    .gestor-session-badge::before {
        content: "";
        width: 5px; height: 5px; border-radius: 50%;
        background: #34d399;
        box-shadow: 0 0 6px rgba(52,211,153,.9);
    }
    .gestor-datetime {
        font-family: 'Geist Mono', monospace;
        font-size: .60rem; color: rgba(148,163,184,.38);
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
        border: 1px solid rgba(148,163,184,.10);
        border-radius: 12px;
        background: linear-gradient(145deg, rgba(12,22,38,.96), rgba(6,12,24,.88));
        box-shadow: 0 8px 28px rgba(0,0,0,.38), inset 0 1px 0 rgba(255,255,255,.025);
        transition: border-color .2s, transform .2s, box-shadow .2s;
        cursor: default;
    }
    .ops-card-v2:hover {
        border-color: var(--tone-c, rgba(56,189,248,.45));
        transform: translateY(-2px);
        box-shadow: 0 0 28px color-mix(in srgb, var(--tone-c, #38bdf8) 12%, transparent),
                    0 18px 44px rgba(0,0,0,.48);
    }
    .ops-card-v2::before {
        content: "";
        position: absolute;
        top: 0; left: 0; right: 0; height: 2.5px;
        background: var(--tone-c, #38bdf8);
        box-shadow: 0 0 16px var(--tone-c, #38bdf8);
        opacity: .85;
    }
    .ops-card-v2::after {
        content: "";
        position: absolute;
        bottom: 0; right: 0;
        width: 70px; height: 70px;
        background: radial-gradient(circle, color-mix(in srgb, var(--tone-c, #38bdf8) 8%, transparent), transparent 70%);
        pointer-events: none;
    }
    .ops-card-v2-icon {
        font-size: 1.4rem; margin-bottom: 10px; display: block; line-height: 1;
    }
    .ops-card-v2-label {
        font-family: 'Geist Mono', monospace;
        font-size: .60rem; font-weight: 800;
        letter-spacing: .12em; text-transform: uppercase;
        color: rgba(148,163,184,.55); margin-bottom: 6px;
    }
    .ops-card-v2-val {
        font-family: 'Geist Mono', monospace;
        font-size: 2.4rem; font-weight: 800;
        line-height: 1; color: var(--tone-c, #f0f6ff);
        margin-bottom: 5px;
        text-shadow: 0 0 20px color-mix(in srgb, var(--tone-c, #38bdf8) 30%, transparent);
    }
    .ops-card-v2-copy {
        font-size: .72rem; color: rgba(125,143,168,.60);
        line-height: 1.35;
    }
    /* ── INSIGHT BANNER V2 ── */
    .insight-banner {
        display: flex; align-items: center; gap: 14px;
        padding: 14px 18px;
        background: linear-gradient(135deg, rgba(56,189,248,.07), rgba(167,139,250,.04));
        border: 1px solid rgba(56,189,248,.15);
        border-radius: 10px;
        margin-bottom: 22px;
    }
    .insight-banner-icon {
        font-size: 1.1rem; flex-shrink: 0;
    }
    .insight-banner-text {
        font-size: .82rem; color: #b7c7dd; line-height: 1.5;
    }
    .insight-banner-text strong { color: #7dd3fc; }
    /* ── TABS V2 ── */
    [data-testid="stTabs"] > div:first-child {
        border-bottom: 1px solid rgba(56,189,248,.14) !important;
        gap: 2px !important;
    }
    [data-testid="stTabs"] button[role="tab"] {
        font-family: 'Geist Mono', monospace !important;
        font-size: .68rem !important;
        font-weight: 800 !important;
        letter-spacing: .10em !important;
        text-transform: uppercase !important;
        color: rgba(148,163,184,.55) !important;
        padding: 10px 20px !important;
        border-radius: 8px 8px 0 0 !important;
        transition: color .18s, background .18s !important;
        border: none !important;
    }
    [data-testid="stTabs"] button[role="tab"]:hover {
        color: #e2e8f0 !important;
        background: rgba(56,189,248,.06) !important;
    }
    [data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
        color: #38bdf8 !important;
        background: rgba(56,189,248,.10) !important;
        border-bottom: 2px solid #38bdf8 !important;
    }
    </style>
    """, unsafe_allow_html=True)

        # ── HERO BAR GESTOR ──
        st.markdown(f"""
    <div class="gestor-hero animate-fade-in">
        <div class="gestor-hero-left">
            <div class="gestor-hero-eyebrow">Operations Console · Acesso Gerencial</div>
            <div class="gestor-hero-title">Painel Administrativo</div>
            <div class="gestor-hero-desc">Gerencie a fila de trabalho, analise métricas em tempo real e decida sobre requisições logísticas.</div>
        </div>
        <div class="gestor-hero-right">
            <div class="gestor-session-badge">sessão ativa · {_sessao_min}min restantes</div>
            <div class="gestor-datetime">🕐 {_agora_str}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

        _col_logout_g = st.columns([6, 1])[1]
        with _col_logout_g:
            if st.button("🔓 Encerrar Sessão", key="gestor_logout", use_container_width=True, help="Encerrar sessão administrativa"):
                st.session_state.gestor_autenticado = False
                st.session_state.gestor_auth_ts     = 0.0
                st.rerun(scope="app")

        # Carga de dados dinâmicos em tempo de execução
        df_base_live = carregar_dados_cacheados(DATABASE, COLUNAS)
        df_compras_live = carregar_dados_cacheados(DATABASE_COMPRAS, COLUNAS_COMPRAS)

        if "qtd_chamados_antiga" not in st.session_state:
            st.session_state.qtd_chamados_antiga = len(df_base_live)
        if "qtd_compras_antiga" not in st.session_state:
            st.session_state.qtd_compras_antiga = len(df_compras_live)

        if len(df_base_live) > st.session_state.qtd_chamados_antiga:
            ultimo_c = df_base_live.iloc[-1]
            st.toast(f"🚨 Novo Chamado de {ultimo_c['Usuário']}", icon="🔔")
            st.session_state.qtd_chamados_antiga = len(df_base_live)

        if len(df_compras_live) > st.session_state.qtd_compras_antiga:
            ultima_c = df_compras_live.iloc[-1]
            st.toast(f"🛒 Nova Compra de {ultima_c['Usuário']}", icon="🛍️")
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

        st.markdown('<div class="sec-label animate-fade-in"><div class="sec-dot"></div>RADAR OPERACIONAL — TEMPO REAL</div>', unsafe_allow_html=True)
        _crit_color = "#f87171" if crit_ch > 0 else "#34d399"
        st.markdown(f"""
        <div class="ops-grid-v2 animate-fade-in">
            <div class="ops-card-v2" style="--tone-c:{_crit_color}">
                <span class="ops-card-v2-icon">🚨</span>
                <div class="ops-card-v2-label">Chamados Críticos</div>
                <div class="ops-card-v2-val">{crit_ch}</div>
                <div class="ops-card-v2-copy">Chamados com prioridade crítica na fila ativa.</div>
            </div>
            <div class="ops-card-v2" style="--tone-c:#fbbf24">
                <span class="ops-card-v2-icon">⚠️</span>
                <div class="ops-card-v2-label">Alta Prioridade</div>
                <div class="ops-card-v2-val">{alta_ch}</div>
                <div class="ops-card-v2-copy">Chamados de alta prioridade aguardando atendimento.</div>
            </div>
            <div class="ops-card-v2" style="--tone-c:#38bdf8">
                <span class="ops-card-v2-icon">🎫</span>
                <div class="ops-card-v2-label">Fila Ativa</div>
                <div class="ops-card-v2-val">{len(df_ch_ativos)}</div>
                <div class="ops-card-v2-copy">Tickets abertos ou em atendimento agora.</div>
            </div>
            <div class="ops-card-v2" style="--tone-c:#a78bfa">
                <span class="ops-card-v2-icon">🛒</span>
                <div class="ops-card-v2-label">Compras Pendentes</div>
                <div class="ops-card-v2-val">{len(df_cp_ativos)}</div>
                <div class="ops-card-v2-copy">Requisições aguardando análise ou decisão de TI.</div>
            </div>
        </div>
        <div class="insight-banner animate-fade-in">
            <span class="insight-banner-icon">🎯</span>
            <div class="insight-banner-text">
                <strong>Próxima ação recomendada:</strong> {html.escape(prox_texto)}
            </div>
        </div>
    """, unsafe_allow_html=True)

        CATEGORIAS_CHAMADO = {
            "💻 Hardware (PC/Note)": [
                "computador", "pc", "notebook", "monitor", "tela",
                "hardware", "liga", "desliga", "travando", "lento", "memoria", "hd",
                "ssd", "fonte", "cabo", "gabinete", "placa", "bateria", "carregador"
            ],
            "🖱️ Periféricos": [
                "periferico", "perifericos", "teclado", "mouse", "headset", "fone",
                "fones", "webcam", "camera", "microfone", "mic", "hub", "usb",
                "adaptador", "dock", "mousepad", "scanner de mao", "leitor"
            ],
            "🌐 Internet/Rede": [
                "internet", "rede", "wifi", "wi-fi", "conexao", "conectividade",
                "cabo de rede", "sem rede", "sem internet", "ping", "vpn",
                "roteador", "switch", "dns", "ip", "ethernet", "lento", "lenta",
                "oscilando", "caindo"
            ],
            "📧 E-mail/Login": [
                "email", "e-mail", "outlook", "login", "senha", "usuario", "conta",
                "autenticacao", "bloqueado", "mfa", "codigo", "entrar",
                "credencial", "office", "microsoft", "gmail", "godaddy"
            ],
            "🖨️ Impressora": [
                "impressora", "imprimir", "impressao", "scanner", "digitalizar",
                "toner", "papel", "fila", "spooler", "driver", "xerox"
            ],
            "📑 Sistema/Software": [
                "sistema", "software", "programa", "aplicativo", "app", "erro",
                "bug", "tela branca", "sistema lento", "instalar", "atualizar",
                "planilha", "navegador", "chrome", "windows"
            ],
            "📱 Telefonia": [
                "telefonia", "telefone", "ramal", "celular", "voip", "chip",
                "sim", "simcard", "ligacao", "chamada", "discar", "fixo",
                "operadora", "linha telefonica", "sinal"
            ],
            "🔐 Acesso/Permissão": [
                "permissao", "permissoes", "pasta", "pastas", "compartilhada",
                "compartilhamento", "diretorio", "autorizacao", "autorizar",
                "liberar acesso", "grupo", "active directory", "nivel de acesso",
                "permitir", "negado"
            ],
            "❓ Outros": [
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

        def detectar_tipo_chamado(descricao, categoria_padrao="❓ Outros"):
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
            return categoria_padrao if categoria_padrao in CATEGORIAS_CHAMADO else "❓ Outros"

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

        # ── MODAL DINÂMICO (@st.dialog) PARA GERENCIAR TICKETS ──
        @st.dialog("📋 Despachar Ticket")
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
            st.caption(f"Anexo: {row['Anexo']}")
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
            if st.button("🧠 Gerar resposta técnica automática", use_container_width=True, key=f"btn_gerar_resposta_{idx}"):
                st.session_state[nota_key] = resposta_resolucao

            resp_tec = st.text_area(
                "Notas Técnicas de Resolução:",
                key=nota_key,
                height=220,
                placeholder="Insira o parecer técnico ou gere uma resposta técnica automática para este chamado..."
            )

            prio_colors = {
                "Baixa": "#38bdf8",
                "Média": "#fbbf24",
                "Alta": "#f87171",
                "Crítica": "#ef4444",
                "Urgente": "#ef4444"
            }
            st.markdown(
                f"<div style='text-align:right; padding-top:10px;'>Prioridade: "
                f"<strong style='color:{prio_colors.get(row['Prioridade'], '#fff')}'>{row['Prioridade']}</strong></div>",
                unsafe_allow_html=True
            )

            if st.button("💾 Gravar Alterações e Fechar", use_container_width=True, type="primary"):
                df_base.at[idx, "Status"] = novo_status
                df_base.at[idx, "Categoria"] = tipo_chamado
                df_base.at[idx, "Resposta do Técnico"] = resp_tec.strip() or "Aguardando"
                df_base.at[idx, "Técnico Responsável"] = tecnico_resp
                df_base.at[idx, "Atualizado em"] = datetime.now().strftime("%d/%m/%Y %H:%M")
                salvar_dataframe_seguro(df_base, DATABASE, COLUNAS)
                st.session_state.pop(nota_key, None)
                st.success("✅ Chamado atualizado com sucesso!")
                st.rerun(scope="app")

        # ── MODAL DINÂMICO (@st.dialog) PARA AUDITORIA DE COMPRAS ──
        @st.dialog("🛒 Analisar Requisição de Compra")
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
        
            if st.button("💾 Registrar Parecer e Fechar", use_container_width=True, type="primary"):
                df_compras.at[idx_c, "Status"] = novo_status_c
                df_compras.at[idx_c, "Observação TI"] = obs_ti
                df_compras.at[idx_c, "Responsável"] = responsavel_compra.strip() or "Não atribuído"
                df_compras.at[idx_c, "Atualizado em"] = datetime.now().strftime("%d/%m/%Y %H:%M")
                salvar_dataframe_seguro(df_compras, DATABASE_COMPRAS, COLUNAS_COMPRAS)
                st.success("✅ Parecer gravado com sucesso!")
                st.rerun(scope="app")

        # ── INTERFACE DO PAINEL GESTOR COM TABS ──
        # ── BOTÃO DE REFRESH MANUAL ──
        st.markdown("""
<style>
    .refresh-bar {
        display: flex; align-items: center; justify-content: flex-end;
        gap: 10px; margin: 0 0 10px 0;
    }
    .refresh-ts {
        font-family: var(--mono, monospace);
        font-size: 0.72rem;
        color: rgba(125,143,168,.65);
        letter-spacing: 0.04em;
    }
    div[data-testid="stButton"] button[kind="secondary"].refresh-btn {
        background: transparent !important;
        border: 1px solid rgba(56,189,248,0.22) !important;
        color: rgba(56,189,248,.85) !important;
        font-family: var(--mono, monospace) !important;
        font-size: 0.72rem !important;
        letter-spacing: 0.08em !important;
        padding: 2px 14px !important;
        border-radius: 6px !important;
        transition: border-color .2s, box-shadow .2s, color .2s !important;
    }
    div[data-testid="stButton"] button[kind="secondary"].refresh-btn:hover {
        border-color: rgba(56,189,248,.7) !important;
        box-shadow: 0 0 10px rgba(56,189,248,.25) !important;
        color: #38bdf8 !important;
    }
    </style>
    """, unsafe_allow_html=True)
        _ref_col_l, _ref_col_r = st.columns([6, 1])
        with _ref_col_r:
            if st.button("↻  SYNC", key="btn_refresh_manual", help="Recarregar dados do banco SQLite agora", use_container_width=True):
                carregar_dados_cacheados.clear()
                st.rerun(scope="app")
        with _ref_col_l:
            _ts_atual = datetime.now().strftime("%d/%m/%Y  %H:%M:%S")
            st.markdown(f'<div class="refresh-ts">última atualização · {_ts_atual}</div>', unsafe_allow_html=True)

        st.markdown('<div class="sec-label animate-fade-in"><div class="sec-dot"></div>MÓDULOS DE GESTÃO</div>', unsafe_allow_html=True)
        tab_analytics, tab_chamados, tab_compras = st.tabs([
            "📊  Analytics & Indicadores",
            "🛠️  Governança de Chamados",
            "🛒  Auditoria de Compras",
        ])

        # ── FRAGMENT COM ANALYTICS APRIMORADO E ANIMADO ──
        @st.fragment
        def renderizar_analytics_realtime():
            df_b_all = carregar_dados_cacheados(DATABASE, COLUNAS)
            df_c_all = carregar_dados_cacheados(DATABASE_COMPRAS, COLUNAS_COMPRAS)

            # ── FILTRO DE PERÍODO (MÊS / ANO) ───────────────────────────
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
                <span style="font-family:var(--mono,monospace);font-size:.72rem;color:#38bdf8;
                border:1px solid rgba(56,189,248,.28);padding:6px 14px;border-radius:7px;
                background:rgba(56,189,248,.06);white-space:nowrap;">
                📅 Exibindo: <strong>{_label_periodo}</strong> — zero-based por mês</span></div>""",
                    unsafe_allow_html=True
                )
            with _fc4:
                # Exportar período filtrado — geração cacheada (só recalcula quando muda o período/dados)
                _zip_bytes = _gerar_export_periodo_zip(
                    df_b_all, df_c_all, _mes_sel, _ano_sel, _meses_pt[_mes_sel]
                )
                st.download_button(
                    "📥 Exportar período",
                    data=_zip_bytes,
                    file_name=f"analytics_{_meses_pt[_mes_sel].lower()}_{_ano_sel}.zip",
                    mime="application/zip",
                    use_container_width=True,
                    key=f"dl_period_{_mes_sel}_{_ano_sel}"
                )

            # Filtrar dados pelo mês/ano selecionado (cacheado — evita re-parsear datas a cada rerun)
            df_b_live = _filtrar_periodo_cacheado(df_b_all, _mes_sel, _ano_sel)
            df_c_live = _filtrar_periodo_cacheado(df_c_all, _mes_sel, _ano_sel)

            # Aviso visual se o período não tiver dados
            if df_b_live.empty and df_c_live.empty:
                st.markdown(
                    f"""<div style="padding:14px 18px;border:1px dashed rgba(56,189,248,.28);
                border-radius:10px;background:rgba(56,189,248,.04);margin:8px 0 16px;">
                <strong style="color:#38bdf8;">Sem dados em {_label_periodo}</strong>
                <span style="color:#7d8fa8;font-size:.82rem;display:block;margin-top:4px;">
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

            # ── KPIs OPERACIONAIS — HEALTH CARDS COM CONTEXT ──
            import streamlit.components.v1 as _comp

            taxa_resol_num = int(concluidos / total_ch * 100) if total_ch > 0 else 0
            pressao_fila = round((abertos_ch + em_atend_ch) / max(total_ch, 1) * 100)
            saude_label = "Crítica" if pressao_fila > 70 else ("Atenção" if pressao_fila > 40 else "Saudável")
            saude_cor = "#f87171" if pressao_fila > 70 else ("#fbbf24" if pressao_fila > 40 else "#34d399")
            saude_icon = "🔴" if pressao_fila > 70 else ("🟡" if pressao_fila > 40 else "🟢")

            kpi_html = f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
*{{box-sizing:border-box;margin:0;padding:0;font-family:'Geist','Inter',sans-serif}}
body{{background:transparent;padding:0 0 4px}}
.row{{display:grid;grid-template-columns:1.8fr 1fr 1fr 1fr 1fr 1fr;gap:10px}}

/* ── HEALTH CARD (first) ── */
.health{{
  background:linear-gradient(145deg,rgba(12,24,40,.98),rgba(6,11,21,.94));
  border:1px solid rgba(56,189,248,.14);border-radius:14px;
  padding:16px 18px;position:relative;overflow:hidden;
  display:flex;flex-direction:column;gap:8px;
}}
.health::before{{content:'';position:absolute;top:0;left:0;right:0;height:2px;
  background:linear-gradient(90deg,#38bdf8,#a78bfa);
  box-shadow:0 0 14px rgba(56,189,248,.5)}}
.health-top{{display:flex;align-items:center;justify-content:space-between}}
.health-title{{font-size:.58rem;text-transform:uppercase;letter-spacing:.14em;color:#4a6480;font-weight:800}}
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
.ring-pct{{font-size:.95rem;font-weight:800;color:#e2e8f0;font-family:monospace;line-height:1}}
.ring-sub{{font-size:.48rem;text-transform:uppercase;letter-spacing:.1em;color:#3d5068;margin-top:1px}}
.health-stats{{flex:1;display:flex;flex-direction:column;gap:5px}}
.stat-row{{display:flex;align-items:center;justify-content:space-between;gap:6px}}
.stat-label{{font-size:.62rem;color:#7d8fa8;font-weight:600}}
.stat-bar{{flex:1;height:4px;background:rgba(255,255,255,.06);border-radius:4px;overflow:hidden}}
.stat-fill{{height:100%;border-radius:4px;transition:width 1s cubic-bezier(.4,0,.2,1)}}
.stat-val{{font-size:.62rem;font-weight:800;font-family:monospace;min-width:20px;text-align:right}}

/* ── KPI CARD (others) ── */
.k{{
  background:linear-gradient(145deg,rgba(12,24,40,.96),rgba(6,11,21,.90));
  border:1px solid rgba(56,189,248,.10);border-radius:14px;
  padding:14px 14px 12px;position:relative;overflow:hidden;
  transition:transform .25s,border-color .25s,box-shadow .25s;cursor:default;
  display:flex;flex-direction:column;justify-content:space-between;
}}
.k:hover{{
  transform:translateY(-3px);
  border-color:color-mix(in srgb,var(--c) 50%,transparent);
  box-shadow:0 0 20px color-mix(in srgb,var(--c) 18%,transparent),0 12px 28px rgba(0,0,0,.5);
}}
.k::before{{content:'';position:absolute;top:0;left:0;right:0;height:2px;
  background:var(--c);opacity:.7;
  box-shadow:0 0 10px var(--c)}}
.k-top{{display:flex;align-items:flex-start;justify-content:space-between}}
.k-ico{{font-size:1.1rem;line-height:1}}
.k-trend{{
  font-size:.55rem;font-weight:800;letter-spacing:.06em;padding:2px 7px;border-radius:999px;
  background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.07);color:#3d5068;
}}
.k-v{{
  font-family:'Geist Mono','Courier New',monospace;font-size:2rem;font-weight:800;
  color:var(--c);line-height:1;margin:6px 0 2px;letter-spacing:-.03em;
}}
.k-l{{font-size:.6rem;text-transform:uppercase;letter-spacing:.11em;color:#3d5068;font-weight:700}}
.k-sub{{font-size:.58rem;color:#4a6480;margin-top:4px;font-weight:500}}
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
          <circle cx="34" cy="34" r="28" fill="none" stroke="rgba(255,255,255,.06)" stroke-width="7"/>
          <circle cx="34" cy="34" r="28" fill="none"
            stroke="{saude_cor}" stroke-width="7"
            stroke-linecap="round"
            stroke-dasharray="{round(2*3.14159*28*taxa_resol_num/100, 1)} {round(2*3.14159*28*(1-taxa_resol_num/100), 1)}"
            style="filter:drop-shadow(0 0 6px {saude_cor})"/>
        </svg>
        <div class="ring-label">
          <span class="ring-pct" id="rpct">0%</span>
          <span class="ring-sub">resolvido</span>
        </div>
      </div>
      <div class="health-stats">
        <div class="stat-row">
          <span class="stat-label">Fila ativa</span>
          <div class="stat-bar"><div class="stat-fill" style="width:{min(abertos_ch/max(total_ch,1)*100,100):.0f}%;background:#f87171"></div></div>
          <span class="stat-val" style="color:#f87171">{abertos_ch}</span>
        </div>
        <div class="stat-row">
          <span class="stat-label">Em atend.</span>
          <div class="stat-bar"><div class="stat-fill" style="width:{min(em_atend_ch/max(total_ch,1)*100,100):.0f}%;background:#fbbf24"></div></div>
          <span class="stat-val" style="color:#fbbf24">{em_atend_ch}</span>
        </div>
        <div class="stat-row">
          <span class="stat-label">Concluídos</span>
          <div class="stat-bar"><div class="stat-fill" style="width:{min(concluidos/max(total_ch,1)*100,100):.0f}%;background:#34d399"></div></div>
          <span class="stat-val" style="color:#34d399">{concluidos}</span>
        </div>
      </div>
    </div>
  </div>

  <!-- ── KPI: TOTAL CHAMADOS ── -->
  <div class="k" style="--c:#38bdf8">
    <div class="k-top">
      <span class="k-ico">🎫</span>
      <span class="k-trend">PERÍODO</span>
    </div>
    <div class="k-v" data-t="{total_ch}">0</div>
    <div class="k-l">Total de chamados</div>
    <div class="k-sub">volume do mês</div>
  </div>

  <!-- ── KPI: FILA ABERTA ── -->
  <div class="k" style="--c:#f87171">
    <div class="k-top">
      <span class="k-ico">🔴</span>
      <span class="k-trend">URGENTE</span>
    </div>
    <div class="k-v" data-t="{abertos_ch}">0</div>
    <div class="k-l">Aguardando</div>
    <div class="k-sub">{"acima do ideal" if abertos_ch > 5 else "dentro do esperado"}</div>
  </div>

  <!-- ── KPI: EM ATENDIMENTO ── -->
  <div class="k" style="--c:#fbbf24">
    <div class="k-top">
      <span class="k-ico">⚙️</span>
      <span class="k-trend">ATIVO</span>
    </div>
    <div class="k-v" data-t="{em_atend_ch}">0</div>
    <div class="k-l">Em atendimento</div>
    <div class="k-sub">técnico alocado</div>
  </div>

  <!-- ── KPI: CONCLUÍDOS ── -->
  <div class="k" style="--c:#34d399">
    <div class="k-top">
      <span class="k-ico">✅</span>
      <span class="k-trend">OK</span>
    </div>
    <div class="k-v" data-t="{concluidos}">0</div>
    <div class="k-l">Encerrados</div>
    <div class="k-sub">resolvidos no período</div>
  </div>

  <!-- ── KPI: COMPRAS ── -->
  <div class="k" style="--c:#a78bfa">
    <div class="k-top">
      <span class="k-ico">🛒</span>
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
            _comp.html(kpi_html, height=142, scrolling=False)

            import streamlit.components.v1 as components
            import json

            st.markdown('<div class="sec-label animate-fade-in"><div class="sec-dot"></div>MÉTRICAS OPERACIONAIS // LIVE DATA</div>', unsafe_allow_html=True)

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

                # Compute insight strings for the dashboard
                top_setor = df_setor.iloc[-1]['Setor'] if not df_setor.empty else "—"
                top_setor_v = int(df_setor.iloc[-1]['Chamados']) if not df_setor.empty else 0
                top_prio = df_prio.sort_values('Qtd',ascending=False).iloc[0] if not df_prio.empty else None
                top_prio_label = top_prio['Prioridade'] if top_prio is not None else "—"
                top_prio_v = int(top_prio['Qtd']) if top_prio is not None else 0
                max_dia = df_dia2.loc[df_dia2['Chamados'].idxmax(),'Dia'] if not df_dia2.empty else "—"
                max_dia_v = int(df_dia2['Chamados'].max()) if not df_dia2.empty else 0
                top_user = df_top2.iloc[-1]['Usuário'] if not df_top2.empty else "—"
                top_user_v = int(df_top2.iloc[-1]['Chamados']) if not df_top2.empty else 0
                # Status donut data
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
*{{box-sizing:border-box;margin:0;padding:0}}
html,body{{background:transparent;font-family:'Geist Mono',ui-monospace,monospace;overflow-x:hidden}}
.dash{{display:grid;grid-template-columns:1fr 1fr;gap:12px;padding:2px}}
.wide{{grid-column:1/-1}}
.card{{
  background:linear-gradient(145deg,rgba(10,18,32,.98),rgba(5,9,18,.94));
  border:1px solid rgba(56,189,248,.10);
  border-radius:12px;padding:15px 17px 13px;
  position:relative;overflow:hidden;
  box-shadow:0 8px 28px rgba(0,0,0,.45),inset 0 1px 0 rgba(255,255,255,.03);
}}
.card::before{{
  content:'';position:absolute;top:0;left:0;right:0;height:2px;
  background:linear-gradient(90deg,var(--accent,#38bdf8),color-mix(in srgb,var(--accent,#38bdf8) 30%,transparent));
  box-shadow:0 0 12px var(--accent,#38bdf8);opacity:.7;
}}
.card-hdr{{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;gap:8px}}
.card-title{{font-size:9.5px;font-weight:800;letter-spacing:.12em;text-transform:uppercase;color:#8aa4c2}}
.card-right{{display:flex;align-items:center;gap:6px}}
.badge{{
  font-size:8.5px;padding:2px 7px;border-radius:999px;
  background:color-mix(in srgb,var(--accent,#38bdf8) 10%,transparent);
  color:color-mix(in srgb,var(--accent,#38bdf8) 90%,white);
  border:1px solid color-mix(in srgb,var(--accent,#38bdf8) 25%,transparent);
  letter-spacing:.06em;font-weight:700;
}}
.insight{{
  display:inline-flex;align-items:center;gap:5px;
  font-size:8.5px;color:#4a6480;font-weight:600;letter-spacing:.04em;
  padding:2px 8px;border-radius:6px;
  background:rgba(255,255,255,.03);border:1px solid rgba(255,255,255,.05);
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:200px;
}}
.insight strong{{color:#7dd3fc;font-weight:800}}
/* donut row */
.donut-wrap{{display:flex;align-items:center;gap:20px;height:130px}}
.donut-box{{position:relative;width:130px;height:130px;flex-shrink:0}}
.donut-box canvas{{position:absolute;inset:0;width:100%!important;height:100%!important}}
.donut-legend{{flex:1;display:flex;flex-direction:column;gap:10px}}
.leg-row{{display:flex;align-items:center;gap:8px}}
.leg-dot{{width:8px;height:8px;border-radius:50%;flex-shrink:0}}
.leg-label{{font-size:10px;color:#7d8fa8;font-weight:600;flex:1}}
.leg-val{{font-size:11px;font-weight:800;color:#e2e8f0;font-family:monospace}}
.leg-pct{{font-size:9px;color:#3d5068;font-weight:700;margin-left:3px}}
/* chart containers com altura fixa */
.ch-box{{position:relative;width:100%}}
.ch-box canvas{{width:100%!important}}
</style>
</head>
<body>
<div class="dash">

  <div class="card" style="--accent:#38bdf8">
    <div class="card-hdr">
      <span class="card-title">📊 Status dos Chamados</span>
      <span class="badge">LIVE</span>
    </div>
    <div class="donut-wrap">
      <div class="donut-box"><canvas id="c0"></canvas></div>
      <div class="donut-legend">
        <div class="leg-row">
          <div class="leg-dot" style="background:#f87171;box-shadow:0 0 6px #f87171"></div>
          <span class="leg-label">Abertos</span>
          <span class="leg-val">{abertos_ch}</span>
          <span class="leg-pct">({round(abertos_ch/max(abertos_ch+em_atend_ch+concluidos,1)*100)}%)</span>
        </div>
        <div class="leg-row">
          <div class="leg-dot" style="background:#fbbf24;box-shadow:0 0 6px #fbbf24"></div>
          <span class="leg-label">Em atendimento</span>
          <span class="leg-val">{em_atend_ch}</span>
          <span class="leg-pct">({round(em_atend_ch/max(abertos_ch+em_atend_ch+concluidos,1)*100)}%)</span>
        </div>
        <div class="leg-row">
          <div class="leg-dot" style="background:#34d399;box-shadow:0 0 6px #34d399"></div>
          <span class="leg-label">Concluídos</span>
          <span class="leg-val">{concluidos}</span>
          <span class="leg-pct">({round(concluidos/max(abertos_ch+em_atend_ch+concluidos,1)*100)}%)</span>
        </div>
      </div>
    </div>
  </div>

  <div class="card" style="--accent:#ef4444">
    <div class="card-hdr">
      <span class="card-title">🚨 Nível de Criticidade</span>
      <div class="card-right">
        <span class="insight">Maior: <strong>{top_prio_label} ({top_prio_v})</strong></span>
        <span class="badge" style="background:rgba(239,68,68,.1);color:#f87171;border-color:rgba(239,68,68,.3)">RISCO</span>
      </div>
    </div>
    <div class="ch-box" style="height:160px"><canvas id="c2"></canvas></div>
  </div>

  <div class="card wide" style="--accent:#38bdf8">
    <div class="card-hdr">
      <span class="card-title">📅 Volume de Chamados por Dia</span>
      <div class="card-right">
        <span class="insight">Pico: <strong>{max_dia} ({max_dia_v} tickets)</strong></span>
        <span class="badge">TENDÊNCIA</span>
      </div>
    </div>
    <div class="ch-box" style="height:110px"><canvas id="c3"></canvas></div>
  </div>

  <div class="card" style="--accent:#38bdf8">
    <div class="card-hdr">
      <span class="card-title">🏢 Incidentes por Setor</span>
      <div class="card-right">
        <span class="insight">Líder: <strong>{top_setor} ({top_setor_v})</strong></span>
        <span class="badge">VOLUME</span>
      </div>
    </div>
    <div class="ch-box" style="height:200px"><canvas id="c1"></canvas></div>
  </div>

  <div class="card" style="--accent:#a78bfa">
    <div class="card-hdr">
      <span class="card-title">👤 Top Solicitantes</span>
      <div class="card-right">
        <span class="insight" style="color:#a78bfa">+ ativo: <strong style="color:#c4b5fd">{top_user} ({top_user_v}×)</strong></span>
        <span class="badge" style="background:rgba(167,139,250,.1);color:#c4b5fd;border-color:rgba(167,139,250,.3)">RANKING</span>
      </div>
    </div>
    <div class="ch-box" style="height:200px"><canvas id="c4"></canvas></div>
  </div>

  <div class="card wide" style="--accent:#a78bfa">
    <div class="card-hdr">
      <span class="card-title">📦 Demandas por Categoria de Compras</span>
      <span class="badge" style="background:rgba(167,139,250,.1);color:#c4b5fd;border-color:rgba(167,139,250,.3)">LOGÍSTICA</span>
    </div>
    <div class="ch-box" style="height:100px"><canvas id="c5"></canvas></div>
  </div>

</div>
<script>
Chart.defaults.color='#7d8fa8';
Chart.defaults.font.family="'Geist Mono',monospace";
Chart.defaults.font.size=10.5;
const SL={SL},SV={SV};
const PL={PL},PV={PV};
const TL={TL},TV={TV};
const UL={UL},UV={UV};
const CL={CL},CV={CV};
const STL={STL},STV={STV};
const tt={{
  backgroundColor:'rgba(3,6,14,0.97)',borderWidth:1,padding:10,
  titleFont:{{weight:'800',size:11}},bodyColor:'#e2e8f0',cornerRadius:8,displayColors:false,
}};

/* ── C0: DONUT STATUS ── */
(()=>{{
  const canvas=document.getElementById('c0');
  const ctx=canvas.getContext('2d');
  const CORES=['#f87171','#fbbf24','#34d399'];
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
        tooltip:{{...tt,borderColor:'rgba(56,189,248,.25)',titleColor:'#38bdf8',
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
        cx.fillStyle='#e2e8f0';cx.font='bold 22px Geist Mono,monospace';
        cx.fillText(total,cxc,cyc-8);
        cx.fillStyle='#3d5068';cx.font='700 9px Geist Mono,monospace';
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
      backgroundColor:'rgba(2,132,199,0.75)',
      hoverBackgroundColor:'#7dd3fc',
      borderRadius:{{topRight:5,bottomRight:5}},borderSkipped:false,barThickness:16,
    }}]}},
    options:{{
      indexAxis:'y',
      responsive:true,maintainAspectRatio:false,
      animation:{{duration:300,easing:'easeOutQuart'}},
      plugins:{{
        legend:{{display:false}},
        tooltip:{{...tt,borderColor:'rgba(56,189,248,.3)',titleColor:'#38bdf8',
          callbacks:{{label:c=>` ${{c.parsed.x}} chamados`}}}}
      }},
      scales:{{
        x:{{display:false}},
        y:{{grid:{{display:false}},border:{{display:false}},
          ticks:{{color:'#cbd5e1',font:{{weight:'600',size:11}},padding:6}}}}
      }}
    }},
    plugins:[{{
      afterDatasetsDraw(chart){{
        const c2=chart.ctx,meta=chart.getDatasetMeta(0);
        meta.data.forEach((bar,i)=>{{
          const v=SV[i];if(!v)return;
          c2.save();c2.fillStyle='rgba(125,211,252,.85)';
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
  const CORES=['#3b82f6','#fbbf24','#f97316','#ef4444','#dc2626'];
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
        tooltip:{{...tt,borderColor:'rgba(239,68,68,.3)',titleColor:'#f87171',
          callbacks:{{label:c=>` ${{c.parsed.y}} chamados`}}}}
      }},
      scales:{{
        x:{{grid:{{display:false}},border:{{display:false}},
          ticks:{{color:'#94a3b8',font:{{weight:'800',size:10}}}}}},
        y:{{grid:{{color:'rgba(255,255,255,.04)',drawTicks:false}},border:{{display:false}},
          ticks:{{color:'#3d5068',padding:4,stepSize:1}},
          min:0}}
      }}
    }},
    plugins:[{{
      afterDatasetsDraw(chart){{
        const c2=chart.ctx,meta=chart.getDatasetMeta(0);
        meta.data.forEach((bar,i)=>{{
          const v=PV[i];if(!v)return;
          c2.save();
          c2.fillStyle='#f8fafc';c2.font='bold 12px Geist Mono,monospace';
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
      data:TV,borderColor:'#38bdf8',borderWidth:2.5,
      backgroundColor:'rgba(56,189,248,0.12)',fill:true,tension:0.38,
      pointBackgroundColor:TV.map(v=>v===maxV?'#fbbf24':'#38bdf8'),
      pointBorderColor:TV.map(v=>v===maxV?'#92400e':'#080d18'),
      pointBorderWidth:TV.map(v=>v===maxV?2.5:1.5),
      pointRadius:TV.map(v=>v===maxV?6:3.5),
      pointHoverRadius:7,pointHoverBackgroundColor:'#7dd3fc',
    }}]}},
    options:{{
      responsive:true,maintainAspectRatio:false,
      animation:{{duration:300,easing:'easeOutCubic'}},
      plugins:{{
        legend:{{display:false}},
        tooltip:{{...tt,borderColor:'rgba(56,189,248,.3)',titleColor:'#38bdf8',
          callbacks:{{label:c=>` ${{c.parsed.y}} tickets`}}}}
      }},
      scales:{{
        x:{{grid:{{color:'rgba(255,255,255,.03)',drawTicks:false}},border:{{display:false}},
          ticks:{{color:'#7d8fa8',maxRotation:30,font:{{size:9.5}}}}}},
        y:{{grid:{{color:'rgba(255,255,255,.03)',drawTicks:false}},border:{{display:false}},
          ticks:{{color:'#3d5068',padding:6,stepSize:1}},min:0}}
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
      backgroundColor:'rgba(109,40,217,0.75)',
      hoverBackgroundColor:'#c4b5fd',
      borderRadius:{{topRight:5,bottomRight:5}},borderSkipped:false,barThickness:15,
    }}]}},
    options:{{
      indexAxis:'y',
      responsive:true,maintainAspectRatio:false,
      animation:{{duration:300,easing:'easeOutQuart'}},
      plugins:{{
        legend:{{display:false}},
        tooltip:{{...tt,borderColor:'rgba(167,139,250,.3)',titleColor:'#a78bfa',
          callbacks:{{label:c=>` ${{c.parsed.x}} chamados`}}}}
      }},
      scales:{{
        x:{{display:false}},
        y:{{grid:{{display:false}},border:{{display:false}},
          ticks:{{color:'#cbd5e1',font:{{weight:'600',size:11}},padding:6}}}}
      }}
    }},
    plugins:[{{
      afterDatasetsDraw(chart){{
        const c2=chart.ctx,meta=chart.getDatasetMeta(0);
        meta.data.forEach((bar,i)=>{{
          c2.save();c2.fillStyle='rgba(196,181,253,.9)';
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
      backgroundColor:'rgba(76,29,149,0.75)',
      hoverBackgroundColor:'#ddd6fe',
      borderRadius:{{topRight:5,bottomRight:5}},borderSkipped:false,barThickness:14,
    }}]}},
    options:{{
      indexAxis:'y',
      responsive:true,maintainAspectRatio:false,
      animation:{{duration:300,easing:'easeOutQuart'}},
      plugins:{{
        legend:{{display:false}},
        tooltip:{{...tt,borderColor:'rgba(167,139,250,.3)',titleColor:'#a78bfa',
          callbacks:{{label:c=>` ${{c.parsed.x}} pedidos`}}}}
      }},
      scales:{{
        x:{{display:false}},
        y:{{grid:{{display:false}},border:{{display:false}},
          ticks:{{color:'#cbd5e1',font:{{weight:'600',size:11}},padding:6}}}}
      }}
    }},
    plugins:[{{
      afterDatasetsDraw(chart){{
        const c2=chart.ctx,meta=chart.getDatasetMeta(0);
        meta.data.forEach((bar,i)=>{{
          c2.save();c2.fillStyle='rgba(221,214,254,.85)';
          c2.font='bold 11px Geist Mono,monospace';
          c2.textAlign='left';c2.textBaseline='middle';
          c2.fillText(CV[i],bar.x+7,bar.y);c2.restore();
        }});
      }}
    }}]
  }});
}})();
</script>
</body></html>"""

                components.html(DASH_HTML, height=1020, scrolling=False)

        # Execução das Métricas no Tab Correspondente
        with tab_analytics:
            renderizar_analytics_realtime()
    # ── 2. TAB GOVERNANÇA DE CHAMADOS ──
        with tab_chamados:
            if not df_base.empty:
                # ── Preparação das opções de filtro ──
                prioridades_disponiveis = sorted([
                    x for x in df_base['Prioridade'].dropna().unique()
                    if str(x).strip().lower() not in ['nan', 'none', '']
                ], key=lambda p: {'Baixa':0,'Média':1,'Alta':2,'Crítica':3,'Urgente':4}.get(p, 99))
                status_disponiveis = ["Aberto", "Em Atendimento", "Concluído"]
                setores_disponiveis = sorted([
                    x for x in df_base['Setor'].dropna().unique()
                    if str(x).strip().lower() not in ['nan', 'none', '']
                ])

                # ── Reset de filtros de chamados ──
                if st.session_state.get("_reset_filters_ch"):
                    for k in ["f_prio_ch", "f_status_ch", "f_setor_ch", "f_txt_mgr", "f_data_ini_ch", "f_data_fim_ch"]:
                        if k in st.session_state:
                            del st.session_state[k]
                    st.session_state["_reset_filters_ch"] = False
                    st.rerun(scope="app")

                # ── PAINEL DE FILTROS — SELECTBOX / MULTISELECT ──
                with st.container(border=True):
                    _fa, _fb, _fc, _fd, _fe, _ff = st.columns([1.4, 1.2, 1.8, 1, 1, 0.7])

                    with _fa:
                        filtro_prio = st.multiselect(
                            "⚡ Urgência",
                            options=prioridades_disponiveis,
                            default=prioridades_disponiveis,
                            key="f_prio_ch",
                            placeholder="Todas",
                        )
                        if not filtro_prio:
                            filtro_prio = prioridades_disponiveis

                    with _fb:
                        filtro_status = st.multiselect(
                            "◎ Status",
                            options=status_disponiveis,
                            default=status_disponiveis,
                            key="f_status_ch",
                            placeholder="Todos",
                        )
                        if not filtro_status:
                            filtro_status = status_disponiveis

                    with _fc:
                        filtro_setor = st.multiselect(
                            "◈ Setor / Unidade",
                            options=setores_disponiveis,
                            default=setores_disponiveis,
                            key="f_setor_ch",
                            placeholder="Todos",
                        )
                        if not filtro_setor:
                            filtro_setor = setores_disponiveis

                    with _fd:
                        filtro_data_ini_ch = st.date_input("📅 De", value=None, key="f_data_ini_ch", format="DD/MM/YYYY")

                    with _fe:
                        filtro_data_fim_ch = st.date_input("📅 Até", value=None, key="f_data_fim_ch", format="DD/MM/YYYY")

                    with _ff:
                        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
                        if st.button("↺", key="btn_reset_filters", use_container_width=True, help="Limpar filtros"):
                            st.session_state["_reset_filters_ch"] = True
                            st.rerun(scope="app")

                busca_texto = st.text_input(
                    "🔍 Busca livre",
                    key="f_txt_mgr",
                    placeholder="Nome, protocolo, setor ou palavra-chave…",
                    label_visibility="visible",
                )

                # Aplica a filtragem tratando maiúsculas/minúsculas e acentuação de forma segura
                # Registros sem Prioridade/Setor preenchidos (legados) sempre passam pelo filtro
                df_filtrado = df_base[
                    (df_base['Prioridade'].fillna('').isin(filtro_prio) | (df_base['Prioridade'].fillna('') == '')) &
                    df_base['Status'].fillna('').str.strip().isin(filtro_status) &
                    (df_base['Setor'].fillna('').isin(filtro_setor) | (df_base['Setor'].fillna('') == ''))
                ]

                # Filtro por data
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

                # Separação das listas baseada no que está ativo no filtro
                df_ativa = ordenar_por_prioridade_e_data(
                    df_filtrado[df_filtrado['Status'].str.strip().isin(["Aberto", "Em Atendimento"])],
                    "Prioridade"
                )
                df_historico = df_filtrado[df_filtrado['Status'].str.strip() == "Concluído"]

                c_space, c_export = st.columns([3, 1])
                with c_export:
                    st.download_button("📥 Exportar PDF filtrado", data=gerar_pdf(df_filtrado), file_name="chamados_ti_filtrado.pdf", use_container_width=True, key="dl_pdf_ch")

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

                st.markdown("#### ⚡ Fila de Trabalho Ativa")
                if not df_ativa.empty:
                    for idx, row in df_ativa.iterrows():
                        prio = str(row.get('Prioridade', '—'))
                        status_atual = str(row.get('Status', '—'))
                        status_html = (
                            '<span class="pill pill-red">🔴 Aberto</span>'
                            if status_atual == "Aberto"
                            else '<span class="pill pill-amber">🟡 Em Atendimento</span>'
                        )
                        desc_texto = str(row.get('Descrição', '')).strip()
                        desc_preview = desc_texto[:150] + '...' if len(desc_texto) > 150 else desc_texto
                        tempo_fila = tempo_decorrido(row.get("Data", ""))
                        anexo_lbl = "Sem anexo" if str(row.get('Anexo', 'Nenhum')) in ["Nenhum", "", "nan"] else "Anexo disponível"

                        with st.container(border=True):
                            st.markdown(f"""
                            <div class="queue-card {classe_fila(prio)}">
                                <div class="work-card-head">
                                    <div>
                                        <div class="work-card-id">Chamado {escapar_html(primeiro_valor(row, 'Protocolo', padrao=f'#{formatar_id(idx)}'))}</div>
                                        <div class="work-card-title">{escapar_html(row.get('Usuário'))}</div>
                                        <div class="work-card-subtitle">{escapar_html(row.get('Setor'))} • {escapar_html(extrair_categoria(row))} • {escapar_html(row.get('Data'))}</div>
                                    </div>
                                    <div class="work-badge-row">
                                        {status_html}
                                    </div>
                                </div>
                                <div class="work-card-desc">{html.escape(desc_preview)}</div>
                                <div class="work-card-meta">
                                    <span>⚡ Urgência <strong>{escapar_html(prio)}</strong></span>
                                    <span>⏱ Em fila <strong>{escapar_html(tempo_fila)}</strong></span>
                                    <span>📎 <strong>{escapar_html(anexo_lbl)}</strong></span>
                                </div>
                            </div>
                        """, unsafe_allow_html=True)

                            _sp_q, c_action = st.columns([3, 1])
                            with c_action:
                                if st.button("🔧 Despachar", key=f"btn_ch_{idx}", use_container_width=True, type="secondary"):
                                    modal_responder_chamado(idx, row)
                else:
                    st.markdown('<div class="empty-state"><strong>Fila de chamados limpa</strong><span>Nenhum chamado aberto ou em atendimento no filtro atual.</span></div>', unsafe_allow_html=True)

                st.write("")
            
                # ── ALTERAÇÃO: Histórico interativo com botão de edição ──
                with st.expander("📁 Histórico e Edição de Chamados Concluídos", expanded=False):
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
                     
                        for idx, row in df_historico.iterrows():
                            prio_hist = str(row.get('Prioridade', '—')).strip() or "—"
                            protocolo_hist = escapar_html(primeiro_valor(row, 'Protocolo', padrao=f'#{formatar_id(idx)}'))
                            usuario_hist = escapar_html(row.get('Usuário'))
                            setor_hist = escapar_html(row.get('Setor'))
                            categoria_hist = escapar_html(extrair_categoria(row))
                            data_hist = escapar_html(row.get('Data'))

                            with st.container(border=True):
                                h_col_info, h_col_action = st.columns([6.6, 1])

                                with h_col_info:
                                    st.markdown(f"""
                                    <div class="arc-row">
                                        <span class="arc-id">{protocolo_hist}</span>
                                        <span class="arc-user">{usuario_hist}</span>
                                        <span class="arc-meta">{setor_hist} • {categoria_hist} • {escapar_html(prio_hist)}</span>
                                        <span class="arc-date">{data_hist}</span>
                                    </div>
                                """, unsafe_allow_html=True)

                                with h_col_action:
                                    # Chama o mesmo modal usado para despachar chamados novos
                                    if st.button("✏️ Editar", key=f"btn_edit_hist_{idx}", use_container_width=True):
                                        modal_responder_chamado(idx, row)
                    else:
                        st.markdown('<div class="empty-state"><strong>Sem histórico concluído</strong><span>Nenhum chamado concluído dentro do filtro atual.</span></div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="empty-state"><strong>Nenhum ticket registrado</strong><span>A base de chamados ainda não recebeu solicitações.</span></div>', unsafe_allow_html=True)

        # ── 3. TAB AUDITORIA DE COMPRAS ──
        with tab_compras:
            if not df_compras.empty:
                # ── Preparação das opções de filtro de compras ──
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

                # ── Reset de filtros de compras ──
                if st.session_state.get("_reset_filters_cp"):
                    for k in ["f_urg_cp", "f_status_cp", "f_setor_cp", "f_cat_cp", "f_data_ini_cp", "f_data_fim_cp", "f_txt_cp"]:
                        if k in st.session_state:
                            del st.session_state[k]
                    st.session_state["_reset_filters_cp"] = False
                    st.rerun(scope="app")

                # ── PAINEL DE FILTROS — SELECTBOX / MULTISELECT ──
                with st.container(border=True):
                    _ca, _cb, _cc, _cd, _ce, _cf, _cg = st.columns([1.2, 1.4, 1.4, 1.4, 0.9, 0.9, 0.6])

                    with _ca:
                        filtro_urg_cp = st.multiselect(
                            "⚡ Urgência",
                            options=urgencias_disponiveis_cp,
                            default=urgencias_disponiveis_cp,
                            key="f_urg_cp",
                            placeholder="Todas",
                        )
                        if not filtro_urg_cp:
                            filtro_urg_cp = urgencias_disponiveis_cp

                    with _cb:
                        filtro_status_cp = st.multiselect(
                            "◎ Status",
                            options=status_disponiveis_cp,
                            default=status_disponiveis_cp,
                            key="f_status_cp",
                            placeholder="Todos",
                        )
                        if not filtro_status_cp:
                            filtro_status_cp = status_disponiveis_cp

                    with _cc:
                        filtro_setor_cp = st.multiselect(
                            "◈ Setor",
                            options=setores_disponiveis_cp,
                            default=setores_disponiveis_cp,
                            key="f_setor_cp",
                            placeholder="Todos",
                        )
                        if not filtro_setor_cp:
                            filtro_setor_cp = setores_disponiveis_cp

                    with _cd:
                        filtro_cat_cp = st.multiselect(
                            "▦ Categoria",
                            options=categorias_disponiveis_cp,
                            default=categorias_disponiveis_cp,
                            key="f_cat_cp",
                            placeholder="Todas",
                        )
                        if not filtro_cat_cp:
                            filtro_cat_cp = categorias_disponiveis_cp

                    with _ce:
                        filtro_data_ini_cp = st.date_input("📅 De", value=None, key="f_data_ini_cp", format="DD/MM/YYYY")

                    with _cf:
                        filtro_data_fim_cp = st.date_input("📅 Até", value=None, key="f_data_fim_cp", format="DD/MM/YYYY")

                    with _cg:
                        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
                        if st.button("↺", key="btn_reset_filters_cp", use_container_width=True, help="Limpar filtros"):
                            st.session_state["_reset_filters_cp"] = True
                            st.rerun(scope="app")

                busca_texto_cp = st.text_input(
                    "🔍 Busca livre",
                    key="f_txt_cp",
                    placeholder="Item, usuário, setor ou protocolo…",
                    label_visibility="visible",
                )

                # ── Aplicação dos filtros de compras ──
                # Registros sem Urgência preenchida (legados) sempre passam pelo filtro de urgência
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

                # Filtro por data
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
                    st.download_button("📥 Exportar PDF Compras", data=gerar_pdf_compras(df_compras_filtrado), file_name="compras_ti.pdf", use_container_width=True, key="dl_pdf_cp")

                # Mini-kanban de compras filtradas
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

                st.markdown("#### ⏳ Requisições de Compra Pendentes")
                if not df_comp_ativa.empty:
                    for idx_c, row_c in df_comp_ativa.iterrows():
                        urgencia_atual = str(row_c.get('Urgência', '—'))
                        status_compra = str(row_c.get('Status', '—'))
                        status_html = (
                            '<span class="pill pill-gray">⏳ Pendente</span>'
                            if status_compra == "Pendente"
                            else '<span class="pill pill-amber">🔍 Em Análise</span>'
                        )
                        justificativa = str(row_c.get('Justificativa', '')).strip()
                        just_preview = justificativa[:150] + '...' if len(justificativa) > 150 else justificativa

                        with st.container(border=True):
                            st.markdown(f"""
                            <div class="queue-card {classe_fila(urgencia_atual)}">
                                <div class="work-card-head">
                                    <div>
                                        <div class="work-card-id">Requisição {escapar_html(primeiro_valor(row_c, 'Protocolo', padrao=f'#{formatar_id(idx_c)}'))}</div>
                                        <div class="work-card-title">{escapar_html(row_c.get('Item'))}</div>
                                        <div class="work-card-subtitle">{escapar_html(row_c.get('Usuário'))} • {escapar_html(row_c.get('Setor'))} • {escapar_html(row_c.get('Data'))}</div>
                                    </div>
                                    <div class="work-badge-row">
                                        {status_html}
                                    </div>
                                </div>
                                <div class="work-card-desc">{html.escape(just_preview)}</div>
                                <div class="work-card-meta">
                                    <span>⚡ Urgência <strong>{escapar_html(urgencia_atual)}</strong></span>
                                    <span>📦 Qtd <strong>{escapar_html(row_c.get("Quantidade"))}</strong></span>
                                    <span>🏷️ <strong>{escapar_html(row_c.get("Categoria"))}</strong></span>
                                </div>
                            </div>
                        """, unsafe_allow_html=True)

                            _sp_c, cc_action = st.columns([3, 1])
                            with cc_action:
                                if st.button("🔍 Analisar", key=f"btn_cp_{idx_c}", use_container_width=True, type="secondary"):
                                    modal_responder_compra(idx_c, row_c)
                else:
                    st.markdown('<div class="empty-state"><strong>Compras auditadas</strong><span>Não há solicitações pendentes ou em análise no momento.</span></div>', unsafe_allow_html=True)

                st.write("")

                with st.expander("📁 Ver Arquivo de Solicitações Finalizadas", expanded=False):
                    if not df_comp_historico.empty:
                        st.markdown(f"""
                        <div class="history-panel-head">
                            <div class="history-panel-title">
                                <span class="history-title-dot" style="background:#38bdf8;box-shadow:0 0 18px rgba(56,189,248,.75);"></span>
                                <span>Arquivo de Compras Finalizadas</span>
                            </div>
                            <span class="history-count">{len(df_comp_historico)} registros</span>
                        </div>
                    """, unsafe_allow_html=True)

                        for idx_c, row_c in df_comp_historico.iterrows():
                            urg_hist  = str(row_c.get('Urgência', '—')).strip() or "—"
                            status_c  = str(row_c.get('Status',   '—')).strip() or "—"

                            if status_c == "Aprovado":
                                status_badge = '<span class="pill pill-blue">✅ Aprovado</span>'
                            elif status_c == "Concluído":
                                status_badge = '<span class="pill pill-green">🟢 Concluído</span>'
                            elif status_c == "Recusado":
                                status_badge = '<span class="pill pill-red">❌ Recusado</span>'
                            else:
                                status_badge = f'<span class="pill pill-gray">{escapar_html(status_c)}</span>'

                            protocolo_c  = escapar_html(primeiro_valor(row_c, 'Protocolo', padrao=f'#{formatar_id(idx_c)}'))
                            item_c       = escapar_html(row_c.get('Item'))
                            usuario_c    = escapar_html(row_c.get('Usuário'))
                            setor_c      = escapar_html(row_c.get('Setor'))
                            categoria_c  = escapar_html(row_c.get('Categoria'))
                            data_c       = escapar_html(row_c.get('Data'))
                            qtd_c        = escapar_html(row_c.get('Quantidade'))

                            card_html = (
                                '<div class="arc-row">'
                                f'<span class="arc-id">{protocolo_c}</span>'
                                f'<span class="arc-user">{item_c}</span>'
                                f'<span class="arc-meta">{usuario_c} • {setor_c} • {categoria_c} • Qtd: {qtd_c}</span>'
                                f'{status_badge}'
                                f'<span class="arc-date">{data_c}</span>'
                                '</div>'
                            )

                            with st.container(border=True):
                                h_col_info, h_col_action = st.columns([6.6, 1])

                                with h_col_info:
                                    st.markdown(card_html, unsafe_allow_html=True)

                                with h_col_action:
                                    if st.button("✏️ Rever", key=f"btn_edit_cp_{idx_c}", use_container_width=True):
                                        modal_responder_compra(idx_c, row_c)
                    else:
                        st.markdown('<div class="empty-state"><strong>Sem arquivo finalizado</strong><span>Nenhuma requisição aprovada, recusada ou concluída até o momento.</span></div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="empty-state"><strong>Nenhuma requisição registrada</strong><span>A base de compras ainda não recebeu solicitações.</span></div>', unsafe_allow_html=True)

                # ═══════════════════════════════════════════════════════
    # 5. ABA: RELATÓRIOS & EXPORTAÇÃO
    # ═══════════════════════════════════════════════════════
    _render_gestor()
elif pagina == "relatorios":

    # Página isolada em fragment: interações reroda só este trecho, não a app inteira.
    @st.fragment
    def _render_relatorios():
        import io, zipfile, os
        from datetime import datetime, date
        import calendar
        import pandas as pd

        # ── CONSTANTES DE SEGURANÇA ────────────────────────────────────────
        _MAX_TENT_R   = 5
        _LOCKOUT_R    = 300
        _SESSION_TTL_R = 3600

        # ── ESTADO INICIAL ─────────────────────────────────────────────────
        if "autenticado_relatorios"    not in st.session_state:
            st.session_state.autenticado_relatorios  = False
        if "relat_auth_ts"             not in st.session_state:
            st.session_state.relat_auth_ts           = 0.0
        if "relat_tentativas"          not in st.session_state:
            st.session_state.relat_tentativas        = 0
        if "relat_lockout_ate"         not in st.session_state:
            st.session_state.relat_lockout_ate       = 0.0

        # ── EXPIRAÇÃO AUTOMÁTICA ───────────────────────────────────────────
        _ts_r = time.time()
        if st.session_state.autenticado_relatorios:
            if (_ts_r - st.session_state.relat_auth_ts) > _SESSION_TTL_R:
                st.session_state.autenticado_relatorios = False
                st.session_state.relat_auth_ts          = 0.0

        # ── TELA DE AUTENTICAÇÃO CENTRALIZADA (LOCKSCREEN) ─────────────────
        if not st.session_state.autenticado_relatorios:
            _em_lock_r   = _ts_r < st.session_state.relat_lockout_ate
            _seg_rest_r  = max(0, int(st.session_state.relat_lockout_ate - _ts_r))
            _tent_rest_r = max(0, _MAX_TENT_R - st.session_state.relat_tentativas)

            # ── CSS da lockscreen: estreito e contido ──
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
            background: rgba(3,7,16,.97);
            border: 1px solid rgba(167,139,250,.22);
            border-radius: 12px 12px 0 0;
            font-family: 'Geist Mono', ui-monospace, monospace;
            box-shadow: 0 0 50px rgba(167,139,250,.10), 0 28px 70px rgba(0,0,0,.70);
        }
        .tr-card.locked {
            border-color: rgba(248,113,113,.35) !important;
        }
        .tr-card::before {
            content:''; position:absolute; top:0; left:0; right:0; height:1px;
            background: linear-gradient(90deg,transparent,rgba(167,139,250,.7),transparent);
            background-size: 200% 1px; animation: _scan_r 3s linear infinite;
        }
        .tr-card.locked::before {
            background: linear-gradient(90deg,transparent,rgba(248,113,113,.8),transparent);
            background-size: 200% 1px;
        }
        .tr-bar {
            display:flex; align-items:center; gap:7px; padding:9px 16px;
            background:rgba(167,139,250,.04); border-bottom:1px solid rgba(167,139,250,.08);
        }
        .tr-bar.locked { background:rgba(248,113,113,.04); border-bottom-color:rgba(248,113,113,.08); }
        .tr-dot { width:8px; height:8px; border-radius:50%; }
        .tr-bar-title {
            flex:1; text-align:center;
            font-size:.50rem; letter-spacing:.22em; text-transform:uppercase;
            color:rgba(167,139,250,.30); margin:0;
        }
        .tr-bar.locked .tr-bar-title { color:rgba(248,113,113,.30); }
        .tr-body { padding:22px 24px 18px; }
        .tr-icon-ring {
            width:52px; height:52px; border-radius:13px;
            margin:0 auto 14px;
            display:flex; align-items:center; justify-content:center;
            background:linear-gradient(135deg,rgba(167,139,250,.15),rgba(56,189,248,.08));
            border:1px solid rgba(167,139,250,.22);
            font-size:1.5rem;
        }
        .tr-icon-ring.locked {
            background:linear-gradient(135deg,rgba(248,113,113,.15),rgba(251,191,36,.06));
            border-color:rgba(248,113,113,.30);
        }
        .tr-badge {
            display:inline-flex; align-items:center; gap:6px;
            font-size:.50rem; letter-spacing:.16em; text-transform:uppercase;
            color:rgba(167,139,250,.50); margin-bottom:8px;
        }
        .tr-badge.locked { color:rgba(248,113,113,.55); }
        .tr-badge-dot { width:5px; height:5px; border-radius:50%; background:#a78bfa; flex-shrink:0; animation:_blink_r 1.5s ease-in-out infinite; }
        .tr-badge-dot.locked { background:#f87171; }
        .tr-title { font-size:.96rem; font-weight:800; color:#dce8f5; letter-spacing:-.01em; margin-bottom:4px; text-align:center; }
        .tr-sub   { font-size:.62rem; color:rgba(148,163,184,.42); text-align:center; }
        .tr-sep   { height:1px; margin:16px 0 14px; background:linear-gradient(90deg,transparent,rgba(167,139,250,.20),transparent); }
        .tr-sep.locked { background:linear-gradient(90deg,transparent,rgba(248,113,113,.18),transparent); }
        .tr-attempts { display:flex; align-items:center; gap:6px; margin-bottom:10px; justify-content:center; }
        .tr-att-dot { width:8px; height:8px; border-radius:50%; }
        .tr-att-dot.used  { background:#f87171; box-shadow:0 0 8px rgba(248,113,113,.6); }
        .tr-att-dot.avail { background:rgba(167,139,250,.25); border:1px solid rgba(167,139,250,.18); }
        .tr-lockout-bar {
            padding:11px 13px; border:1px solid rgba(248,113,113,.28); border-radius:8px;
            background:rgba(248,113,113,.06);
        }
        .tr-lockout-title { font-size:.62rem; font-weight:800; color:#f87171; letter-spacing:.08em; text-transform:uppercase; margin-bottom:3px; }
        .tr-lockout-msg   { font-size:.60rem; color:rgba(248,113,113,.65); }
        .tr-warn-token {
            padding:9px 12px; margin-bottom:10px;
            border:1px solid rgba(251,191,36,.24); border-radius:7px;
            background:rgba(251,191,36,.05);
            font-size:.60rem; color:rgba(251,191,36,.75); line-height:1.45;
        }
        .tr-warn-token strong { color:#fbbf24; }

        /* Form Streamlit — integrado embaixo do card */
        div[data-testid="stForm"][data-form-key="flr"] {
            background:rgba(3,7,16,.97) !important;
            border:1px solid rgba(167,139,250,.22) !important;
            border-top:none !important;
            border-radius:0 0 12px 12px !important;
            padding:4px 24px 18px !important;
            margin:0 !important;
            box-shadow:0 28px 70px rgba(0,0,0,.70) !important;
        }
        div[data-testid="stForm"][data-form-key="flr"] input[type="password"] {
            background:rgba(167,139,250,.04) !important;
            border:1px solid rgba(167,139,250,.18) !important;
            border-radius:7px !important; color:#c4b5f4 !important;
            font-family:'Geist Mono',monospace !important;
            font-size:.78rem !important; letter-spacing:.06em !important;
            height:40px !important; padding:0 14px !important;
        }
        div[data-testid="stForm"][data-form-key="flr"] input[type="password"]:focus {
            border-color:rgba(167,139,250,.55) !important;
            box-shadow:0 0 0 3px rgba(167,139,250,.10) !important;
        }
        div[data-testid="stForm"][data-form-key="flr"] .stFormSubmitButton button {
            width:100% !important; height:34px !important;
            background:transparent !important;
            border:1px solid rgba(167,139,250,.18) !important;
            border-radius:6px !important; color:rgba(167,139,250,.55) !important;
            font-family:'Geist Mono',monospace !important;
            font-size:.56rem !important; font-weight:700 !important;
            letter-spacing:.18em !important; text-transform:uppercase !important;
            padding:0 !important; margin-top:10px !important;
            transition:all .18s !important;
        }
        div[data-testid="stForm"][data-form-key="flr"] .stFormSubmitButton button:hover {
            background:rgba(167,139,250,.08) !important;
            border-color:rgba(167,139,250,.42) !important;
            color:#a78bfa !important;
            box-shadow:0 0 14px rgba(167,139,250,.14) !important;
        }
        .tr-footer-r {
            display:flex; justify-content:space-between; align-items:center;
            padding:10px 0 0;
            font-family:'Geist Mono',monospace; font-size:.46rem; letter-spacing:.10em;
        }
        </style>
        """, unsafe_allow_html=True)

            # ── Monta conteúdo do card ──
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
              <div class="tr-lockout-title">🔒 Acesso Bloqueado</div>
              <div class="tr-lockout-msg">Aguarde <strong style="color:#f87171;">{_seg_rest_r}s</strong> para tentar novamente.</div>
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

            # Colunas para centralizar tudo
            _ca, _cb, _cc = st.columns([1, 2.4, 1])
            with _cb:
                st.markdown(f"""
            <div class="tr-card {_locked_r}">
              <div class="tr-bar {_locked_r}">
                <div class="tr-dot" style="background:#f87171;opacity:.65;"></div>
                <div class="tr-dot" style="background:#fbbf24;opacity:.65;"></div>
                <div class="tr-dot" style="background:#34d399;opacity:.65;"></div>
                <span class="tr-bar-title">central · relatórios · exportação</span>
              </div>
              <div class="tr-body">
                <div class="tr-icon-ring {_locked_r}">{"🔒" if _em_lock_r else "📊"}</div>
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
                        _btn_r = st.form_submit_button("→  autenticar", use_container_width=True)
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
                                    st.markdown(f"""<div style="font-family:'Geist Mono',monospace;font-size:.58rem;
                                    color:rgba(248,113,113,.9);text-align:center;padding:4px 0 0;">
                                    🔒 &nbsp;Bloqueado por {_LOCKOUT_R//60} min.</div>""", unsafe_allow_html=True)
                                else:
                                    _rest_r = _MAX_TENT_R - st.session_state.relat_tentativas
                                    st.markdown(f"""<div style="font-family:'Geist Mono',monospace;font-size:.58rem;
                                    color:rgba(248,113,113,.85);text-align:center;padding:4px 0 0;">
                                    ✗ &nbsp;Token inválido — {_rest_r} tentativa(s) restante(s)</div>""", unsafe_allow_html=True)

                st.markdown(f"""
            <div class="tr-footer-r">
              <span style="color:rgba(167,139,250,.18);">IT ServiceDesk Pro · v2.1</span>
              <span style="color:rgba(52,211,153,.35);display:flex;align-items:center;gap:5px;">
                <span style="width:4px;height:4px;border-radius:50%;background:#34d399;
                  box-shadow:0 0 4px #34d399;display:inline-block;"></span>
                hmac · sha256
              </span>
            </div>
            """, unsafe_allow_html=True)

            st.stop()


        # ── CONTEÚDO LIBERADO (SÓ APARECE SE AUTENTICADO) ───────────────────
        _sessao_rest_r  = max(0, int(_SESSION_TTL_R - (_ts_r - st.session_state.relat_auth_ts)))
        _sessao_min_r   = _sessao_rest_r // 60
        _agora_str_r    = datetime.now().strftime("%d/%m/%Y  %H:%M")

        # ── CSS EXCLUSIVO DA PÁGINA RELATÓRIOS ──
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
        background: linear-gradient(135deg, rgba(14,20,42,.98), rgba(8,12,32,.95));
        border: 1px solid rgba(167,139,250,.18);
        border-radius: 14px;
        overflow: hidden;
        box-shadow: 0 2px 40px rgba(0,0,0,.45), 0 0 0 1px rgba(167,139,250,.05);
    }
    .rel-hero::before {
        content: "";
        position: absolute;
        top: 0; left: 0; right: 0; height: 2px;
        background: linear-gradient(90deg, #a78bfa, #38bdf8, #a78bfa);
        background-size: 200% 100%;
        animation: relHeroShimmer 4s linear infinite;
    }
    @keyframes relHeroShimmer { 0%{background-position:0} 100%{background-position:200%} }
    .rel-hero::after {
        content: "";
        position: absolute;
        right: 0; top: 0; bottom: 0;
        width: 35%;
        background: radial-gradient(ellipse at right center, rgba(167,139,250,.06), transparent 70%);
        pointer-events: none;
    }
    .rel-hero-left { position: relative; z-index: 1; }
    .rel-hero-eyebrow {
        font-family: 'Geist Mono', monospace;
        font-size: .58rem; font-weight: 800;
        letter-spacing: .18em; text-transform: uppercase;
        color: rgba(167,139,250,.60);
        display: flex; align-items: center; gap: 8px;
        margin-bottom: 6px;
    }
    .rel-hero-eyebrow::before {
        content: "";
        width: 6px; height: 6px; border-radius: 50%;
        background: #a78bfa;
        box-shadow: 0 0 10px rgba(167,139,250,.8);
    }
    .rel-hero-title {
        font-size: 1.32rem; font-weight: 800;
        color: #f0f6ff; letter-spacing: -.01em;
        margin: 0 0 4px;
        text-shadow: 0 0 30px rgba(167,139,250,.14);
    }
    .rel-hero-desc {
        font-size: .76rem; color: rgba(148,163,184,.62);
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
        background: rgba(167,139,250,.06);
        border: 1px solid rgba(167,139,250,.22);
        border-radius: 999px;
        font-family: 'Geist Mono', monospace;
        font-size: .58rem; font-weight: 700;
        letter-spacing: .08em; color: rgba(167,139,250,.75);
    }
    .rel-session-badge::before {
        content: "";
        width: 5px; height: 5px; border-radius: 50%;
        background: #a78bfa;
        box-shadow: 0 0 6px rgba(167,139,250,.9);
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
        background: linear-gradient(145deg, rgba(12,20,36,.96), rgba(6,12,24,.90));
        border: 1px solid rgba(148,163,184,.12);
        box-shadow: 0 8px 24px rgba(0,0,0,.35), inset 0 1px 0 rgba(255,255,255,.02);
        transition: transform .2s, border-color .2s, box-shadow .2s;
    }
    .kpi-v2:hover {
        transform: translateY(-2px);
        border-color: var(--kc, rgba(56,189,248,.40));
        box-shadow: 0 0 24px color-mix(in srgb, var(--kc, #38bdf8) 12%, transparent),
                    0 18px 44px rgba(0,0,0,.45);
    }
    .kpi-v2::before {
        content: "";
        position: absolute;
        left: 0; top: 0; bottom: 0; width: 3px;
        background: var(--kc, #38bdf8);
        box-shadow: 0 0 18px var(--kc, #38bdf8);
    }
    .kpi-v2::after {
        content: "";
        position: absolute;
        inset: 0;
        background: linear-gradient(135deg, rgba(255,255,255,.025), transparent 50%);
        pointer-events: none;
    }
    .kpi-v2-icon { font-size: 1.1rem; margin-bottom: 8px; display: block; }
    .kpi-v2-val {
        font-family: 'Geist Mono', monospace;
        font-size: 2rem; font-weight: 800; line-height: 1;
        color: var(--kc, #f0f6ff);
        margin-bottom: 5px;
        text-shadow: 0 0 18px color-mix(in srgb, var(--kc, #38bdf8) 25%, transparent);
    }
    .kpi-v2-label {
        font-size: .62rem; text-transform: uppercase;
        letter-spacing: .10em; color: rgba(138,164,194,.65); font-weight: 700;
    }
    /* ── EXEC RECOMENDATION BANNER ── */
    .exec-banner {
        display: flex; align-items: flex-start; gap: 14px;
        padding: 16px 20px;
        background: linear-gradient(135deg, rgba(167,139,250,.07), rgba(56,189,248,.04));
        border: 1px solid rgba(167,139,250,.16);
        border-radius: 10px;
        margin-bottom: 22px;
    }
    .exec-banner-icon { font-size: 1.2rem; flex-shrink: 0; margin-top: 1px; }
    .exec-banner-content { flex: 1; }
    .exec-banner-title {
        font-size: .62rem; font-weight: 800;
        font-family: 'Geist Mono', monospace;
        letter-spacing: .12em; text-transform: uppercase;
        color: rgba(167,139,250,.70); margin-bottom: 4px;
    }
    .exec-banner-text { font-size: .82rem; color: #b7c7dd; line-height: 1.5; }
    </style>
    """, unsafe_allow_html=True)

        # ── HERO BAR RELATÓRIOS ──
        st.markdown(f"""
    <div class="rel-hero animate-fade-in">
        <div class="rel-hero-left">
            <div class="rel-hero-eyebrow">Data Operations · Exportação & Auditoria</div>
            <div class="rel-hero-title">Relatórios & Analytics</div>
            <div class="rel-hero-desc">Visualize indicadores consolidados, exporte dados e monitore o desempenho operacional de TI.</div>
        </div>
        <div class="rel-hero-right">
            <div class="rel-session-badge">sessão segura · {_sessao_min_r}min</div>
            <div style="font-family:'Geist Mono',monospace;font-size:.60rem;color:rgba(148,163,184,.38);">🕐 {_agora_str_r}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

        _logout_col = st.columns([6, 1])[1]
        with _logout_col:
            if st.button("🔓 Sair", key="relatorios_logout", use_container_width=True, help="Encerrar sessão de relatórios"):
                st.session_state.autenticado_relatorios = False
                st.session_state.relat_auth_ts          = 0.0
                st.rerun(scope="app")

        df_rel   = carregar_dados_cacheados(DATABASE, COLUNAS)
        df_comp  = carregar_dados_cacheados(DATABASE_COMPRAS, COLUNAS_COMPRAS)
        # Cópias íntegras (pré-filtro) — usadas apenas no fechamento mensal automático,
        # que deve sempre refletir o dado oficial completo, independente do filtro em tela.
        df_rel_raw  = df_rel.copy()
        df_comp_raw = df_comp.copy()
        agora    = datetime.now()
        hoje     = agora.date()
        ult_dia  = calendar.monthrange(hoje.year, hoje.month)[1]
        is_last  = (hoje.day == ult_dia)
        is_hora  = (agora.hour == 18)

        # ── FILTROS DO RELATÓRIO ─────────────────────────────────────────────
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

        with st.container(border=True):
            ff1, ff2, ff3, ff4, ff5 = st.columns([1.6, 1.3, 1.3, 1.3, 0.7])
            with ff1:
                _periodo_sel = st.date_input(
                    "Período (Data)", value=(_min_dt_rel, _max_dt_rel),
                    min_value=_min_dt_rel, max_value=_max_dt_rel,
                    key="rel_filtro_periodo",
                )
            with ff2:
                _f_setor = st.multiselect("Setor", options=_setores_opts, key="rel_filtro_setor")
            with ff3:
                _f_status = st.multiselect("Status (Chamados)", options=_status_ch_opts, key="rel_filtro_status")
            with ff4:
                _f_prio = st.multiselect("Prioridade", options=_prio_opts, key="rel_filtro_prioridade")
            with ff5:
                st.markdown("<div style='height:1.7rem'></div>", unsafe_allow_html=True)
                if st.button("🧹 Limpar", use_container_width=True, key="rel_filtro_limpar"):
                    for _k in ("rel_filtro_periodo", "rel_filtro_setor", "rel_filtro_status",
                               "rel_filtro_prioridade", "rel_filtro_status_compra"):
                        st.session_state.pop(_k, None)
                    st.rerun(scope="fragment")
            _f_status_compra = st.multiselect("Status (Compras)", options=_status_comp_opts, key="rel_filtro_status_compra")

        # Normaliza o range de datas (o widget pode devolver 1 data só enquanto o usuário
        # ainda está selecionando o segundo extremo do período)
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
            st.caption(f"🔎 Filtros ativos — exibindo {len(df_rel)} de {len(df_rel_raw)} chamado(s) e {len(df_comp)} de {len(df_comp_raw)} compra(s).")

        # ── KPI SNAPSHOT — PROCESSAMENTO DE DADOS LÓGICOS ───────────────────
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

        st.markdown("""<div class="sec-label animate-fade-in"><div class="sec-dot"></div>SNAPSHOT DE KPIs — PERÍODO COMPLETO</div>""", unsafe_allow_html=True)

        # Grid de KPIs V2 — design premium, anti-quebra
        st.markdown(f"""
        <div class="kpi-grid-v2 animate-fade-in">
            <div class="kpi-v2" style="--kc:#38bdf8">
                <span class="kpi-v2-icon">🎫</span>
                <div class="kpi-v2-val">{total_ch}</div>
                <div class="kpi-v2-label">Total de Chamados</div>
            </div>
            <div class="kpi-v2" style="--kc:#34d399">
                <span class="kpi-v2-icon">✅</span>
                <div class="kpi-v2-val">{taxa}</div>
                <div class="kpi-v2-label">Taxa de Resolução</div>
            </div>
            <div class="kpi-v2" style="--kc:#f87171">
                <span class="kpi-v2-icon">🔴</span>
                <div class="kpi-v2-val">{urgentes}</div>
                <div class="kpi-v2-label">Alta / Crítica</div>
            </div>
            <div class="kpi-v2" style="--kc:#a78bfa">
                <span class="kpi-v2-icon">🛒</span>
                <div class="kpi-v2-val">{total_comp}</div>
                <div class="kpi-v2-label">Requisições de Compra</div>
            </div>
            <div class="kpi-v2" style="--kc:#fbbf24">
                <span class="kpi-v2-icon">⏳</span>
                <div class="kpi-v2-val">{abertos}</div>
                <div class="kpi-v2-label">Chamados Abertos</div>
            </div>
            <div class="kpi-v2" style="--kc:#f97316">
                <span class="kpi-v2-icon">⚙️</span>
                <div class="kpi-v2-val">{atendimento}</div>
                <div class="kpi-v2-label">Em Atendimento</div>
            </div>
            <div class="kpi-v2" style="--kc:#34d399">
                <span class="kpi-v2-icon">📦</span>
                <div class="kpi-v2-val">{comp_aprov}</div>
                <div class="kpi-v2-label">Compras Aprovadas</div>
            </div>
            <div class="kpi-v2" style="--kc:#f87171">
                <span class="kpi-v2-icon">⏰</span>
                <div class="kpi-v2-val">{comp_pend}</div>
                <div class="kpi-v2-label">Compras Pendentes</div>
            </div>
        </div>
    """, unsafe_allow_html=True)

        # ── BANNER EXECUTIVO ──
        st.markdown(f"""
        <div class="exec-banner animate-fade-in">
            <span class="exec-banner-icon">📋</span>
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

        # ── INFO ADICIONAL PARAMETRIZADA ─────────────────────────────────────
        c_inf1, c_inf2, c_inf3, c_inf4 = st.columns(4)
        _info_cards = [
            (c_inf1, "🏆", "Setor Mais Demandante", top_setor, "#38bdf8"),
            (c_inf2, "⚠️", "Prioridade Mais Frequente", top_prio, "#fbbf24"),
            (c_inf3, "🔴", "Chamados Críticos (ativos)", str(crit_ativos), "#f87171"),
            (c_inf4, "✅", "Concluídos no Período", str(concluidos), "#34d399"),
        ]
        for col, ico, lbl, val, color in _info_cards:
            with col:
                with st.container(border=True):
                    st.markdown(f"""
                    <div style="padding:2px 0;">
                        <div style="font-size:.60rem;color:rgba(148,163,184,.55);
                            text-transform:uppercase;letter-spacing:.10em;
                            font-weight:800;margin-bottom:6px;">{ico} {lbl}</div>
                        <div style="font-size:1.55rem;font-weight:800;color:{color};
                            font-family:'Geist Mono',monospace;line-height:1;">{val}</div>
                    </div>
                """, unsafe_allow_html=True)

        # ── PANORAMA VISUAL — GRÁFICOS ───────────────────────────────────────
        st.markdown("---")
        st.markdown("""<div class="sec-label animate-fade-in"><div class="sec-dot"></div>PANORAMA VISUAL — INDICADORES GRÁFICOS</div>""", unsafe_allow_html=True)

        import plotly.graph_objects as go  # lazy: só carrega ao abrir relatórios

        # Paleta alinhada ao tema do portal
        _C_ACCENT, _C_VIOLET, _C_GREEN = "#38bdf8", "#a78bfa", "#34d399"
        _C_AMBER, _C_RED, _C_ORANGE = "#fbbf24", "#f87171", "#f97316"
        _STATUS_COLORS = {
            "Aberto": _C_AMBER, "Em Atendimento": _C_ACCENT,
            "Concluído": _C_GREEN, "Cancelado": "#64748b", "Outros": "#64748b",
        }
        _PRIO_COLORS = {
            "Baixa": "#64748b", "Média": _C_ACCENT, "Alta": _C_AMBER,
            "Crítica": _C_ORANGE, "Urgente": _C_RED,
        }

        def _aplicar_tema(fig, altura=300):
            fig.update_layout(
                template="plotly_dark",
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font=dict(family="Geist, sans-serif", color="#e2e8f0", size=12),
                margin=dict(l=10, r=10, t=14, b=10),
                height=altura,
                legend=dict(orientation="h", yanchor="bottom", y=-0.22, x=0, font=dict(size=11)),
                hoverlabel=dict(bgcolor="#0c1828", bordercolor="rgba(56,189,248,.4)", font_size=12),
            )
            fig.update_xaxes(showgrid=True, gridcolor="rgba(148,163,184,.08)", zeroline=False, color="#7d8fa8")
            fig.update_yaxes(showgrid=True, gridcolor="rgba(148,163,184,.08)", zeroline=False, color="#7d8fa8")
            return fig

        _PLOT_CFG = {"displayModeBar": False, "responsive": True}

        if df_rel.empty and df_comp.empty:
            st.info("Sem dados suficientes para gerar os gráficos analíticos.")
        else:
            g1c1, g1c2 = st.columns(2)
            # Donut — distribuição de status dos chamados
            with g1c1:
                with st.container(border=True):
                    st.markdown("###### 🎫 Distribuição de chamados por status")
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
                            marker=dict(colors=[_STATUS_COLORS.get(x, "#64748b") for x in sc.index],
                                        line=dict(color="#02040b", width=2)),
                            textinfo="percent", textposition="inside",
                            insidetextorientation="horizontal",
                            textfont=dict(size=12, color="#02040b"),
                        ))
                        fig.update_layout(annotations=[dict(
                            text=f"<b>{total_ch}</b><br><span style='font-size:11px;color:#7d8fa8'>chamados</span>",
                            x=0.5, y=0.5, xref="paper", yref="paper",
                            showarrow=False, font=dict(size=22, color="#e2e8f0"))])
                        st.plotly_chart(_aplicar_tema(fig, 300), use_container_width=True, config=_PLOT_CFG)
                    else:
                        st.caption("Sem chamados registrados.")
            # Barras horizontais — chamados por setor
            with g1c2:
                with st.container(border=True):
                    st.markdown("###### 🏢 Chamados por setor")
                    if not df_rel.empty:
                        ss = df_rel["Setor"].value_counts().head(8).sort_values()
                        fig = go.Figure(go.Bar(
                            x=ss.values.tolist(), y=ss.index.tolist(), orientation="h",
                            marker=dict(color=ss.values.tolist(),
                                        colorscale=[[0, "#1e3a5f"], [1, _C_ACCENT]],
                                        line=dict(width=0)),
                            text=ss.values.tolist(), textposition="outside",
                            textfont=dict(color="#e2e8f0", size=11),
                        ))
                        st.plotly_chart(_aplicar_tema(fig, 300), use_container_width=True, config=_PLOT_CFG)
                    else:
                        st.caption("Sem chamados registrados.")

            # Barras — prioridade (criticidade) — largura cheia
            with st.container(border=True):
                st.markdown("###### 🚨 Chamados por nível de criticidade")
                if not df_rel.empty:
                    ordem = ["Baixa", "Média", "Alta", "Crítica"]
                    pc = df_rel["Prioridade"].value_counts()
                    cats = [p for p in ordem if p in pc.index] + [p for p in pc.index if p not in ordem]
                    vals = [int(pc.get(p, 0)) for p in cats]
                    fig = go.Figure(go.Bar(
                        x=cats, y=vals,
                        marker=dict(color=[_PRIO_COLORS.get(p, _C_ACCENT) for p in cats], line=dict(width=0)),
                        text=vals, textposition="outside", textfont=dict(color="#e2e8f0", size=11),
                    ))
                    st.plotly_chart(_aplicar_tema(fig, 300), use_container_width=True, config=_PLOT_CFG)
                else:
                    st.caption("Sem chamados registrados.")

            # Tendência temporal — chamados ao longo do tempo
            with st.container(border=True):
                st.markdown("###### 📈 Evolução de chamados no tempo")
                if not df_rel.empty:
                    _dft = df_rel.copy()
                    _dft["__dt"] = pd.to_datetime(_dft["Data"], format="%d/%m/%Y %H:%M", errors="coerce")
                    _dft = _dft.dropna(subset=["__dt"])
                    if not _dft.empty:
                        serie = _dft.groupby(_dft["__dt"].dt.date).size().sort_index()
                        fig = go.Figure(go.Scatter(
                            x=list(serie.index), y=serie.values.tolist(),
                            mode="lines", line=dict(color=_C_ACCENT, width=2.5, shape="spline"),
                            fill="tozeroy", fillcolor="rgba(56,189,248,.12)",
                            hovertemplate="%{x}<br><b>%{y}</b> chamado(s)<extra></extra>",
                        ))
                        st.plotly_chart(_aplicar_tema(fig, 280), use_container_width=True, config=_PLOT_CFG)
                    else:
                        st.caption("Datas insuficientes para a série temporal.")
                else:
                    st.caption("Sem chamados registrados.")

        # ── FUNÇÕES INTERNAS DE EXTRAÇÃO DE LOGS ────────────────────────────
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
                # Cabeçalho
                ws.append(list(df.columns))
                for cell in ws[1]:
                    cell.font      = HDR_FONT
                    cell.fill      = HDR_FILL
                    cell.alignment = HDR_ALIGN
                    cell.border    = BORDER
                ws.row_dimensions[1].height = 28
                # Dados
                for r_idx, row in enumerate(df.itertuples(index=False), start=2):
                    ws.append(list(row))
                    fill = ROW_FILL_E if r_idx % 2 == 0 else ROW_FILL_O
                    for cell in ws[r_idx]:
                        cell.fill      = fill
                        cell.alignment = CELL_ALIGN
                        cell.border    = BORDER
                        cell.font      = Font(name="Arial", size=9)
                # Largura automática
                for col_idx, col_cells in enumerate(ws.columns, start=1):
                    max_len = max((len(str(c.value or "")) for c in col_cells), default=8)
                    ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 52)
                # Congelar primeira linha
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

        # ── DOWNLOAD MANUAL ─────────────────────────────────────────
        st.markdown("""<div class="sec-label animate-fade-in"><div class="sec-dot"></div>EXPORTAÇÃO DE DADOS SOB DEMANDA</div>""", unsafe_allow_html=True)

        with st.container(border=True):
            st.markdown("""
            <div style='font-size:.76rem;color:rgba(148,163,184,.55);
                margin-bottom:16px;line-height:1.5;'>
                Exporte qualquer módulo operacional para análise externa, auditoria ou arquivo.
                Arquivos gerados com timestamp automático.
            </div>
        """, unsafe_allow_html=True)
            d1, d2, d3, d4 = st.columns(4)

            with d1:
                st.download_button(
                    "📊 KPIs Consolidados (XLSX)",
                    data=gerar_xlsx_kpis(),
                    file_name=f"kpis_ti_{agora.strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with d2:
                if not df_rel.empty:
                    st.download_button(
                        "🎫 Banco de Chamados (XLSX)",
                        data=_xlsx_bytes({"Chamados": df_rel}),
                        file_name=f"chamados_{agora.strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.button("🎫 Banco de Chamados (XLSX)", disabled=True, use_container_width=True)
            with d3:
                if not df_rel.empty:
                    st.download_button(
                        "📄 Relatório de Incidentes (PDF)",
                        data=gerar_pdf(df_rel),
                        file_name=f"chamados_{agora.strftime('%Y%m%d_%H%M')}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                else:
                    st.button("📄 Relatório de Incidentes (PDF)", disabled=True, use_container_width=True)
            with d4:
                st.download_button(
                "📦 Pacote Completo (.ZIP)",
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
                        "🛒 Fluxo de Compras (XLSX)",
                        data=_xlsx_bytes({"Compras": df_comp}),
                        file_name=f"compras_{agora.strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                with d6:
                    st.download_button(
                        "📄 Auditoria Logística (PDF)",
                        data=gerar_pdf_compras(df_comp),
                        file_name=f"compras_{agora.strftime('%Y%m%d_%H%M')}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )

        # ── AGENDAMENTO AUTOMÁTICO CRONJOB ───────────────────────────
        st.markdown("---")
        st.markdown("""<div class="sec-label animate-fade-in"><div class="sec-dot"></div>AUTOMAÇÃO MENSAL & REPOSITÓRIO DE BACKUPS</div>""", unsafe_allow_html=True)

        EXPORT_DIR = "exports"
        if not os.path.exists(EXPORT_DIR):
            os.makedirs(EXPORT_DIR)

        flag_file = os.path.join(EXPORT_DIR, f"auto_{hoje.strftime('%Y%m')}.flag")
        ja_exportou = os.path.exists(flag_file)

        with st.container(border=True):
            col_ag1, col_ag2 = st.columns([1.8, 1.2])
            with col_ag1:
                st.markdown("""
                <div style='display:flex;align-items:center;gap:10px;margin-bottom:10px;'>
                    <span style='font-size:1.2rem;'>🎛️</span>
                    <div>
                        <div style='font-size:.82rem;font-weight:800;color:#e2e8f0;'>Motor de Fechamento Mensal</div>
                        <div style='font-size:.68rem;color:rgba(148,163,184,.50);'>Executa automaticamente no último dia do mês às 18h00</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
                st.markdown(f"""
                <p style='color:rgba(148,163,184,.60); font-size:.76rem; line-height:1.6; margin-bottom:15px;'>
                O microsserviço de auditoria monitora em tempo real a janela de fechamento do ciclo. 
                Ao atingir o limite, compila os logs automaticamente, salva em <code style='color:#38bdf8;background:rgba(56,189,248,.08);padding:1px 5px;border-radius:4px;'>exports/</code> e notifica a equipe.
                </p>
            """, unsafe_allow_html=True)

                if is_last and is_hora and not ja_exportou:
                    # Fechamento oficial do mês: sempre com os dados completos (df_rel_raw/df_comp_raw),
                    # nunca com o recorte de filtros que porventura esteja aplicado na tela.
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
                        "📊 RELATÓRIO MENSAL GERADO",
                        _corpo_rel,
                        url=_url_pagina("relatorios"),
                    )
                    st.success(f"✅ Relatório mensal de {hoje.strftime('%m/%Y')} gerado e salvo automaticamente!")
                elif ja_exportou:
                    st.success(f"✅ Status: Sincronização mensal de {hoje.strftime('%m/%Y')} validada e armazenada no servidor.")
                else:
                    prox = date(hoje.year if hoje.month < 12 else hoje.year+1,
                                hoje.month+1 if hoje.month < 12 else 1, 1)
                    dias_restantes = (prox - hoje).days - 1
                    st.info(f"⏰ Monitoramento em execução. Próximo fechamento em **{dias_restantes} dia(s)** (dia {ult_dia:02d}/{hoje.month:02d}/{hoje.year} às 18h00).")

            with col_ag2:
                st.markdown("""
                <div style='display:flex;align-items:center;gap:8px;margin-bottom:12px;'>
                    <span style='font-size:1.1rem;'>📁</span>
                    <div style='font-size:.80rem;font-weight:800;color:#e2e8f0;'>Repositório de Backups</div>
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
                            f"⬇️ {arq[:24]}...",
                            data=_ler_zip_cache(arq_path, _mtime),
                            file_name=arq,
                            mime="application/zip",
                            use_container_width=True,
                            key=f"dl_saved_{arq}"
                        )
                else:
                    st.caption("Nenhum snapshot automático localizado no cluster.")

        # ── TABELAS ANALÍTICAS DETALHADAS ────────────────────────────────────
        st.markdown("---")
        st.markdown("""<div class="sec-label animate-fade-in"><div class="sec-dot"></div>MATRIZ DE DADOS DE INFRAESTRUTURA E AUDITORIA</div>""", unsafe_allow_html=True)

        # ── Tabela HTML compacta com destaque (badges + barras de volume) ──
        st.markdown('''<style>
        .tbl-pro{width:100%;border-collapse:collapse;font-family:var(--font);font-size:.78rem;margin-top:2px;}
        .tbl-pro thead th{text-align:left;padding:8px 14px;color:#8aa0bb;font-weight:700;text-transform:uppercase;letter-spacing:.05em;font-size:.62rem;border-bottom:1px solid rgba(56,189,248,.16);white-space:nowrap;}
        .tbl-pro tbody td{padding:8px 14px;color:#e6edf6;border-bottom:1px solid rgba(148,163,184,.05);vertical-align:middle;}
        .tbl-pro tbody tr:hover{background:rgba(56,189,248,.05);}
        .tbl-pro tbody tr:last-child td{border-bottom:none;}
        .tp-bar{position:relative;background:rgba(56,189,248,.08);border-radius:5px;height:18px;min-width:78px;display:flex;align-items:center;overflow:hidden;}
        .tp-bar-fill{position:absolute;left:0;top:0;height:100%;border-radius:5px;background:linear-gradient(90deg,#1e3a5f,#38bdf8);}
        .tp-bar-val{position:relative;padding-left:8px;font-family:var(--mono);font-weight:700;font-size:.70rem;color:#eaf2fc;}
        .tp-badge{display:inline-block;padding:2px 10px;border-radius:99px;font-size:.68rem;font-weight:700;background:rgba(148,163,184,.12);color:#cbd5e1;}
        .tp-badge.b-red{background:rgba(248,113,113,.16);color:#fca5a5;}
        .tp-badge.b-amber{background:rgba(251,191,36,.16);color:#fcd34d;}
        .tp-badge.b-green{background:rgba(52,211,153,.16);color:#6ee7b7;}
        .tp-badge.b-blue{background:rgba(56,189,248,.16);color:#7dd3fc;}
        .tp-badge.b-orange{background:rgba(249,115,22,.16);color:#fdba74;}
        </style>''', unsafe_allow_html=True)

        _PRIO_BADGE = {"Baixa":"b-blue", "Média":"b-amber", "Alta":"b-orange", "Crítica":"b-red"}
        _STATUS_BADGE = {"Aberto":"b-amber","Em Atendimento":"b-blue","Concluído":"b-green","Pendente":"b-amber","Aprovado":"b-green","Em Análise":"b-blue","Cancelado":""}

        def _tabela_pro(df, bar_col=None, badge_map=None, headers=None):
            cols = list(df.columns)
            hd = headers or {}
            maxv = float(df[bar_col].max()) if (bar_col and bar_col in df.columns and len(df)) else 0.0
            th = "".join(f"<th>{html.escape(str(hd.get(c, c)))}</th>" for c in cols)
            body = []
            for _, r in df.iterrows():
                tds = []
                for c in cols:
                    val = r[c]
                    if bar_col and c == bar_col and maxv:
                        pct = int(float(val) / maxv * 100)
                        cell = (f'<div class="tp-bar"><div class="tp-bar-fill" style="width:{pct}%"></div>'
                                f'<span class="tp-bar-val">{html.escape(str(val))}</span></div>')
                    elif badge_map and c in badge_map:
                        cls = badge_map[c].get(str(val), "")
                        cell = f'<span class="tp-badge {cls}">{html.escape(str(val))}</span>'
                    else:
                        cell = html.escape(str(val))
                    tds.append(f"<td>{cell}</td>")
                body.append("<tr>" + "".join(tds) + "</tr>")
            st.markdown(f'<table class="tbl-pro"><thead><tr>{th}</tr></thead><tbody>{"".join(body)}</tbody></table>',
                        unsafe_allow_html=True)

        tab_t1, tab_t2, tab_t3, tab_t4 = st.tabs([
            "📋 Volumetria por Setor",
            "🚨 Criticidade e Prioridades",
            "🛒 Demandas Logísticas",
            "🧭 Categorias de Chamados"
        ])

        with tab_t1:
            with st.container(border=True):
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
            with st.container(border=True):
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
            with st.container(border=True):
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
            with st.container(border=True):
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

# ─────────────────────────────────────────────────────────────
# FIM DO ARQUIVO
# Linhas-tampão: protegem o código acima de eventual corrupção no
# final do arquivo durante a sincronização. Não remover.

# -----------------------------------------------------------------
# FIM DO ARQUIVO
# Linhas-tampao (ASCII puro, sem caracteres multibyte): protegem o
# codigo acima de corrupcao no final do arquivo durante o save.
# Nao remover.
# --------------------------------------------------------------
# --------------------------------------------------------------
# --------------------------------------------------------------
# --------------------------------------------------------------
