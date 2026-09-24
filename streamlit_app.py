from __future__ import annotations

from pathlib import Path
from io import BytesIO, StringIO
from datetime import datetime, timezone, timedelta
from zoneinfo import ZoneInfo
from email.utils import parsedate_to_datetime
import base64
import csv
import hashlib
import hmac
import html
import json
import math
import os
import re
import unicodedata

import pandas as pd
import altair as alt
import pydeck as pdk
import requests
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Painel Geográfico de Nodes — Rio Grande",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

ROOT = Path(__file__).resolve().parent
POA_TZ = ZoneInfo("America/Sao_Paulo")

def now_poa():
    return datetime.now(POA_TZ)

BASE = ROOT / "base_nodes_rio_grande.csv"
CONFIG = ROOT / "config_status.json"
DATA_DIR = ROOT / "data"
LOCAL_XLSX = DATA_DIR / "xpertrack_integridade_atual.xlsx"
LOCAL_CSV = DATA_DIR / "xpertrack_integridade_atual.csv"
REGION_FILE = DATA_DIR / "regioes_nodes.csv"
MANUAL_CROSSWALK_FILE = DATA_DIR / "cruzamento_manual.csv"
LOCATION_ANCHOR_FILE = DATA_DIR / "cruzamento_variantes_recuperadas.csv"
NEIGHBORHOOD_FILE = DATA_DIR / "bairros_nodes.csv"
TREATMENTS_FILE = DATA_DIR / "tratativas.csv"
HISTORY_FILE = DATA_DIR / "historico_crise.csv"
OUTAGES_CURRENT_FILE = DATA_DIR / "outages_atual.csv"
TEAM_FILE = DATA_DIR / "equipe_dia.csv"
ASSIGNMENTS_FILE = DATA_DIR / "atribuicoes.csv"
PHOTO_DIR = DATA_DIR / "fotos_atendimento"

# Ajustes operacionais validados em campo.
# Os nomes antigos permanecem disponíveis apenas como âncoras geográficas quando necessário,
# mas não aparecem como nodes ativos no mapa, busca ou indicadores.
RETIRED_NODES = {"GRE1", "PBLAB", "VIPAN", "VIPAM", "CNTAO", "VJDALA", "JCVAGA", "SSOAAB", "NODE INTER2", "NODE INTER3", "NODE INTER4", "NODE INTER5", "NODE INTER6"}

# Nodes UNI não possuem coleta no XPERTrack por definição. Eles permanecem no mapa,
# mas não devem ser tratados como falha/ausência de leitura.
UNI_NODES = {"SRDAJ", "SRDAG", "SRDAE", "SRDAC", "SRDAB", "VNVAT"}

# Correções de identidade: somente casos explicitamente validados.
IDENTITY_ALIAS_OVERRIDES = {
    "PRTAN": "PRTANA",
    "VJDALA": "VJDAL",
    "JCVAGA": "JCVAG",
    "SSOAAB": "SSOABB",
}

# Âncoras geográficas para cadastros atuais que reutilizam a posição de um cadastro antigo.
# Isso NÃO funde identidade de nodes.
LOCATION_ANCHOR_OVERRIDES = {
    "IGUI": "VIPAN",
    # SSOABB deve usar a posição correta do antigo ponto SSOAAB;
    # qualquer posição anterior de SSOABB é ignorada.
    "SSOABB": "SSOAAB",
    "INT1": "PBLAB",
    "INT2": "NODE INTER2",
    "INT3": "NODE INTER3",
    "INT4": "NODE INTER4",
    "INT5": "NODE INTER5",
    "INT6": "NODE INTER6",
}

# Origem/HUB confirmados pelo usuário. Internamente mantemos as chaves já usadas no painel.
ORIGIN_OVERRIDES = {
    "VNVAV": "HUB SUL",
    "FLOAHC": "HUB BV",
    "CTLAM": "HUB SUL",
    "BFMADC": "HUB BV",
    "RBRATA": "HUB BV",
    "URBRATB": "HUB BV",
    "AZNAEC": "HEADEND",
    "HMTAH": "HUB BV",
}

# Bairro/localidade confirmados para nodes fora do recorte de bairros de Porto Alegre.
NEIGHBORHOOD_OVERRIDES = {
    "VIA01": "VIAMAO",
    "UMBAL": "ALVORADA",
    "UMBAK": "ALVORADA",
    "ALV01": "ALVORADA",
    "ALV02": "ALVORADA",
}

LIGHT_MAP_STYLE = "https://basemaps.cartocdn.com/gl/voyager-gl-style/style.json"


def norm_txt(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip().upper()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s)


def norm_col(c) -> str:
    return re.sub(r"[^A-Z0-9]+", "_", norm_txt(c)).strip("_")


def esc(v) -> str:
    return html.escape(str(v if v is not None else ""))


def fmt_int(v):
    try:
        return f"{int(v):,}".replace(",", ".")
    except Exception:
        return "—"


def split_node_port(v, known_bases=None):
    """Normaliza o identificador do XPERTrack em (node lógico, porta).

    Mantém a regra padrão NODE-1..4 e acrescenta apenas formatos seguros/validados:
    - UMBAK1 / UMBAL1 quando a base sem o dígito existe no mapa;
    - ALV01-A / ALV02-A / VIA01-A quando a base existe no mapa (A=porta 1, B=porta 2);
    - INT1A..INT6A, que são os nodes atuais INT1..INT6 do Beira-Rio, cada um com uma leitura.
    """
    s = norm_txt(v)

    # Rio Grande: cadastro do XPERTrack sem hífen na porta 3.
    # TRVABA3 NÃO é um node: é o node TRVABA, porta 3.
    RG_COMPACT_PORTS = {"TRVABA3": ("TRVABA", 3)}
    if s in RG_COMPACT_PORTS:
        return RG_COMPACT_PORTS[s]

    # Formato normal: NODE-1, NODE-2... Mantém TRVACD e TRVACC
    # como identidades totalmente independentes.
    m = re.match(r"^(.+?)[\s_\-/]+([1-4])$", s)
    if m:
        return m.group(1).strip(), int(m.group(2))

    bases = {norm_txt(x) for x in (known_bases or set()) if norm_txt(x)}

    # Formato sem separador: UMBAK1 -> UMBAK / porta 1, somente quando a base é conhecida.
    m = re.match(r"^(.+?)([1-4])$", s)
    if m and norm_txt(m.group(1)) in bases:
        return norm_txt(m.group(1)), int(m.group(2))

    # Porta alfabética A/B ligada a um node base já existente: VIA01-A -> VIA01 / porta 1.
    m = re.match(r"^(.+?)-([AB])$", s)
    if m and norm_txt(m.group(1)) in bases:
        return norm_txt(m.group(1)), 1 if m.group(2) == "A" else 2

    # Cadastros atuais do Estádio Beira-Rio. INT1A..INT6A são nodes distintos,
    # cada um com uma única leitura no XPERTrack.
    m = re.match(r"^INT([1-6])A$", s)
    if m:
        return f"INT{m.group(1)}", 1

    return s, None


def get_secret(name: str, default=""):
    try:
        value = st.secrets.get(name, default)
        if value is None:
            return default
        return str(value)
    except Exception:
        return os.getenv(name, default)


def request_is_local() -> bool:
    try:
        host = str(st.context.headers.get("Host", "")).lower()
        return host.startswith("localhost") or host.startswith("127.0.0.1")
    except Exception:
        return False


def load_config() -> dict:
    cfg = {"crisis_threshold": 50, "refresh_minutes": 60, "stale_after_minutes": 90, "degraded_threshold": 20}
    try:
        if CONFIG.exists():
            cfg.update(json.loads(CONFIG.read_text(encoding="utf-8")) or {})
    except Exception:
        pass
    return cfg


CFG = load_config()
REFRESH_MINUTES = max(5, int(CFG.get("refresh_minutes", 60) or 60))
DEGRADED_THRESHOLD = max(1, int(CFG.get("degraded_threshold", 20) or 20))  # limite superior inclusivo da porta crítica

# Atualização horária sem dependência streamlit_autorefresh.
components.html(
    f"""<script>window.setTimeout(function(){{window.parent.location.reload();}}, {REFRESH_MINUTES * 60 * 1000});</script>""",
    height=0,
)

st.markdown(
    """
<style>
:root { color-scheme:light !important; --navy:#0b2443; --blue:#1677ff; --bg:#f4f7fb; --card:#fff; --text:#0b1736; --muted:#6b7892; --line:#e4eaf2; --yellow:#f5b700; --orange:#ff7a00; --red:#ef3340; --green:#10a55a; --cyan:#00a9d6; }
html, body, [class*="css"] { font-family: Inter, "Segoe UI", Arial, sans-serif; color-scheme:light !important; }
html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] { background:#f4f7fb !important; color:#0b1736 !important; }
[data-testid="stMarkdownContainer"], [data-testid="stCaptionContainer"], [data-testid="stText"], label, p, span { color:inherit; }
[data-baseweb="input"] > div, [data-baseweb="select"] > div, textarea, input { background:#ffffff !important; color:#0b1736 !important; -webkit-text-fill-color:#0b1736 !important; }
[data-baseweb="select"] span, [data-baseweb="input"] input, [data-testid="stNumberInput"] input, [data-testid="stTextInput"] input, [data-testid="stTextArea"] textarea { color:#0b1736 !important; -webkit-text-fill-color:#0b1736 !important; }
[data-testid="stForm"], [data-testid="stExpander"], [data-testid="stDataFrame"] { color:#0b1736 !important; }
[data-testid="stAppViewContainer"] { background:var(--bg); }
[data-testid="stHeader"] { background:transparent; }
[data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"] { display:none !important; }
#MainMenu, footer { visibility:hidden; }
.block-container { padding-top:.8rem; padding-bottom:1.5rem; max-width:1900px; }
.topbar { display:block; margin-bottom:9px; }
.title-wrap h1 { margin:0; font-size:30px; line-height:1.08; color:var(--text); letter-spacing:-.7px; }
.title-wrap .sub { color:#4b5d7d; font-size:13px; margin-top:6px; font-weight:650; }
.header-meta { margin-top:4px; color:#60708f; font-size:12.5px; line-height:1.45; }
.header-meta .warn-inline { color:#8a6811 !important; }
.sys { display:none; }
.dot { display:none; }
.dot.wait { display:none; }
.stale-note { display:none; }
.viewbar { margin:3px 0 10px; }
.kpi-grid { display:grid; grid-template-columns:repeat(7,minmax(0,1fr)); gap:9px; margin-bottom:9px; }
.kpi-card { background:#fff; border:1px solid var(--line); border-radius:14px; padding:12px 12px; min-height:105px; box-shadow:0 3px 12px rgba(12,38,75,.05); }
.kpi-label { color:#1b2b4b; font-size:11.5px; margin-bottom:5px; font-weight:650; }
.kpi-val { color:#071735; font-size:24px; line-height:1; font-weight:800; letter-spacing:-.5px; overflow-wrap:anywhere; }
.kpi-sub { color:#75829a; font-size:11px; margin-top:8px; }
.kpi-red .kpi-val { color:#d92c3a; } .kpi-orange .kpi-val { color:#ec6e00; } .kpi-yellow .kpi-val { color:#b88700; } .kpi-blue .kpi-val { color:#0c66d6; } .kpi-green .kpi-val { color:#078d4a; }
.mini-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(175px,1fr)); gap:8px; margin:0 0 10px; }
.mini-card { background:#fff; border:1px solid var(--line); border-radius:12px; padding:10px 12px; min-height:74px; box-shadow:0 3px 12px rgba(12,38,75,.04); }
.mini-label { color:#5f6f8b; font-size:10.7px; font-weight:650; } .mini-val { color:#0b1736; font-size:21px; font-weight:800; margin-top:4px; line-height:1; } .mini-sub { color:#7a879c; font-size:9.7px; margin-top:5px; }
.panel { background:#fff; border:1px solid var(--line); border-radius:14px; padding:12px 13px; box-shadow:0 3px 12px rgba(12,38,75,.045); overflow-x:auto; }
.panel-title { display:flex; justify-content:space-between; align-items:center; gap:8px; margin:0 0 8px; }
.panel-title b { color:#13213f; font-size:17px; } .panel-title span { color:#0d6efd; font-size:10.8px; }
.rank { width:100%; border-collapse:collapse; font-size:11.3px; min-width:390px; }
.rank th { text-align:left; padding:7px 5px; color:#6d7890; font-weight:600; border-bottom:1px solid #e8edf4; } .rank td { padding:7px 5px; color:#1f2d49; border-bottom:1px solid #eef2f7; }
.badge { display:inline-block; padding:4px 7px; border-radius:999px; font-size:9.6px; font-weight:750; white-space:nowrap; }
.badge-normal { background:#e8f8ef; color:#0a8f4d; } .badge-deg { background:#fff5cf; color:#8c6900; border:1px solid #e3b600; } .badge-leve { background:#fff5cf; color:#8c6900; } .badge-critico { background:#fff5cf; color:#8c6900; } .badge-off { background:#ffe1e4; color:#c72231; } .badge-sem { background:#eef1f5; color:#647089; } .badge-atend { background:#e2f7fc; color:#007e9f; } .badge-tratado { background:#e8f8ef; color:#0a8f4d; } .badge-semt { background:#f0f2f5; color:#667085; }
.legend-card { background:#fff; border:1px solid var(--line); border-radius:12px; padding:7px 10px; margin-top:7px; display:flex; align-items:center; gap:8px 14px; flex-wrap:wrap; }
.legend-title { color:#13213f; font-size:11px; font-weight:800; margin-right:2px; }
.legend-line { display:flex; align-items:center; gap:5px; margin:0; color:#263654; font-size:9.7px; white-space:nowrap; }
.sw { width:9px; height:9px; border-radius:50%; display:inline-block; flex:0 0 auto; }
.ring { width:11px; height:11px; border-radius:50%; display:inline-block; background:#fff; border:2px solid #00a9d6; box-sizing:border-box; flex:0 0 auto; }
.kpi-action { display:inline-block; margin-top:5px; padding:4px 9px; border-radius:8px; background:#0c66d6; color:#fff !important; text-decoration:none !important; font-weight:750; font-size:10px; line-height:1.2; }
.kpi-action:hover { background:#0958b8; color:#fff !important; }
.action-high { font-weight:800; color:#c72231; } .action-mid { font-weight:750; color:#be5700; } .action-ok { font-weight:750; color:#078d4a; }
.map-focus-card { background:#fff; border:1px solid #cfe0f5; border-radius:12px; padding:10px 12px; margin:0 0 8px; box-shadow:0 3px 10px rgba(12,38,75,.06); color:#0b1736 !important; }
.map-focus-name { font-size:16px; font-weight:800; color:#0b1736 !important; } .map-focus-meta { margin-top:3px; font-size:11px; color:#4c5d79 !important; } .map-focus-ports { margin-top:6px; font-size:11px; font-weight:650; color:#1f2d49 !important; } .map-focus-trat { margin-top:5px; font-size:10.5px; color:#51617c !important; }
.mobile-hint { display:none; }
@media (max-width:768px) {
 .block-container { padding:.45rem .6rem 1.2rem !important; max-width:100% !important; }
 .topbar { display:block; } .title-wrap h1 { font-size:21px; } .title-wrap .sub { font-size:10.8px; }
 .header-meta { font-size:10px; margin-top:4px; }
 .stale-note { display:none; }
 .kpi-grid { grid-template-columns:repeat(2,minmax(0,1fr)); gap:6px; } .kpi-card { min-height:87px; padding:9px 10px; border-radius:11px; } .kpi-label{font-size:10px}.kpi-val{font-size:20px}.kpi-sub{font-size:8.9px;margin-top:6px}
 .mini-grid { grid-template-columns:repeat(2,minmax(0,1fr)); gap:6px; } .mini-card{min-height:66px;padding:8px 9px}.mini-label{font-size:9.3px}.mini-val{font-size:18px}.mini-sub{font-size:8.5px}
 [data-testid="stHorizontalBlock"] { flex-wrap:wrap !important; gap:.45rem !important; } [data-testid="stHorizontalBlock"] > [data-testid="column"] { flex:1 1 100% !important; width:100% !important; min-width:100% !important; }
 .panel-title b{font-size:14.5px;color:#13213f !important}.panel-title span{font-size:9px;color:#0d6efd !important}.rank{font-size:9.6px;min-width:350px;background:#fff !important}.rank th,.rank td{padding:5px 4px;color:#1f2d49 !important}.badge{font-size:8.2px;padding:3px 5px}.mobile-hint{display:block;color:#53627d !important;font-size:9.5px;margin:3px 0 6px}
 .map-focus-card{padding:9px 10px}.map-focus-name{font-size:14px}.map-focus-meta,.map-focus-ports,.map-focus-trat{font-size:9.5px}
}
</style>
""",
    unsafe_allow_html=True,
)


def kpi_html(label, value, sub, cls="", sub_is_html=False):
    sub_content = str(sub) if sub_is_html else esc(sub)
    return f'<div class="kpi-card {cls}"><div class="kpi-label">{esc(label)}</div><div class="kpi-val">{esc(value)}</div><div class="kpi-sub">{sub_content}</div></div>'


def mini_html(label, value, sub=""):
    return f'<div class="mini-card"><div class="mini-label">{esc(label)}</div><div class="mini-val">{esc(value)}</div><div class="mini-sub">{esc(sub)}</div></div>'


def status_publico(status):
    return {
        "NORMAL":"ONLINE",
        "DEGRADADO":"ONLINE · PORTA CRÍTICA",
        "PARCIAL":"SS PARCIAL",
        "PARCIAL CRÍTICO":"SS PARCIAL",  # compatibilidade com históricos antigos
        "OFF TOTAL":"SS TOTAL",
        "SEM DADO":"AGUARDANDO STATUS",
        "UNI":"UNI · SEM COLETA XPERTRACK",
    }.get(str(status), str(status))


def badge(status):
    cl = {"NORMAL":"normal", "DEGRADADO":"deg", "PARCIAL":"critico", "PARCIAL CRÍTICO":"critico", "OFF TOTAL":"off"}.get(status, "sem")
    return f'<span class="badge badge-{cl}">{esc(status_publico(status))}</span>'


def treatment_badge(v):
    s = norm_txt(v)
    if s == "EM ATENDIMENTO": return '<span class="badge badge-atend">EM ATENDIMENTO</span>'
    if s == "TRATADO / NORMALIZADO": return '<span class="badge badge-tratado">TRATADO</span>'
    return '<span class="badge badge-semt">SEM ATENDIMENTO</span>'


# -----------------------------
# Arquivos de referência
# -----------------------------
@st.cache_data(show_spinner=False)
def load_manual_crosswalk() -> pd.DataFrame:
    cols = ["XPER_Node","Node_Correto","Map_Node","Regiao","Bairro","Endereco","Observacao","Latitude","Longitude","Precisao"]
    if not MANUAL_CROSSWALK_FILE.exists(): return pd.DataFrame(columns=cols)
    try:
        d = pd.read_csv(MANUAL_CROSSWALK_FILE)
        for c in cols:
            if c not in d.columns: d[c] = ""
        for c in ["XPER_Node","Node_Correto","Map_Node","Regiao","Bairro"]:
            d[c] = d[c].fillna("").map(norm_txt)
        d["Latitude"] = pd.to_numeric(d["Latitude"], errors="coerce")
        d["Longitude"] = pd.to_numeric(d["Longitude"], errors="coerce")
        return d[cols].copy()
    except Exception:
        return pd.DataFrame(columns=cols)


@st.cache_data(show_spinner=False)
def load_location_anchors() -> dict:
    """Crosswalk histórico usado SOMENTE para localizar um node no mapa.
    Não funde identidade de nodes. HPCANA e HPCANB, por exemplo, continuam distintos.
    """
    out = {}
    if LOCATION_ANCHOR_FILE.exists():
        try:
            d = pd.read_csv(LOCATION_ANCHOR_FILE)
            cols = {norm_col(c): c for c in d.columns}
            if "XPER_NODE" in cols and "MAP_NODE" in cols:
                for _, r in d.iterrows():
                    x = norm_txt(r[cols["XPER_NODE"]]); m = norm_txt(r[cols["MAP_NODE"]])
                    if x and m: out[x] = m
        except Exception: pass
    manual = load_manual_crosswalk()
    if not manual.empty:
        for _, r in manual.iterrows():
            x = norm_txt(r.get("XPER_Node")); m = norm_txt(r.get("Map_Node"))
            if x and m: out[x] = m
            corr = norm_txt(r.get("Node_Correto"))
            if corr and m: out[corr] = m
    out.update(LOCATION_ANCHOR_OVERRIDES)
    return out


@st.cache_data(show_spinner=False)
def load_identity_aliases() -> dict:
    aliases = {}
    manual = load_manual_crosswalk()
    if not manual.empty:
        for _, r in manual.iterrows():
            x = norm_txt(r.get("XPER_Node")); corr = norm_txt(r.get("Node_Correto"))
            if x and corr and x != corr: aliases[x] = corr
    # Correções de identidade explicitamente validadas.
    aliases.update(IDENTITY_ALIAS_OVERRIDES)
    return aliases


@st.cache_data(show_spinner=False)
def load_region_map() -> pd.DataFrame:
    if not REGION_FILE.exists(): return pd.DataFrame(columns=["Node","Regiao_Mapa"])
    try:
        r = pd.read_csv(REGION_FILE)
        cols = {norm_col(c): c for c in r.columns}
        if "NODE" not in cols or "REGIAO" not in cols: return pd.DataFrame(columns=["Node","Regiao_Mapa"])
        out = r[[cols["NODE"], cols["REGIAO"]]].copy(); out.columns=["Node","Regiao_Mapa"]
        out["Node"] = out["Node"].map(norm_txt); out["Regiao_Mapa"] = out["Regiao_Mapa"].fillna("").map(norm_txt)
        return out.drop_duplicates("Node")
    except Exception: return pd.DataFrame(columns=["Node","Regiao_Mapa"])


@st.cache_data(show_spinner=False)
def load_manual_neighborhood_map() -> dict:
    if not NEIGHBORHOOD_FILE.exists(): return dict(NEIGHBORHOOD_OVERRIDES)
    try:
        d = pd.read_csv(NEIGHBORHOOD_FILE)
        cols = {norm_col(c):c for c in d.columns}
        if "NODE" not in cols or "BAIRRO" not in cols: return {}
        out={norm_txt(r[cols["NODE"]]): norm_txt(r[cols["BAIRRO"]]) for _,r in d.iterrows() if norm_txt(r[cols["NODE"]]) and norm_txt(r[cols["BAIRRO"]])}
        out.update(NEIGHBORHOOD_OVERRIDES)
        return out
    except Exception: return dict(NEIGHBORHOOD_OVERRIDES)


@st.cache_data(show_spinner=False)
def load_base() -> pd.DataFrame:
    df = pd.read_csv(BASE)
    for c in ["Node","RX","Origem","Status"]:
        if c not in df.columns: df[c]=""
        df[c]=df[c].fillna("").astype(str)
    df["Node"] = df["Node"].map(norm_txt); df["Origem"] = df["Origem"].map(norm_txt); df["RX"] = df["RX"].map(norm_txt)
    # Proteção contra base antiga: TRVABA3 é porta 3 do TRVABA, nunca node físico.
    df["Node"] = df["Node"].replace({"TRVABA3": "TRVABA"})
    df = df.drop_duplicates("Node", keep="first").copy()
    df["Latitude"] = pd.to_numeric(df["Latitude"], errors="coerce"); df["Longitude"] = pd.to_numeric(df["Longitude"], errors="coerce")
    return df


# -----------------------------
# Leitura XPERTrack
# -----------------------------
@st.cache_data(show_spinner=False)
def read_csv_bytes(data: bytes) -> pd.DataFrame:
    last_error=None
    for enc in ("utf-8-sig","utf-8","latin-1"):
        try:
            text=data.decode(enc, errors="strict")
            df=pd.read_csv(StringIO(text), sep=None, engine="python")
            if df.shape[1]==1 and "," in str(df.columns[0]):
                outer=list(csv.reader(StringIO(text))); inner=[]
                for row in outer:
                    inner.append(next(csv.reader([row[0]])) if len(row)==1 else row)
                if inner and len(inner[0])>1:
                    width=len(inner[0]); good=[r for r in inner[1:] if len(r)==width]
                    return pd.DataFrame(good, columns=inner[0])
            return df
        except Exception as e: last_error=e
    raise ValueError(f"Não foi possível abrir o CSV: {last_error}")


@st.cache_data(show_spinner=False)
def read_excel_bytes(data: bytes) -> pd.DataFrame:
    return pd.read_excel(BytesIO(data))


def load_table_bytes(data: bytes, name_hint: str):
    low=(name_hint or "").lower()
    if low.endswith(".csv"): return read_csv_bytes(data)
    if low.endswith((".xlsx",".xls")): return read_excel_bytes(data)
    try: return read_csv_bytes(data)
    except Exception: return read_excel_bytes(data)


@st.cache_data(ttl=300, show_spinner=False)
def fetch_remote_bytes(url: str, token: str=""):
    headers={"User-Agent":"Painel-Nodes-RG/10.0.1"}
    if token: headers["Authorization"]=f"Bearer {token}"
    r=requests.get(url, headers=headers, timeout=35); r.raise_for_status()
    return r.content, r.headers.get("Last-Modified","")


def xpertrack_schema(df: pd.DataFrame):
    by={norm_col(c):c for c in df.columns}
    if "NODE" not in by or "PONTUACAO" not in by: return None
    return {"NODE":by["NODE"], "PONTUACAO":by["PONTUACAO"], "IMPACTADO":by.get("IMPACTADO"), "ESTRESSADO":by.get("ESTRESSADO"), "TOTAL":by.get("TOTAL"), "REGIAO":by.get("REGIAO") or by.get("REGION") or by.get("ZONA")}


def score_to_port_state(score):
    if score is None or pd.isna(score): return ""
    try: v=float(score)
    except Exception: return ""
    if v==0: return "OFF"
    if 0<v<=DEGRADED_THRESHOLD: return "CRÍTICA"
    if v>DEGRADED_THRESHOLD: return "ON"
    return ""


def build_operational_status(xdf: pd.DataFrame):
    schema=xpertrack_schema(xdf)
    if not schema: return pd.DataFrame(), pd.DataFrame()
    d=xdf.copy(); known_bases=set(load_base()["Node"].map(norm_txt).tolist()); parsed=d[schema["NODE"]].map(lambda v: split_node_port(v, known_bases))
    d["_raw_node"]=[x[0] for x in parsed]; d["_port"]=[x[1] for x in parsed]
    aliases=load_identity_aliases(); d["_node"]=d["_raw_node"].map(lambda n:aliases.get(norm_txt(n),norm_txt(n)))
    d["_score"]=pd.to_numeric(d[schema["PONTUACAO"]], errors="coerce")
    d=d[d["_node"].astype(str).str.len()>0 & d["_port"].isin([1,2,3,4])].copy()
    d=d[~d["_node"].isin(RETIRED_NODES)].copy()
    if d.empty: return pd.DataFrame(), d
    d["_region"] = d[schema["REGIAO"]].fillna("").map(norm_txt) if schema.get("REGIAO") else ""
    for src,tgt in [("IMPACTADO","_impactado"),("ESTRESSADO","_estressado"),("TOTAL","_total")]:
        col=schema.get(src); d[tgt]=pd.to_numeric(d[col],errors="coerce") if col else 0
    rows=[]
    for node,g in d.groupby("_node", dropna=False):
        scores={}; states={}
        for p in range(1,5):
            vals=g.loc[g["_port"]==p,"_score"].dropna(); score=float(vals.min()) if len(vals) else None
            scores[p]=score; states[p]=score_to_port_state(score)
        existing=[p for p in range(1,5) if scores[p] is not None]
        total_ports=len(existing); off=sum(states[p]=="OFF" for p in existing); deg=sum(states[p]=="CRÍTICA" for p in existing); on=sum(states[p]=="ON" for p in existing)
        if total_ports==0: status="SEM DADO"
        elif off==total_ports: status="OFF TOTAL"
        elif off>=1: status="PARCIAL"
        elif deg>0: status="DEGRADADO"
        elif on==total_ports: status="NORMAL"
        else: status="SEM DADO"
        regs=g.loc[g["_region"].astype(str).str.len()>0,"_region"]; reg=""
        if len(regs):
            md=regs.mode(); reg=str(md.iloc[0] if not md.empty else regs.iloc[0])
        row={"Node":node,"Status":status,"Total_portas":total_ports,"Portas_OFF":int(off),"Portas_Degradadas":int(deg),"Portas_Criticas":int(deg),"Portas_ON":int(on),"Regiao":reg,"Impactado_integridade":int(g["_impactado"].fillna(0).sum()),"Estressado_integridade":int(g["_estressado"].fillna(0).sum()),"Total_integridade":int(g["_total"].fillna(0).sum()),"Raw_Nodes":", ".join(sorted(g["_raw_node"].dropna().map(norm_txt).unique().tolist()))}
        for p in range(1,5): row[f"Porta_{p}"]=states[p]; row[f"Nota_P{p}"]=scores[p]
        rows.append(row)
    return pd.DataFrame(rows), d


# -----------------------------
# Bairro oficial — GIS Prefeitura de Porto Alegre
# -----------------------------
BAIRROS_API = ""

@st.cache_data(ttl=86400, show_spinner=False)
def fetch_bairros_geojson():
    try:
        if not BAIRROS_API:
            return None
        params={"where":"1=1","outFields":"nome","returnGeometry":"true","f":"geojson","outSR":"4326"}
        r=requests.get(BAIRROS_API, params=params, timeout=30, headers={"User-Agent":"Painel-Nodes-RG/10.0.1"}); r.raise_for_status()
        js=r.json()
        return js if js.get("features") else None
    except Exception:
        return None


def _point_in_ring(x,y,ring):
    inside=False; n=len(ring)
    if n<3: return False
    j=n-1
    for i in range(n):
        xi,yi=ring[i][0],ring[i][1]; xj,yj=ring[j][0],ring[j][1]
        cross=((yi>y)!=(yj>y)) and (x < (xj-xi)*(y-yi)/((yj-yi) or 1e-15)+xi)
        if cross: inside=not inside
        j=i
    return inside


def _point_in_geometry(lon,lat,geom):
    typ=geom.get("type"); coords=geom.get("coordinates",[])
    polys=[coords] if typ=="Polygon" else coords if typ=="MultiPolygon" else []
    for poly in polys:
        if not poly: continue
        if _point_in_ring(lon,lat,poly[0]):
            if not any(_point_in_ring(lon,lat,h) for h in poly[1:]): return True
    return False


def build_neighborhood_index(geojson):
    out=[]
    if not geojson: return out
    for f in geojson.get("features",[]):
        geom=f.get("geometry") or {}; name=norm_txt((f.get("properties") or {}).get("nome",""))
        if not name: continue
        vals=[]
        def walk(o):
            if isinstance(o,(list,tuple)):
                if len(o)>=2 and isinstance(o[0],(int,float)) and isinstance(o[1],(int,float)): vals.append((o[0],o[1]))
                else:
                    for z in o: walk(z)
        walk(geom.get("coordinates",[]))
        if vals:
            xs=[v[0] for v in vals]; ys=[v[1] for v in vals]
            out.append((name,geom,(min(xs),min(ys),max(xs),max(ys))))
    return out


def neighborhood_for_point(lon,lat,index):
    if pd.isna(lon) or pd.isna(lat): return ""
    for name,geom,b in index:
        if b[0] <= lon <= b[2] and b[1] <= lat <= b[3] and _point_in_geometry(float(lon),float(lat),geom): return name
    return ""


# -----------------------------
# Consistência de região por bairro
# -----------------------------
VALID_REGIONS = {"NORTE", "CENTRAL", "SUL"}

# Exceções podem ser adicionadas aqui somente quando validadas operacionalmente.
# Petrópolis foi usado como caso de validação: todos os nodes do bairro pertencem à mesma região.
BAIRRO_REGION_OVERRIDES = {
    "PETROPOLIS": "CENTRAL",
}

# Região validada operacionalmente por node. Aplicada por último para não ser
# sobrescrita pela consolidação automática do bairro.
NODE_REGION_OVERRIDES = {
    "MDRAA": "SUL",
    "MDRAB": "SUL",
    "CCTAA": "SUL",
    "CCTABB": "SUL",
    # Cadastros existentes na base geográfica correspondentes à família CCT.
    "CCTAAA": "SUL",
    "CCTAAB": "SUL",
}

def build_bairro_region_map(df: pd.DataFrame) -> dict:
    """Define uma única região operacional para cada bairro.

    A regra usa a maioria dos nodes físicos (Anchor_Node) já classificados no
    mapa, evitando que variantes XPERTrack do mesmo ponto distorçam a contagem.
    Depois aplica exceções explicitamente validadas.
    """
    if df is None or df.empty or "Bairro" not in df.columns or "Regiao" not in df.columns:
        return dict(BAIRRO_REGION_OVERRIDES)
    d = df.copy()
    d["_bairro"] = d["Bairro"].fillna("").map(norm_txt)
    d["_regiao"] = d["Regiao"].fillna("").map(norm_txt)
    if "Anchor_Node" in d.columns:
        d["_anchor"] = d["Anchor_Node"].fillna("").map(norm_txt)
    else:
        d["_anchor"] = d.get("Node", "").fillna("").map(norm_txt)
    d = d[(d["_bairro"].str.len() > 0) & d["_regiao"].isin(VALID_REGIONS)].copy()
    if d.empty:
        return dict(BAIRRO_REGION_OVERRIDES)
    d = d.drop_duplicates(["_bairro", "_anchor"])
    counts = d.groupby(["_bairro", "_regiao"])["_anchor"].nunique().reset_index(name="n")
    mapping = {}
    for bairro, g in counts.groupby("_bairro"):
        g = g.sort_values(["n", "_regiao"], ascending=[False, True])
        if len(g):
            mapping[bairro] = str(g.iloc[0]["_regiao"])
    mapping.update(BAIRRO_REGION_OVERRIDES)
    return mapping

def apply_bairro_region_consistency(df: pd.DataFrame, bairro_region_map: dict) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    out = df.copy()
    if "Bairro" not in out.columns:
        out["Bairro"] = ""
    if "Regiao" not in out.columns:
        out["Regiao"] = ""
    out["Bairro"] = out["Bairro"].fillna("").map(norm_txt)
    out["Regiao"] = out["Regiao"].fillna("").map(norm_txt)
    canonical = out["Bairro"].map(lambda b: bairro_region_map.get(norm_txt(b), ""))
    mask = canonical.astype(str).str.len() > 0
    out.loc[mask, "Regiao"] = canonical.loc[mask]
    return out

# -----------------------------
# Localização sem fundir identidade
# -----------------------------
def _port_label(state,p):
    if state=="ON": return f"P{p} 🟢 ONLINE"
    if state=="OFF": return f"P{p} 🔴 OFF"
    if state=="CRÍTICA": return f"P{p} 🟡 CRÍTICA"
    return ""


def build_logical_geo(status_df: pd.DataFrame, base_df: pd.DataFrame, reg_map: pd.DataFrame):
    if status_df is None or status_df.empty: return pd.DataFrame()
    base_idx=base_df.drop_duplicates("Node",keep="first").set_index("Node")
    anchors=load_location_anchors(); manual=load_manual_crosswalk(); manual_bairro=load_manual_neighborhood_map()
    manual_by={}
    for _,r in manual.iterrows():
        keys={norm_txt(r.get("XPER_Node")), norm_txt(r.get("Node_Correto"))}
        for k in keys:
            if k: manual_by[k]=r
    reg_dict={norm_txt(r.Node):norm_txt(r.Regiao_Mapa) for r in reg_map.itertuples()} if reg_map is not None and not reg_map.empty else {}
    neigh_index=build_neighborhood_index(fetch_bairros_geojson())
    rows=[]
    for _,r in status_df.iterrows():
        node=norm_txt(r.get("Node")); mr=manual_by.get(node)
        anchor=""; lat=lon=float("nan"); origem=""; precision=""; bairro=manual_bairro.get(node,"")
        if node in base_idx.index:
            anchor=node; br=base_idx.loc[node]; lat=br["Latitude"]; lon=br["Longitude"]; origem=norm_txt(br.get("Origem")); precision="KMZ_EXATO"
        elif mr is not None and pd.notna(mr.get("Latitude")) and pd.notna(mr.get("Longitude")):
            anchor=node; lat=float(mr.get("Latitude")); lon=float(mr.get("Longitude")); origem="NAO INFORMADA"; precision=str(mr.get("Precisao","") or "MANUAL")
            fallback_anchor=LOCATION_ANCHOR_OVERRIDES.get(node) or anchors.get(node,"")
            if fallback_anchor in base_idx.index:
                origem=norm_txt(base_idx.loc[fallback_anchor].get("Origem")) or origem
        else:
            manual_anchor=norm_txt(mr.get("Map_Node")) if mr is not None else ""
            anchor=manual_anchor or anchors.get(node,"")
            if anchor in base_idx.index:
                br=base_idx.loc[anchor]; lat=br["Latitude"]; lon=br["Longitude"]; origem=norm_txt(br.get("Origem")); precision="ANCORA_KMZ"
        reg=norm_txt(r.get("Regiao",""))
        if mr is not None and norm_txt(mr.get("Regiao")): reg=norm_txt(mr.get("Regiao"))
        if not reg: reg=reg_dict.get(anchor or node,"")
        if not bairro and mr is not None and norm_txt(mr.get("Bairro")): bairro=norm_txt(mr.get("Bairro"))
        if not bairro and not pd.isna(lat) and not pd.isna(lon): bairro=neighborhood_for_point(float(lon),float(lat),neigh_index)
        bairro=NEIGHBORHOOD_OVERRIDES.get(node,bairro)
        origem=ORIGIN_OVERRIDES.get(node,origem)
        parts=[]
        for p in range(1,5):
            z=_port_label(str(r.get(f"Porta_{p}","") or ""),p)
            if z: parts.append(z)
        row=r.to_dict(); row.update({"Anchor_Node":anchor,"Latitude_Base":lat,"Longitude_Base":lon,"Latitude":lat,"Longitude":lon,"Origem":origem,"Regiao":reg,"Bairro":bairro,"Precisao_Local":precision,"Portas_popup":"  |  ".join(parts) if parts else "Sem leitura de porta no XPERTrack"})
        rows.append(row)
    out=pd.DataFrame(rows)
    # Se vários nodes lógicos compartilham o mesmo ponto físico, apenas desloca o marcador visualmente.
    located=out[out["Latitude_Base"].notna() & out["Longitude_Base"].notna()].copy()
    for _,idxs in located.groupby([located["Latitude_Base"].round(6),located["Longitude_Base"].round(6)]).groups.items():
        idxs=list(idxs)
        if len(idxs)<=1: continue
        n=len(idxs); radius=0.00018
        for j,idx in enumerate(sorted(idxs, key=lambda z:str(out.loc[z,"Node"]))):
            ang=2*math.pi*j/n; lat0=float(out.loc[idx,"Latitude_Base"]); lon0=float(out.loc[idx,"Longitude_Base"])
            out.loc[idx,"Latitude"] = lat0 + radius*math.sin(ang)
            out.loc[idx,"Longitude"] = lon0 + radius*math.cos(ang)/max(0.5,math.cos(math.radians(lat0)))
    return out


def build_map_frame(logical_geo: pd.DataFrame, base_df: pd.DataFrame, reg_map: pd.DataFrame):
    used=set(logical_geo["Anchor_Node"].dropna().map(norm_txt).tolist()) if not logical_geo.empty else set()
    placeholders=base_df[(~base_df["Node"].isin(used)) & (~base_df["Node"].isin(RETIRED_NODES))].copy()
    reg_dict={norm_txt(r.Node):norm_txt(r.Regiao_Mapa) for r in reg_map.itertuples()} if reg_map is not None and not reg_map.empty else {}
    manual_bairro=load_manual_neighborhood_map(); neigh_index=build_neighborhood_index(fetch_bairros_geojson())
    if not placeholders.empty:
        placeholders["Status"]="SEM DADO"; placeholders["Total_portas"]=0; placeholders["Portas_OFF"]=0; placeholders["Portas_Degradadas"]=0; placeholders["Portas_Criticas"]=0; placeholders["Portas_ON"]=0
        uni_mask=placeholders["Node"].map(norm_txt).isin(UNI_NODES)
        placeholders.loc[uni_mask,"Status"]="UNI"
        placeholders["Regiao"]=placeholders["Node"].map(reg_dict).fillna(""); placeholders["Bairro"]=placeholders["Node"].map(manual_bairro).fillna("")
        placeholders["Origem"]=placeholders.apply(lambda r:ORIGIN_OVERRIDES.get(norm_txt(r.get("Node")),norm_txt(r.get("Origem"))),axis=1)
        if neigh_index:
            miss=placeholders["Bairro"].astype(str).str.len()==0
            placeholders.loc[miss,"Bairro"] = placeholders.loc[miss].apply(lambda r: neighborhood_for_point(r["Longitude"],r["Latitude"],neigh_index),axis=1)
        placeholders["Portas_popup"]="Sem leitura de porta no XPERTrack"
        placeholders.loc[uni_mask,"Portas_popup"]="Node UNI · sem coleta no XPERTrack"
        placeholders["Anchor_Node"]=placeholders["Node"]; placeholders["Precisao_Local"]="KMZ"; placeholders["Latitude_Base"]=placeholders["Latitude"]; placeholders["Longitude_Base"]=placeholders["Longitude"]
        for p in range(1,5): placeholders[f"Porta_{p}"]=""; placeholders[f"Nota_P{p}"]=pd.NA
    cols=set(placeholders.columns).union(set(logical_geo.columns))
    for d in (placeholders,logical_geo):
        for c in cols:
            if c not in d.columns: d[c]=""
    return pd.concat([logical_geo[list(cols)],placeholders[list(cols)]], ignore_index=True)


def apply_node_region_overrides(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    out=df.copy()
    if "Node" not in out.columns:
        return out
    if "Regiao" not in out.columns:
        out["Regiao"]=""
    nodes=out["Node"].fillna("").map(norm_txt)
    forced=nodes.map(NODE_REGION_OVERRIDES)
    mask=forced.notna() & forced.astype(str).str.len().gt(0)
    out.loc[mask,"Regiao"]=forced.loc[mask]
    return out


# -----------------------------
# GitHub/local persistence
# -----------------------------
def github_settings():
    return {
        "repo":get_secret("GITHUB_REPO","").strip(), "branch":get_secret("GITHUB_BRANCH","main").strip() or "main", "token":get_secret("GITHUB_TOKEN","").strip(),
        "treatments":get_secret("TRATATIVAS_PATH","data/tratativas.csv").strip() or "data/tratativas.csv", "history":get_secret("HISTORICO_CRISE_PATH","data/historico_crise.csv").strip() or "data/historico_crise.csv",
        "outages_current":get_secret("OUTAGES_ATUAL_PATH","data/outages_atual.csv").strip() or "data/outages_atual.csv",
        "team":get_secret("EQUIPE_DIA_PATH","data/equipe_dia.csv").strip() or "data/equipe_dia.csv", "assignments":get_secret("ATRIBUICOES_PATH","data/atribuicoes.csv").strip() or "data/atribuicoes.csv", "photos":get_secret("FOTOS_PATH","data/fotos_atendimento").strip().strip("/") or "data/fotos_atendimento",
    }


def github_headers(token): return {"Authorization":f"Bearer {token}","Accept":"application/vnd.github+json","X-GitHub-Api-Version":"2022-11-28","User-Agent":"Painel-Nodes-RG/10.0.1"}


@st.cache_data(ttl=120,show_spinner=False)
def github_file_updated_at(path:str):
    """Data do último commit do arquivo XPERTrack no repositório; evita usar o relógio do servidor Streamlit."""
    gh=github_settings()
    if not (gh["repo"] and gh["token"] and path): return None
    try:
        url=f"https://api.github.com/repos/{gh['repo']}/commits"
        r=requests.get(url,headers=github_headers(gh["token"]),params={"path":path,"sha":gh["branch"],"per_page":1},timeout=20)
        r.raise_for_status(); items=r.json()
        if not items: return None
        iso=((items[0].get("commit") or {}).get("committer") or {}).get("date") or ((items[0].get("commit") or {}).get("author") or {}).get("date")
        return datetime.fromisoformat(str(iso).replace("Z","+00:00")) if iso else None
    except Exception:
        return None


def _gh_read_bytes(path):
    gh=github_settings()
    if not (gh["repo"] and gh["token"]): return None
    try:
        url=f"https://api.github.com/repos/{gh['repo']}/contents/{path}"; r=requests.get(url,headers=github_headers(gh["token"]),params={"ref":gh["branch"]},timeout=20)
        if r.status_code==404: return None
        r.raise_for_status(); return base64.b64decode(r.json().get("content",""))
    except Exception: return None


@st.cache_data(ttl=60,show_spinner=False)
def _gh_list_data_files():
    """Lista os arquivos atuais da pasta data diretamente no GitHub.

    Isso evita depender do checkout local do Streamlit Cloud para enxergar uma
    nova extração diária recém-enviada ao repositório.
    """
    gh=github_settings()
    if not (gh["repo"] and gh["token"]): return []
    try:
        url=f"https://api.github.com/repos/{gh['repo']}/contents/data"
        r=requests.get(url,headers=github_headers(gh["token"]),params={"ref":gh["branch"]},timeout=20)
        r.raise_for_status(); items=r.json()
        if not isinstance(items,list): return []
        out=[]
        for item in items:
            if str(item.get("type","")).lower()!="file": continue
            name=str(item.get("name","")).strip()
            path=str(item.get("path","")).strip()
            if not name or not path: continue
            out.append({"name":name,"path":path,"sha":str(item.get("sha","")).strip()})
        return out
    except Exception:
        return []


def _daily_xper_date_name(name:str):
    m=re.search(r"(20\d{2})[-_](\d{2})[-_](\d{2}).*Integridade atual de Node", str(name), flags=re.I)
    if not m: return None
    try:
        return datetime(int(m.group(1)),int(m.group(2)),int(m.group(3)),12,0,tzinfo=POA_TZ)
    except Exception:
        return None


def find_best_xpertrack_github():
    """Escolhe e baixa a extração mais recente diretamente do GitHub.

    Prioridade: arquivo diário AAAA-MM-DD - Integridade atual de Node; depois os
    nomes fixos xpertrack_integridade_atual.csv/xlsx. Retorna bytes, nome, path
    e horário do último commit do arquivo.
    """
    files=_gh_list_data_files()
    if not files: return None,None,None,None
    candidates=[]
    for item in files:
        name=item["name"]
        lower=name.lower()
        if not (lower.endswith('.csv') or lower.endswith('.xlsx') or lower.endswith('.xls')):
            continue
        dated=_daily_xper_date_name(name)
        fixed=lower in {'xpertrack_integridade_atual.csv','xpertrack_integridade_atual.xlsx','xpertrack_integridade_atual.xls'}
        if dated is None and not fixed:
            continue
        if dated is not None:
            score=(2,dated.timestamp(),name)
        else:
            upd=github_file_updated_at(item["path"])
            score=(1,upd.timestamp() if upd is not None else 0,name)
        candidates.append((score,item))
    if not candidates: return None,None,None,None
    candidates.sort(key=lambda x:x[0],reverse=True)
    item=candidates[0][1]
    raw=_gh_read_bytes(item["path"])
    if raw is None: return None,None,None,None
    updated=github_file_updated_at(item["path"])
    return raw,item["name"],item["path"],updated


def _gh_write_bytes(path,data:bytes,message:str):
    gh=github_settings()
    if not (gh["repo"] and gh["token"]): return False,"GitHub não configurado."
    try:
        url=f"https://api.github.com/repos/{gh['repo']}/contents/{path}"; headers=github_headers(gh["token"]); getr=requests.get(url,headers=headers,params={"ref":gh["branch"]},timeout=20); sha=getr.json().get("sha") if getr.status_code==200 else None
        if getr.status_code not in (200,404): getr.raise_for_status()
        body={"message":message,"content":base64.b64encode(data).decode("ascii"),"branch":gh["branch"]}
        if sha: body["sha"]=sha
        r=requests.put(url,headers=headers,json=body,timeout=35); r.raise_for_status(); return True,"Salvo no GitHub."
    except Exception as e: return False,f"Falha ao salvar no GitHub: {e}"


def load_csv_store(local_path:Path, gh_path:str, columns:list[str]):
    raw=_gh_read_bytes(gh_path)
    try:
        if raw is not None and raw.strip(): d=pd.read_csv(BytesIO(raw))
        elif local_path.exists() and local_path.stat().st_size: d=pd.read_csv(local_path)
        else: d=pd.DataFrame(columns=columns)
    except Exception: d=pd.DataFrame(columns=columns)
    for c in columns:
        if c not in d.columns: d[c]=""
    return d[columns].copy().fillna("")


def save_csv_store(df:pd.DataFrame, local_path:Path, gh_path:str, columns:list[str], message:str):
    out=df.copy()
    for c in columns:
        if c not in out.columns: out[c]=""
    out=out[columns]; data=out.to_csv(index=False).encode("utf-8-sig"); gh=github_settings()
    if gh["repo"] and gh["token"]: return _gh_write_bytes(gh_path,data,message)
    try:
        local_path.parent.mkdir(parents=True,exist_ok=True); local_path.write_bytes(data); return True,"Salvo localmente."
    except Exception as e: return False,f"Falha ao salvar localmente: {e}"


TREATMENT_COLUMNS=["Node","Tratativa","Situacao","Observacao","Responsavel","Atualizado_em","Status_tecnico_no_registro"]
HISTORY_COLUMNS=["Evento","DataHora","Fuso","Portas_OFF","Portas_OFF_Mapeadas","Portas_OFF_Sem_Localizacao","Nodes_OFF_Total","Nodes_Parcial_Critico","Nodes_Parcial","Nodes_Degradados","Outages_Sem_Sinal","Nodes_Em_Atendimento","Regiao_Mais_Impactada","Pct_Regiao_Crise","Observacao","Tipo_Registro","Chave_Fonte"]
OUTAGES_CURRENT_COLUMNS=["Outages_Sem_Sinal","DataHora","Fuso","Observacao"]
TEAM_COLUMNS=["Data","Tecnico","Equipe","Turno","Regiao_Preferencial","Ativo","Atualizado_em"]
ASSIGN_COLUMNS=["Node","Tecnico","Equipe","Prioridade","Ordem","Status","Situacao","Observacao","Atribuido_em","Atualizado_em","Foto_Path"]


def load_treatments(): return load_csv_store(TREATMENTS_FILE,github_settings()["treatments"],TREATMENT_COLUMNS)
def save_treatments(d): return save_csv_store(d,TREATMENTS_FILE,github_settings()["treatments"],TREATMENT_COLUMNS,f"Atualiza tratativas {now_poa():%Y-%m-%d %H:%M}")
def load_history(): return load_csv_store(HISTORY_FILE,github_settings()["history"],HISTORY_COLUMNS)
def save_history(d): return save_csv_store(d,HISTORY_FILE,github_settings()["history"],HISTORY_COLUMNS,f"Registra snapshot {now_poa():%Y-%m-%d %H:%M}")
def load_outages_current(): return load_csv_store(OUTAGES_CURRENT_FILE,github_settings()["outages_current"],OUTAGES_CURRENT_COLUMNS)
def save_outages_current(d): return save_csv_store(d,OUTAGES_CURRENT_FILE,github_settings()["outages_current"],OUTAGES_CURRENT_COLUMNS,f"Atualiza outages {now_poa():%Y-%m-%d %H:%M}")
def load_team(): return load_csv_store(TEAM_FILE,github_settings()["team"],TEAM_COLUMNS)
def save_team(d): return save_csv_store(d,TEAM_FILE,github_settings()["team"],TEAM_COLUMNS,f"Atualiza equipe do dia {now_poa():%Y-%m-%d %H:%M}")
def load_assignments(): return load_csv_store(ASSIGNMENTS_FILE,github_settings()["assignments"],ASSIGN_COLUMNS)
def save_assignments(d): return save_csv_store(d,ASSIGNMENTS_FILE,github_settings()["assignments"],ASSIGN_COLUMNS,f"Atualiza distribuição de nodes {now_poa():%Y-%m-%d %H:%M}")


def save_photo(upload,node,tech):
    if upload is None: return True,"",""
    data=upload.getvalue()
    if len(data)>6*1024*1024: return False,"","Foto acima de 6 MB."
    ext=Path(upload.name or "foto.jpg").suffix.lower(); ext=ext if ext in {".jpg",".jpeg",".png",".webp"} else ".jpg"
    safe_node=re.sub(r"[^A-Z0-9_-]","_",norm_txt(node)); safe_tech=re.sub(r"[^A-Z0-9_-]","_",norm_txt(tech)); name=f"{now_poa():%Y%m%d_%H%M%S}_{safe_node}_{safe_tech}{ext}"
    gh=github_settings(); path=f"{gh['photos']}/{name}"
    if gh["repo"] and gh["token"]:
        ok,msg=_gh_write_bytes(path,data,f"Foto atendimento {safe_node} {now_poa():%Y-%m-%d %H:%M}")
        return ok,path,msg
    try:
        PHOTO_DIR.mkdir(parents=True,exist_ok=True); (PHOTO_DIR/name).write_bytes(data); return True,f"data/fotos_atendimento/{name}","Foto salva localmente."
    except Exception as e: return False,"",f"Falha ao salvar foto: {e}"


# -----------------------------
# Fonte de dados
# -----------------------------
def _daily_xper_date(path: Path):
    m=re.search(r"(20\d{2})[-_](\d{2})[-_](\d{2}).*Integridade atual de Node", path.name, flags=re.I)
    if not m:
        return None
    try:
        return datetime(int(m.group(1)),int(m.group(2)),int(m.group(3)),12,0,tzinfo=POA_TZ)
    except Exception:
        return None


def _xpertrack_file_candidates():
    candidates=[]
    # Arquivos diários enviados diretamente ao GitHub/Downloads sincronizado.
    for p in DATA_DIR.glob("*Integridade atual de Node*"):
        if p.is_file() and p.suffix.lower() in {".csv",".xlsx",".xls"}:
            candidates.append(p)
    # Nomes fixos continuam compatíveis com a automação local existente.
    for p in (LOCAL_XLSX,LOCAL_CSV):
        if p.exists() and p not in candidates:
            candidates.append(p)
    return candidates


def _candidate_updated_at(path: Path):
    gh_path=f"data/{path.name}"
    gh_dt=github_file_updated_at(gh_path)
    if gh_dt is not None:
        return gh_dt
    try:
        return datetime.fromtimestamp(path.stat().st_mtime,tz=timezone.utc)
    except Exception:
        return None


def find_best_xpertrack_file():
    candidates=_xpertrack_file_candidates()
    if not candidates:
        return None,None
    scored=[]
    for p in candidates:
        upd=_candidate_updated_at(p)
        dated=_daily_xper_date(p)
        # A data explícita do arquivo diário evita que um arquivo antigo de nome fixo
        # prevaleça apenas porque o Streamlit recriou o checkout. Quando existe data
        # de commit do GitHub, ela serve como desempate e horário real da fonte.
        if dated is not None:
            primary=dated.astimezone(timezone.utc).timestamp()
            secondary=upd.timestamp() if upd is not None else 0
            kind=2
        else:
            primary=upd.timestamp() if upd is not None else 0
            secondary=primary
            kind=1
        scored.append((primary,kind,secondary,p,upd))
    scored.sort(key=lambda x:(x[0],x[1],x[2]),reverse=True)
    _,_,_,chosen,upd=scored[0]
    return chosen,upd

base=load_base(); reg_map=load_region_map()
remote_url=get_secret("XPERTRACK_DATA_URL","https://drive.google.com/uc?export=download&id=1khmqXK9bpeTub9eXgluJrkCYm8491sCY").strip(); remote_token=get_secret("XPERTRACK_BEARER_TOKEN","").strip(); source_name=get_secret("DATA_SOURCE_NAME","XPERTrack").strip() or "XPERTrack"
xraw=None; source_updated_at=None; source_file_name=""

# No Streamlit Cloud, lê primeiro a pasta data diretamente do GitHub usando o
# token já configurado. Assim uma nova extração diária passa a valer sem depender
# de o checkout local do app ter sido reconstruído.
gh=github_settings()
if False and gh["repo"] and gh["token"]:
    try:
        gh_raw,gh_name,gh_path,gh_updated=find_best_xpertrack_github()
        if gh_raw is not None and gh_name:
            xraw=load_table_bytes(gh_raw,gh_name)
            source_updated_at=gh_updated
            source_file_name=gh_name
    except Exception as e:
        st.warning(f"Não foi possível ler a extração mais recente diretamente do GitHub: {e}")

# Fonte remota explícita continua disponível como fallback/alternativa.
if xraw is None and remote_url:
    try:
        raw,last_mod=fetch_remote_bytes(remote_url,remote_token); hint=remote_url.split("?")[0].split("/")[-1] or "xpertrack.csv"; xraw=load_table_bytes(raw,hint); source_file_name=hint
        if last_mod:
            try: source_updated_at=parsedate_to_datetime(last_mod)
            except Exception: pass
    except Exception as e: st.error(f"Falha ao consultar a fonte online do XPERTrack: {e}")

# Local permanece como fallback para execução no PC e para ambientes sem token GitHub.
if xraw is None:
    local,detected_updated_at=find_best_xpertrack_file()
    if local:
        try:
            xraw=load_table_bytes(local.read_bytes(),local.name); source_file_name=local.name
            gh_xper_path=get_secret("XPERTRACK_GITHUB_PATH",f"data/{local.name}").strip() or f"data/{local.name}"
            if get_secret("XPERTRACK_GITHUB_PATH","").strip():
                source_updated_at=github_file_updated_at(gh_xper_path) or detected_updated_at
            else:
                source_updated_at=detected_updated_at or github_file_updated_at(f"data/{local.name}")
            if source_updated_at is None:
                source_updated_at=datetime.fromtimestamp(local.stat().st_mtime,tz=timezone.utc)
        except Exception as e: st.error(f"Não consegui abrir {local.name}: {e}")

status_df=pd.DataFrame(); xnorm=pd.DataFrame()
if xraw is not None and not xraw.empty and xpertrack_schema(xraw):
    status_df,xnorm=build_operational_status(xraw)
manual=load_manual_crosswalk()
if not status_df.empty and not manual.empty:
    reg_override={}
    for _,r in manual.iterrows():
        reg=norm_txt(r.get("Regiao")); x=norm_txt(r.get("XPER_Node")); corr=norm_txt(r.get("Node_Correto")) or load_identity_aliases().get(x,x)
        if reg:
            if x: reg_override[x]=reg
            if corr: reg_override[corr]=reg
    status_df["Regiao"]=status_df.apply(lambda r:reg_override.get(norm_txt(r.Node),norm_txt(r.Regiao)),axis=1)

logical_geo=build_logical_geo(status_df,base,reg_map) if not status_df.empty else pd.DataFrame()
mapdf=build_map_frame(logical_geo,base,reg_map)

# Região passa a ser consolidada pelo bairro: um mesmo bairro não pode aparecer
# simultaneamente como NORTE/CENTRAL/SUL. O mapa completo é usado como base
# para escolher a região predominante de cada bairro, e a regra é aplicada
# também aos nodes lógicos do XPERTrack.
bairro_region_map = build_bairro_region_map(mapdf)
mapdf = apply_bairro_region_consistency(mapdf, bairro_region_map)
logical_geo = apply_bairro_region_consistency(logical_geo, bairro_region_map)
# Correções explícitas de região sempre prevalecem sobre a inferência por bairro.
mapdf = apply_node_region_overrides(mapdf)
logical_geo = apply_node_region_overrides(logical_geo)

mapping_ready=not logical_geo.empty

# Tratativas + atribuições
trdf=load_treatments(); treatments_latest=trdf.copy()
if not treatments_latest.empty:
    treatments_latest["Node"]=treatments_latest["Node"].map(norm_txt); treatments_latest=treatments_latest.drop_duplicates("Node",keep="last")
assignments=load_assignments()
if not assignments.empty:
    assignments["Node"]=assignments["Node"].map(norm_txt); assignments["Tecnico"]=assignments["Tecnico"].map(norm_txt); assignments["Status"]=assignments["Status"].map(norm_txt)
    assignments_latest=assignments.drop_duplicates("Node",keep="last").copy()
else: assignments_latest=pd.DataFrame(columns=ASSIGN_COLUMNS)


def overlay_operations(df):
    out=df.copy()
    if not treatments_latest.empty: out=out.merge(treatments_latest,on="Node",how="left")
    else:
        for c in TREATMENT_COLUMNS[1:]: out[c]=""
    for c in TREATMENT_COLUMNS[1:]: out[c]=out[c].fillna("")
    out["Tratativa"]=out["Tratativa"].replace("","SEM TRATATIVA")
    if not assignments_latest.empty:
        a=assignments_latest[["Node","Tecnico","Equipe","Status","Situacao","Observacao","Atualizado_em","Foto_Path"]].copy()
        a=a.rename(columns={"Status":"Status_Atribuicao","Situacao":"Situacao_Atribuicao","Observacao":"Observacao_Atribuicao","Atualizado_em":"Atualizado_Atribuicao"})
        out=out.merge(a,on="Node",how="left")
        for c in ["Tecnico","Equipe","Status_Atribuicao","Situacao_Atribuicao","Observacao_Atribuicao","Atualizado_Atribuicao","Foto_Path"]: out[c]=out[c].fillna("")
        def ap(r):
            s=norm_txt(r.get("Status_Atribuicao"))
            if s in {"NORMALIZADO","CONCLUÍDO","CONCLUIDO"}: return "TRATADO / NORMALIZADO"
            if s and s not in {"CANCELADO","CANCELADA"}: return "EM ATENDIMENTO"
            return r.get("Tratativa","SEM TRATATIVA")
        out["Tratativa"]=out.apply(ap,axis=1)
        out["Situacao"]=out.apply(lambda r:r.get("Situacao_Atribuicao") or r.get("Status_Atribuicao") or r.get("Situacao",""),axis=1)
        out["Responsavel"]=out.apply(lambda r:r.get("Tecnico") or r.get("Responsavel",""),axis=1)
        out["Observacao"]=out.apply(lambda r:r.get("Observacao_Atribuicao") or r.get("Observacao",""),axis=1)
    else:
        for c in ["Tecnico","Equipe","Status_Atribuicao","Situacao_Atribuicao","Observacao_Atribuicao","Atualizado_Atribuicao","Foto_Path"]: out[c]=""
    return out

logical=overlay_operations(logical_geo) if not logical_geo.empty else logical_geo.copy()
mapdf=overlay_operations(mapdf)

def status_color(r):
    """Cores principais do mapa: verde ONLINE, amarelo SS Parcial/porta crítica, vermelho SS Total."""
    s=norm_txt(r.get("Status","SEM DADO"))
    if s=="OFF TOTAL": return [239,51,64,238]
    if s in {"PARCIAL CRITICO","PARCIAL","DEGRADADO"}: return [245,183,0,238]
    if s=="NORMAL": return [16,165,90,232]
    if s=="UNI": return [110,124,146,190]
    return [145,155,170,185]


def line_color(r):
    return [255,255,255,235]

mapdf["color"]=mapdf.apply(status_color,axis=1); mapdf["line_color"]=mapdf.apply(line_color,axis=1); mapdf["Tem_Critica"]=pd.to_numeric(mapdf.get("Portas_Criticas",0),errors="coerce").fillna(0).gt(0); mapdf["Em_Atendimento"]=mapdf["Tratativa"].map(norm_txt).eq("EM ATENDIMENTO"); mapdf["Status_exibicao"]=mapdf["Status"].map(status_publico); mapdf["Regiao_exibicao"]=mapdf["Regiao"].apply(lambda x:x if str(x).strip() else "Não cadastrada"); mapdf["Bairro_exibicao"]=mapdf["Bairro"].apply(lambda x:x if str(x).strip() else "Não identificado")
mapdf["Origem_exibicao"]=mapdf["Origem"].map(lambda x:{"HUB BV":"HUB BELA VISTA","HUB SUL":"HUB ZONA SUL"}.get(norm_txt(x),norm_txt(x) or "Não informada"))
mapdf["Tratativa_txt"]=mapdf.apply(lambda r:("SEM ATENDIMENTO" if norm_txt(r.get("Tratativa"))=="SEM TRATATIVA" else (norm_txt(r.get("Tratativa")) + (f" · {norm_txt(r.get('Situacao'))}" if norm_txt(r.get("Situacao")) else ""))),axis=1)

# -----------------------------
# Métricas e resumos
# -----------------------------
ports_off=int(status_df["Portas_OFF"].fillna(0).sum()) if not status_df.empty else None
ports_critical=int(status_df.get("Portas_Criticas", status_df.get("Portas_Degradadas", pd.Series(dtype=float))).fillna(0).sum()) if not status_df.empty else None
ports_degraded=ports_critical  # compatibilidade com histórico
nodes_degraded=int((status_df["Status"]=="DEGRADADO").sum()) if not status_df.empty else None
off_total=int((status_df["Status"]=="OFF TOTAL").sum()) if not status_df.empty else None
parcial_crit=0  # classificação antiga removida; mantido apenas para compatibilidade do CSV histórico
parcial=int(status_df["Status"].isin(["PARCIAL","PARCIAL CRÍTICO"]).sum()) if not status_df.empty else None
mapped_ports_off=int(logical.loc[logical["Latitude_Base"].notna(),"Portas_OFF"].fillna(0).sum()) if not logical.empty else 0
ports_off_unmapped=max(0,int((ports_off or 0)-mapped_ports_off)) if ports_off is not None else None

def region_summary_df():
    if logical.empty: return pd.DataFrame(columns=["Regiao","Portas_OFF","Pct_crise"])
    d=logical.copy(); d["Regiao"]=d["Regiao"].replace("","SEM REGIÃO"); total=int(d["Portas_OFF"].sum()); rows=[]
    order={"NORTE":0,"CENTRAL":1,"SUL":2,"SEM REGIÃO":9}
    for reg,g in d.groupby("Regiao"):
        off=int(g["Portas_OFF"].sum())
        if off>0: rows.append({"Regiao":reg,"Portas_OFF":off,"Pct_crise":off/total*100 if total else 0,"_ord":order.get(reg,5)})
    out=pd.DataFrame(rows)
    if not out.empty: out=out.sort_values(["_ord","Portas_OFF"],ascending=[True,False]).drop(columns="_ord")
    return out


def neighborhood_summary_df():
    if logical.empty: return pd.DataFrame(columns=["Bairro","Portas_OFF","Pct_crise"])
    d=logical.copy(); d["Bairro"]=d["Bairro"].replace("","SEM BAIRRO IDENTIFICADO"); total=int(d["Portas_OFF"].sum()); rows=[]
    for b,g in d.groupby("Bairro"):
        off=int(g["Portas_OFF"].sum())
        if off>0: rows.append({"Bairro":b,"Portas_OFF":off,"Pct_crise":off/total*100 if total else 0})
    out=pd.DataFrame(rows)
    if not out.empty: out=out.sort_values(["Portas_OFF","Bairro"],ascending=[False,True]).reset_index(drop=True)
    return out

region_summary=region_summary_df(); neighborhood_summary=neighborhood_summary_df()
region_label="Sem portas OFF"; region_pct=0.0
if not region_summary.empty:
    top=region_summary.sort_values("Portas_OFF",ascending=False).iloc[0]; region_label=str(top.Regiao); region_pct=float(top.Pct_crise)
bairro_label="Sem portas OFF"; bairro_off=0
if not neighborhood_summary.empty:
    bt=neighborhood_summary.iloc[0]; bairro_label=str(bt.Bairro); bairro_off=int(bt.Portas_OFF)

em_atendimento=int((logical["Tratativa"].map(norm_txt)=="EM ATENDIMENTO").sum()) if not logical.empty else 0
sem_atendimento=int(((logical["Portas_OFF"]>0)&(logical["Tratativa"].map(norm_txt)=="SEM TRATATIVA")).sum()) if not logical.empty else 0

# Histórico
def _source_fingerprint():
    """Identifica uma coleta XPERTrack sem depender de reruns da página."""
    stamp=""
    if source_updated_at is not None:
        try:
            dt=source_updated_at if source_updated_at.tzinfo else source_updated_at.replace(tzinfo=timezone.utc)
            stamp=dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        except Exception:
            stamp=str(source_updated_at)
    digest=""
    try:
        if xraw is not None and not xraw.empty:
            hv=pd.util.hash_pandas_object(xraw.astype(str),index=True).values.tobytes()
            digest=hashlib.sha1(hv).hexdigest()[:16]
    except Exception:
        pass
    if stamp and digest:
        return f"{stamp}|{digest}"
    return stamp or digest


def _source_local_datetime():
    if source_updated_at is None:
        return now_poa()
    try:
        dt=source_updated_at if source_updated_at.tzinfo else source_updated_at.replace(tzinfo=timezone.utc)
        return dt.astimezone(POA_TZ)
    except Exception:
        return now_poa()


def _snapshot_row(datahora, outages_value, tipo, source_key="", observacao=""):
    return {
        "Evento":datahora.strftime("%Y-%m-%d"),"DataHora":datahora.strftime("%Y-%m-%d %H:%M"),"Fuso":"America/Sao_Paulo",
        "Portas_OFF":int(ports_off or 0),"Portas_OFF_Mapeadas":int(mapped_ports_off),"Portas_OFF_Sem_Localizacao":int(ports_off_unmapped or 0),
        "Nodes_OFF_Total":int(off_total or 0),"Nodes_Parcial_Critico":0,"Nodes_Parcial":int(parcial or 0),"Nodes_Degradados":int(nodes_degraded or 0),
        "Outages_Sem_Sinal":"" if outages_value is None else int(outages_value),"Nodes_Em_Atendimento":int(em_atendimento),
        "Regiao_Mais_Impactada":region_label,"Pct_Regiao_Crise":round(region_pct,2),"Observacao":observacao,
        "Tipo_Registro":tipo,"Chave_Fonte":source_key,
    }


def _history_datetime_local(row):
    """Converte o horário histórico para Porto Alegre sem deslocar registros novos.

    Registros novos trazem Fuso=America/Sao_Paulo. Em registros legados sem
    marcador, corrigimos automaticamente os casos claramente gravados em UTC
    (por exemplo, horário no dia seguinte ao Evento ou no futuro local).
    """
    raw=pd.to_datetime(row.get("DataHora"),errors="coerce",dayfirst=True)
    if pd.isna(raw): return pd.NaT
    ts=pd.Timestamp(raw)
    fuso=norm_txt(row.get("Fuso",""))
    if "SAO_PAULO" in fuso or "AMERICA/SAO_PAULO" in str(row.get("Fuso","")).upper():
        return ts.tz_localize(None) if ts.tzinfo else ts
    if ts.tzinfo is not None:
        try: return ts.tz_convert(POA_TZ).tz_localize(None)
        except Exception: return ts.tz_localize(None)
    event_date=pd.to_datetime(row.get("Evento"),errors="coerce")
    now_local=pd.Timestamp(now_poa().replace(tzinfo=None))
    looks_utc=False
    if pd.notna(event_date) and ts.date()>event_date.date(): looks_utc=True
    if ts>now_local+pd.Timedelta(minutes=10): looks_utc=True
    if looks_utc:
        try: return ts.tz_localize("UTC").tz_convert(POA_TZ).tz_localize(None)
        except Exception: return ts-pd.Timedelta(hours=3)
    return ts

# Valor atual de outages é persistido separadamente do histórico.
outages_current=None
outages_current_df=load_outages_current()
if not outages_current_df.empty:
    vals=pd.to_numeric(outages_current_df["Outages_Sem_Sinal"],errors="coerce").dropna()
    if len(vals): outages_current=int(vals.iloc[-1])

history_df=load_history()
source_key=_source_fingerprint()
# Cada coleta real do XPERTrack deve existir uma única vez no histórico.
# Se a mesma coleta já estiver salva com valores antigos/incorretos, fazemos UPSERT
# do snapshot atual em vez de ignorá-la. Rerun da página não cria um novo ponto.
if mapping_ready and source_key:
    source_local=_source_local_datetime()
    snap=_snapshot_row(source_local,outages_current,"XPERTRACK",source_key,"")
    hd=history_df.copy() if not history_df.empty else pd.DataFrame(columns=HISTORY_COLUMNS)
    for c in HISTORY_COLUMNS:
        if c not in hd.columns: hd[c]=""
    key_series=hd.get("Chave_Fonte",pd.Series(index=hd.index,dtype=str)).fillna("").astype(str)
    matches=hd.index[key_series.eq(str(source_key))].tolist()
    changed=False
    if matches:
        target=matches[-1]
        before=hd.loc[target,HISTORY_COLUMNS].astype(str).to_dict()
        for c in HISTORY_COLUMNS:
            hd.at[target,c]=snap.get(c,"")
        after=hd.loc[target,HISTORY_COLUMNS].astype(str).to_dict()
        changed=before!=after
    else:
        hd=pd.concat([hd,pd.DataFrame([snap])],ignore_index=True)
        changed=True
    if changed:
        save_history(hd)
    history_df=hd

if not history_df.empty:
    history_df["DataHora_dt"]=history_df.apply(_history_datetime_local,axis=1)
    history_df=history_df.sort_values("DataHora_dt",na_position="last").reset_index(drop=True)
    today=now_poa().date()
    local_dates=history_df["DataHora_dt"].map(lambda v:v.date() if pd.notna(v) else None)
    event_history=history_df[local_dates.eq(today)].copy()
    current_event=today.isoformat()
    # Limpa legado do próprio dia: registros sem tipo eram gerados por reruns.
    # Mantemos apenas mudanças efetivas de portas OFF ou outages.
    if not event_history.empty:
        event_history["_off_num"]=pd.to_numeric(event_history["Portas_OFF"],errors="coerce")
        event_history["_out_num"]=pd.to_numeric(event_history["Outages_Sem_Sinal"],errors="coerce")
        tipo=event_history.get("Tipo_Registro",pd.Series(index=event_history.index,dtype=str)).fillna("").map(norm_txt)
        explicit=tipo.isin(["XPERTRACK","OUTAGES","MANUAL"])
        off_cmp=event_history["_off_num"].fillna(-10**12); out_cmp=event_history["_out_num"].fillna(-10**12)
        legacy_changed=(off_cmp.ne(off_cmp.shift()) | out_cmp.ne(out_cmp.shift()))
        keep=explicit | legacy_changed
        if len(keep): keep.iloc[0]=True
        event_history=event_history.loc[keep].drop(columns=["_off_num","_out_num"],errors="ignore").reset_index(drop=True)
else:
    current_event=now_poa().date().isoformat(); event_history=pd.DataFrame(columns=HISTORY_COLUMNS+["DataHora_dt"])

# Fallback apenas para históricos antigos, antes da criação do arquivo outages_atual.csv.
if outages_current is None and not event_history.empty:
    outs=pd.to_numeric(event_history["Outages_Sem_Sinal"],errors="coerce").dropna()
    outages_current=int(outs.iloc[-1]) if len(outs) else None

def _num(v):
    try:
        if pd.isna(v) or str(v).strip()=="": return None
        return float(v)
    except Exception: return None

# Para portas OFF, a comparação é sempre entre coletas XPERTrack, nunca contra um salvamento manual de outages.
if not event_history.empty:
    tipos=event_history.get("Tipo_Registro",pd.Series(index=event_history.index,dtype=str)).fillna("").map(norm_txt)
    xper_hist=event_history[tipos.eq("XPERTRACK")].copy()
    if xper_hist.empty:
        # Legado do dia: usa somente registros que contenham valor técnico.
        xper_hist=event_history[pd.to_numeric(event_history.get("Portas_OFF"),errors="coerce").notna()].copy()
    if not xper_hist.empty:
        xper_hist=xper_hist.sort_values("DataHora_dt",na_position="last")
        # Hotfix de migração: a V9.x podia ter salvo duas versões da mesma coleta/minuto.
        # Mantemos a última (já corrigida pelo UPSERT acima).
        xper_hist=xper_hist.drop_duplicates(subset=["DataHora_dt"],keep="last").reset_index(drop=True)
else:
    xper_hist=pd.DataFrame(columns=event_history.columns)

baseline_off=_num(xper_hist.iloc[0].get("Portas_OFF")) if not xper_hist.empty else None
if len(xper_hist)>=2:
    previous_xper_off=_num(xper_hist.iloc[-2].get("Portas_OFF"))
elif len(xper_hist)==1:
    previous_xper_off=_num(xper_hist.iloc[-1].get("Portas_OFF"))
else:
    previous_xper_off=None
baseline_off=float(ports_off or 0) if baseline_off is None else baseline_off
previous_xper_off=float(ports_off or 0) if previous_xper_off is None else previous_xper_off
histvals=pd.to_numeric(event_history.get("Portas_OFF",pd.Series(dtype=float)),errors="coerce").dropna().tolist(); peak_off=int(max(histvals+[ports_off or 0])) if histvals else int(ports_off or 0)
delta_baseline=int((ports_off or 0)-baseline_off); delta_last=int((ports_off or 0)-previous_xper_off); recovered=int(max(0,baseline_off-(ports_off or 0))); recovery_pct=recovered/baseline_off*100 if baseline_off else 0
if len(xper_hist)<=1 and delta_last==0: crisis_trend="ESTÁVEL"
elif delta_last<0: crisis_trend="NORMALIZANDO"
elif delta_last>0: crisis_trend="PIORANDO"
else: crisis_trend="ESTÁVEL"
crisis_mode=bool((ports_off or 0)>0)

# -----------------------------
# Acesso / cabeçalho
# -----------------------------
def source_age_minutes(dt):
    if dt is None: return None
    try:
        now=datetime.now(timezone.utc); dt=dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc); return max(0,(now-dt.astimezone(timezone.utc)).total_seconds()/60)
    except Exception: return None
age_min=source_age_minutes(source_updated_at); stale_after=int(CFG.get("stale_after_minutes",90) or 90); is_stale=age_min is not None and age_min>stale_after
if source_updated_at:
    try: updated_txt=source_updated_at.astimezone(POA_TZ).strftime("%d/%m/%Y %H:%M")
    except Exception: updated_txt=source_updated_at.strftime("%d/%m/%Y %H:%M")
else: updated_txt="sem horário da fonte"

header_warn = f' • sem atualização há <b>{int(age_min)} min</b>' if is_stale else ''
st.markdown(
    f'<div class="topbar">'
    f'<div class="title-wrap">'
    f'<h1>Painel Geográfico de Nodes – Rio Grande</h1>'
    f'<div class="sub">XPERTrack • Pontuação por porta • RG V3</div>'
    f'<div class="header-meta">Fonte XPERTrack: {esc(updated_txt)}{header_warn}' + (f' • arquivo: {esc(source_file_name)}' if source_file_name else '') + '</div>'
    f'</div></div>',
    unsafe_allow_html=True,
)

# V9.1: por enquanto o painel exibe somente a visão principal.
# As visões Supervisor e Técnico permanecem preservadas no código para uma etapa futura, mas não aparecem na interface.
view="Executiva"


def check_admin_access():
    if request_is_local(): return True
    pin=get_secret("ADMIN_PIN","").strip()
    if not pin:
        st.info("A Visão Supervisor está protegida. Configure ADMIN_PIN nos Secrets do Streamlit."); return False
    entered=st.text_input("PIN do supervisor",type="password",key="admin_view_pin")
    ok=bool(entered) and hmac.compare_digest(str(entered),pin)
    if entered and not ok: st.error("PIN inválido.")
    return ok


def check_operational_access():
    if request_is_local(): return True
    pin=get_secret("ADMIN_PIN","").strip()
    if not pin:
        st.info("Para editar informações online, configure ADMIN_PIN nos Secrets do Streamlit.")
        return False
    entered=st.text_input("PIN de atualização",type="password",key="operational_update_pin")
    ok=bool(entered) and hmac.compare_digest(str(entered),pin)
    if entered and not ok: st.error("PIN inválido.")
    return ok


def _close_operational_dialog():
    try:
        st.query_params.clear()
    except Exception:
        pass


@st.dialog("Atualização operacional", width="large")
def render_operational_update_dialog():
    st.caption("Atualização manual protegida pelo ADMIN_PIN. O status técnico continua vindo do XPERTrack.")
    if not check_operational_access():
        return

    # 1) Outages primeiro, conforme fluxo operacional validado.
    st.markdown("#### 1. Outages sem sinal")
    with st.form("outages_operational_form_dialog"):
        outages_input=st.number_input("Quantidade atual",min_value=0,step=1,value=int(outages_current or 0))
        obs_out=st.text_input("Observação (opcional)",key="outages_note_dialog")
        save_out=st.form_submit_button("💾 Registrar outages",use_container_width=True)
    if save_out:
        now=now_poa()
        current_row=pd.DataFrame([{"Outages_Sem_Sinal":int(outages_input),"DataHora":now.strftime("%Y-%m-%d %H:%M"),"Fuso":"America/Sao_Paulo","Observacao":obs_out.strip()}])
        ok_current,msg_current=save_outages_current(current_row)
        if not ok_current:
            st.error(msg_current)
        else:
            row=_snapshot_row(now,int(outages_input),"OUTAGES",source_key,obs_out.strip())
            hd=history_df.drop(columns=["DataHora_dt"],errors="ignore") if not history_df.empty else pd.DataFrame(columns=HISTORY_COLUMNS)
            hd=pd.concat([hd,pd.DataFrame([row])],ignore_index=True)
            ok_hist,msg_hist=save_history(hd)
            st.session_state["operational_flash"]=("success" if ok_hist else "warning", f"Outages atualizado para {int(outages_input)}. {msg_current}" + ("" if ok_hist else f" Histórico: {msg_hist}"))
            _close_operational_dialog(); st.rerun()

    st.divider()
    # 2) Node depois de outages.
    st.markdown("#### 2. Informação do node")
    node_options=sorted([n for n in logical.get("Node",pd.Series(dtype=str)).dropna().map(norm_txt).unique().tolist() if n and n not in RETIRED_NODES])
    if not node_options:
        node_options=sorted([n for n in mapdf.get("Node",pd.Series(dtype=str)).dropna().map(norm_txt).unique().tolist() if n and n not in RETIRED_NODES])
    selected_node=st.selectbox("Node",node_options,key="op_update_node_dialog") if node_options else None
    current_rec=None
    if selected_node and not treatments_latest.empty:
        tmp=treatments_latest[treatments_latest["Node"].map(norm_txt)==norm_txt(selected_node)]
        if not tmp.empty: current_rec=tmp.iloc[-1]
    treatment_opts=["SEM TRATATIVA","EM ATENDIMENTO","TRATADO / NORMALIZADO"]
    current_t=norm_txt(current_rec.get("Tratativa")) if current_rec is not None else "SEM TRATATIVA"
    t_idx=treatment_opts.index(current_t) if current_t in treatment_opts else 0
    with st.form("node_operational_form_dialog"):
        trat=st.selectbox("Status do atendimento",treatment_opts,index=t_idx)
        situacao=st.text_input("Situação",value=str(current_rec.get("Situacao","") if current_rec is not None else ""),placeholder="Ex.: técnico atuando, aguardando energia, gerador acionado")
        responsavel=st.text_input("Equipe / responsável",value=str(current_rec.get("Responsavel","") if current_rec is not None else ""))
        observacao=st.text_area("Observação",value=str(current_rec.get("Observacao","") if current_rec is not None else ""),height=80)
        save_node=st.form_submit_button("💾 Salvar informação do node",use_container_width=True,disabled=not selected_node)
    if save_node and selected_node:
        all_t=load_treatments(); now=now_poa().strftime("%d/%m/%Y %H:%M")
        new_row={"Node":norm_txt(selected_node),"Tratativa":norm_txt(trat),"Situacao":situacao.strip(),"Observacao":observacao.strip(),"Responsavel":responsavel.strip(),"Atualizado_em":now,"Status_tecnico_no_registro":norm_txt(trat)}
        all_t=pd.concat([all_t,pd.DataFrame([new_row])],ignore_index=True)
        ok,msg=save_treatments(all_t)
        if ok:
            st.success(msg); _close_operational_dialog(); st.rerun()
        else: st.error(msg)

    if st.button("Fechar",use_container_width=True,key="close_operational_dialog"):
        _close_operational_dialog(); st.rerun()


# O link no card de Outages abre este diálogo sem ocupar espaço no corpo do painel.
if str(st.query_params.get("atualizar", "")).lower() in {"1","sim","outages","operacional"}:
    render_operational_update_dialog()

_flash=st.session_state.pop("operational_flash",None)
if _flash:
    kind,msg=_flash
    (st.warning if kind=="warning" else st.success)(msg)

def tech_users():
    raw=get_secret("TECH_USERS_JSON","").strip()
    if not raw: return {}
    try:
        obj=json.loads(raw); out={}
        for user,v in obj.items():
            if isinstance(v,dict) and v.get("pin") and v.get("tecnico"): out[norm_txt(user)]={"pin":str(v["pin"]),"tecnico":norm_txt(v["tecnico"])}
        return out
    except Exception: return {}


def check_tech_access():
    if request_is_local():
        team=load_team(); today=now_poa().strftime("%Y-%m-%d"); t=team[(team["Data"].astype(str)==today)&(team["Ativo"].astype(str).str.upper().isin(["TRUE","1","SIM","YES"]))]
        choices=sorted(t["Tecnico"].map(norm_txt).dropna().unique().tolist())
        if not choices: st.info("Cadastre a Equipe do dia na Visão Supervisor."); return None
        return st.selectbox("Técnico (teste local)",choices)
    users=tech_users()
    if not users:
        st.info("A Visão Técnico está protegida. Configure TECH_USERS_JSON nos Secrets do Streamlit."); return None
    c1,c2=st.columns(2); user=norm_txt(c1.text_input("Usuário do técnico")); pin=c2.text_input("PIN",type="password")
    if not user or not pin: return None
    rec=users.get(user)
    if not rec or not hmac.compare_digest(pin,rec["pin"]): st.error("Usuário ou PIN inválido."); return None
    return rec["tecnico"]


# -----------------------------
# Componentes de visão
# -----------------------------
def render_kpis():
    off_sub=(f"{mapped_ports_off} localizadas" + (f" • {ports_off_unmapped} sem localização" if ports_off_unmapped else "")) if ports_off is not None else "aguardando XPERTrack"
    outages_action='<a class="kpi-action" href="?atualizar=operacional" target="_self">Atualizar</a>'
    cards=[
        kpi_html("Portas OFF",fmt_int(ports_off),off_sub,"kpi-red"),
        kpi_html("Outages sem sinal",fmt_int(outages_current) if outages_current is not None else "—",outages_action,"kpi-blue",sub_is_html=True),
        kpi_html("Sem sinal total",fmt_int(off_total),"todas as portas existentes OFF","kpi-red"),
        kpi_html("Sem sinal parcial",fmt_int(parcial),"1 ou mais portas OFF, sem perda total","kpi-yellow"),
        kpi_html("Portas críticas",fmt_int(ports_critical),"pontuação de 1 a 20","kpi-yellow"),
        kpi_html("Região mais impactada",region_label,f"{region_pct:.1f}% das portas OFF","kpi-blue"),
        kpi_html("Bairro mais impactado",bairro_label,f"{bairro_off} portas OFF","kpi-blue"),
    ]
    st.markdown('<div class="kpi-grid">'+''.join(cards)+'</div>',unsafe_allow_html=True)


def render_evolution():
    def dt(v): return f"+{v}" if v>0 else str(v)
    base_time="--"
    if not xper_hist.empty and pd.notna(xper_hist.iloc[0].get("DataHora_dt")):
        try: base_time=xper_hist.iloc[0]["DataHora_dt"].strftime("%Hh%M")
        except Exception: pass
    cards=[
        mini_html(f"Desde {base_time}",dt(delta_baseline),f"{recovered} portas recuperadas • {recovery_pct:.1f}%"),
        mini_html("Vs. última atualização",dt(delta_last),"negativo = melhora"),
        mini_html("Pico do dia",fmt_int(peak_off),"maior volume registrado"),
        mini_html("Tendência",crisis_trend,"leitura das portas zeradas"),
    ]
    st.markdown('<div class="mini-grid">'+''.join(cards)+'</div>',unsafe_allow_html=True)

    st.markdown('<div class="panel-title"><b>📈 Evolução dos eventos</b><span>somente coletas reais do dia</span></div>',unsafe_allow_html=True)
    # Uma única coleta não representa evolução. O bloco permanece informativo, mas
    # o gráfico só aparece quando existe uma segunda coleta real do XPERTrack.
    if len(xper_hist)<2:
        st.info("Aguardando próxima coleta para formar a evolução dos eventos.")
        return

    tech=xper_hist.copy()
    tech["DataHora_dt"]=tech.apply(_history_datetime_local,axis=1)
    tech["Valor"]=pd.to_numeric(tech["Portas_OFF"],errors="coerce")
    tech=tech.dropna(subset=["DataHora_dt","Valor"]).sort_values("DataHora_dt")
    tech["Hora"]=tech["DataHora_dt"].dt.strftime("%H:%M")
    tech["Rotulo"]=tech["Valor"].round(0).astype(int).astype(str)

    if tech.empty:
        st.info("Aguardando dados válidos para formar a evolução dos eventos.")
        return

    # Reserva espaço para os rótulos sobre os pontos.
    vmax=float(tech["Valor"].max() or 0)
    ymax=max(5.0, math.ceil(vmax*1.18))

    xenc=alt.X(
        "DataHora_dt:T",
        title=None,
        axis=alt.Axis(format="%H:%M",labelAngle=0,labelOverlap=True,labelFontSize=10,tickCount=min(len(tech),12)),
    )
    yenc=alt.Y(
        "Valor:Q",
        title=None,
        scale=alt.Scale(domain=[0,ymax],nice=False),
        axis=alt.Axis(tickMinStep=1,format="d"),
    )

    # Série principal: Portas OFF.
    base_tech=alt.Chart(tech).encode(
        x=xenc,
        y=yenc,
        tooltip=[
            alt.Tooltip("Hora:N",title="Hora"),
            alt.Tooltip("Valor:Q",title="Portas OFF",format=".0f"),
        ],
    )
    tech_line=base_tech.mark_line(strokeWidth=2.6,color="#0c66d6")
    tech_points=base_tech.mark_point(filled=True,size=70,color="#0c66d6")
    tech_labels=alt.Chart(tech).mark_text(
        dy=-17,
        fontSize=12,
        fontWeight="bold",
        color="#0b1736",
        baseline="bottom",
    ).encode(
        x=xenc,
        y=yenc,
        text=alt.Text("Rotulo:N"),
    )

    layers=[tech_line,tech_points,tech_labels]

    # Outages só aparece quando houve atualização manual salva.
    if not event_history.empty:
        tipos=event_history.get("Tipo_Registro",pd.Series(index=event_history.index,dtype=str)).fillna("").map(norm_txt)
        od=event_history[tipos.isin(["OUTAGES","MANUAL"])].copy()
        if not od.empty:
            od["DataHora_dt"]=od.apply(_history_datetime_local,axis=1)
            od["Valor"]=pd.to_numeric(od["Outages_Sem_Sinal"],errors="coerce")
            od=od.dropna(subset=["DataHora_dt","Valor"]).sort_values("DataHora_dt")
            if not od.empty:
                od["Hora"]=od["DataHora_dt"].dt.strftime("%H:%M")
                od["Rotulo"]=od["Valor"].round(0).astype(int).astype(str)
                out_base=alt.Chart(od).encode(
                    x=xenc,
                    y=yenc,
                    tooltip=[
                        alt.Tooltip("Hora:N",title="Hora"),
                        alt.Tooltip("Valor:Q",title="Outages sem sinal",format=".0f"),
                    ],
                )
                layers.extend([
                    out_base.mark_line(strokeWidth=1.8,color="#7ab8ff",strokeDash=[5,4]),
                    out_base.mark_point(filled=True,size=55,color="#7ab8ff"),
                    alt.Chart(od).mark_text(
                        dy=16,fontSize=10,fontWeight="bold",color="#0c66d6",baseline="top"
                    ).encode(x=xenc,y=yenc,text=alt.Text("Rotulo:N")),
                ])

    chart=alt.layer(*layers).properties(height=260).configure_view(strokeWidth=0)
    st.altair_chart(chart,use_container_width=True)



def _capture_map_selection(map_key: str):
    """Ao tocar/clicar num node, usa a seleção do PyDeck para centralizar o mapa."""
    try:
        state=st.session_state.get(map_key)
        selection=getattr(state,"selection",None)
        if selection is None and isinstance(state,dict):
            selection=state.get("selection",{})
        objects=getattr(selection,"objects",None) if selection is not None else None
        if objects is None and isinstance(selection,dict):
            objects=selection.get("objects",{})
        selected=[]
        if isinstance(objects,dict):
            selected=objects.get("nodes",[]) or []
        else:
            try:
                selected=objects.get("nodes",[]) or []
            except Exception:
                selected=[]
        if selected:
            node=norm_txt(selected[0].get("Node",""))
            if node and node not in RETIRED_NODES:
                st.session_state["map_focus_node"]=node
                st.session_state["node_locator"]=node
    except Exception:
        pass


def render_map_search():
    """Busca apenas navega no mapa; não filtra nem altera os indicadores."""
    active=sorted([n for n in mapdf.get("Node",pd.Series(dtype=str)).dropna().map(norm_txt).unique().tolist() if n and n not in RETIRED_NODES])
    active_set=set(active)
    # Permite pesquisar aliases explicitamente validados, sem fundir nodes por semelhança.
    aliases={k:v for k,v in load_identity_aliases().items() if k not in RETIRED_NODES and v in active_set}
    options=sorted(set(active) | set(aliases.keys()))
    if not options:
        return
    current=norm_txt(st.session_state.get("map_focus_node",""))
    idx=(options.index(current)+1) if current in options else 0
    chosen=st.selectbox(
        "🔎 Localizar node",
        [""]+options,
        index=idx,
        format_func=lambda x:"Digite ou selecione o node..." if not x else (f"{x} → {aliases[x]}" if x in aliases else x),
        key="node_locator",
        help="A busca apenas centraliza e destaca o node no mapa; não altera os indicadores.",
    )
    if chosen:
        st.session_state["map_focus_node"]=norm_txt(aliases.get(norm_txt(chosen),norm_txt(chosen)))
    elif current and st.session_state.get("node_locator","")=="":
        st.session_state.pop("map_focus_node",None)


def _focus_card(row):
    node=esc(row.get("Node","")); status=esc(row.get("Status_exibicao","")); origem=esc(row.get("Origem_exibicao",row.get("Origem","")))
    reg=esc(row.get("Regiao_exibicao","")); bairro=esc(row.get("Bairro_exibicao","")); ports=esc(row.get("Portas_popup","")); trat=esc(row.get("Tratativa_txt",""))
    trat_html=("<div class='map-focus-trat'>Atendimento: "+trat+"</div>") if trat else ""
    lat=row.get("Latitude_Base",row.get("Latitude")); lon=row.get("Longitude_Base",row.get("Longitude"))
    route_html=""
    try:
        lat=float(lat); lon=float(lon)
        if math.isfinite(lat) and math.isfinite(lon):
            url=f"https://www.google.com/maps/dir/?api=1&destination={lat:.7f},{lon:.7f}&travelmode=driving"
            precision=norm_txt(row.get("Precisao_Local",""))
            note=""
            if precision and precision!="KMZ_EXATO":
                note="<div class='map-focus-trat'>⚠️ Localização vinculada/corrigida no cadastro. Confirme o ponto antes de iniciar a rota.</div>"
            route_html=note+f"<a class='kpi-action' href='{url}' target='_blank' rel='noopener noreferrer'>🧭 Traçar rota no Google Maps</a>"
    except Exception:
        pass
    card=("<div class='map-focus-card'><div class='map-focus-name'>📍 "+node+"</div>"
          "<div class='map-focus-meta'><b>"+status+"</b> • "+origem+" • "+reg+" • "+bairro+"</div>"
          "<div class='map-focus-ports'>"+ports+"</div>"+trat_html+route_html+"</div>")
    st.markdown(card,unsafe_allow_html=True)


def render_map(df=None,height=535,assigned_only=False):
    d=(df.copy() if df is not None else mapdf.copy()); d=d[d["Latitude"].notna()&d["Longitude"].notna()].copy()
    if "Status_exibicao" not in d.columns: d["Status_exibicao"]=d.get("Status","").map(status_publico) if isinstance(d.get("Status"),pd.Series) else ""
    if "Regiao_exibicao" not in d.columns: d["Regiao_exibicao"]=d.get("Regiao","").map(lambda x:x if str(x).strip() else "Não cadastrada") if isinstance(d.get("Regiao"),pd.Series) else "Não cadastrada"
    if "Bairro_exibicao" not in d.columns: d["Bairro_exibicao"]=d.get("Bairro","").map(lambda x:x if str(x).strip() else "Não identificado") if isinstance(d.get("Bairro"),pd.Series) else "Não identificado"
    if "Origem_exibicao" not in d.columns: d["Origem_exibicao"]=d.get("Origem","").map(lambda x:{"HUB BV":"HUB BELA VISTA","HUB SUL":"HUB ZONA SUL"}.get(norm_txt(x),norm_txt(x) or "Não informada")) if isinstance(d.get("Origem"),pd.Series) else "Não informada"
    if "Tratativa_txt" not in d.columns: d["Tratativa_txt"]="SEM ATENDIMENTO"
    if d.empty:
        st.info("Sem pontos localizados para exibir.")
        return

    focus_node=norm_txt(st.session_state.get("map_focus_node",""))
    focus=d[d["Node"].map(norm_txt)==focus_node].copy() if focus_node else pd.DataFrame()
    if not focus.empty:
        fr=focus.iloc[0]
        center_lat=float(fr["Latitude"]); center_lon=float(fr["Longitude"]); zoom=12.7 if not assigned_only else 13.2
        _focus_card(fr)
    else:
        center_lat=float(d["Latitude"].mean()); center_lon=float(d["Longitude"].mean()); zoom=11.0 if assigned_only else 10.65

    layers=[]; affected=d[d["Portas_OFF"]>0].copy()
    if crisis_mode and not affected.empty:
        affected["weight"]=affected["Portas_OFF"].clip(lower=1)+affected["Status"].map({"OFF TOTAL":3,"PARCIAL CRÍTICO":2,"PARCIAL":2}).fillna(1)
        layers.append(pdk.Layer("HeatmapLayer",data=affected,id="off-heat",get_position="[Longitude, Latitude]",get_weight="weight",radius_pixels=45,intensity=1.0,threshold=.03,opacity=.35,pickable=False))

    layers.append(pdk.Layer("ScatterplotLayer",data=d,id="nodes",get_position="[Longitude, Latitude]",get_radius=58 if crisis_mode else 48,radius_min_pixels=6,radius_max_pixels=14,get_fill_color="color",get_line_color="line_color",line_width_min_pixels=2.2,stroked=True,pickable=True,auto_highlight=True))

    # Porta crítica: mantém o amarelo e recebe um anel dourado mais escuro.
    critical_values=pd.to_numeric(d["Portas_Criticas"],errors="coerce").fillna(0) if "Portas_Criticas" in d.columns else pd.Series(0,index=d.index,dtype=float)
    critical=d[critical_values>0].copy()
    if not critical.empty:
        critical["critical_line"]=[[184,135,0,255]]*len(critical)
        layers.append(pdk.Layer("ScatterplotLayer",data=critical,id="critical-ring",get_position="[Longitude, Latitude]",get_radius=76,radius_min_pixels=9,radius_max_pixels=17,filled=False,stroked=True,get_line_color="critical_line",line_width_min_pixels=3,pickable=False))

    # Atendimento é uma camada independente do status técnico: contorno azul por fora da cor principal.
    attending=d[d.get("Tratativa",pd.Series(index=d.index,dtype=str)).map(norm_txt)=="EM ATENDIMENTO"].copy()
    if not attending.empty:
        attending["attend_line"]=[[0,145,230,255]]*len(attending)
        layers.append(pdk.Layer("ScatterplotLayer",data=attending,id="attendance-ring",get_position="[Longitude, Latitude]",get_radius=92,radius_min_pixels=11,radius_max_pixels=20,filled=False,stroked=True,get_line_color="attend_line",line_width_min_pixels=4,pickable=False))

    if not focus.empty:
        ring=focus.copy(); ring["focus_line"]=[[11,102,214,255]]*len(ring)
        layers.append(pdk.Layer("ScatterplotLayer",data=ring,id="focus-ring",get_position="[Longitude, Latitude]",get_radius=112,radius_min_pixels=15,radius_max_pixels=25,filled=False,stroked=True,get_line_color="focus_line",line_width_min_pixels=4,pickable=False))

    tooltip={"html":"<div style='font-size:13px;line-height:1.5;min-width:230px'><div style='font-size:18px;font-weight:800;margin-bottom:5px'>{Node}</div><div><b>Status:</b> {Status_exibicao}</div><div><b>Origem:</b> {Origem_exibicao}</div><div><b>Região:</b> {Regiao_exibicao}</div><div><b>Bairro:</b> {Bairro_exibicao}</div><div style='border-top:1px solid rgba(255,255,255,.2);margin:7px 0 5px'></div><div style='font-weight:700'>Portas</div><div style='white-space:pre-line'>{Portas_popup}</div><div style='border-top:1px solid rgba(255,255,255,.2);margin:7px 0 5px'></div><div><b>Atendimento:</b> {Tratativa_txt}</div></div>","style":{"backgroundColor":"rgba(9,26,51,.97)","color":"white","borderRadius":"10px","padding":"10px 12px"}}
    view=pdk.ViewState(latitude=center_lat,longitude=center_lon,zoom=zoom,pitch=0)
    map_key="nodes_map_assigned" if assigned_only else "nodes_map_main"
    deck=pdk.Deck(layers=layers,initial_view_state=view,tooltip=tooltip,map_style=LIGHT_MAP_STYLE)  # CARTO Voyager: ruas, parques e água mais visíveis
    st.pydeck_chart(deck,use_container_width=True,height=height,on_select=lambda:_capture_map_selection(map_key),selection_mode="single-object",key=map_key)

def render_region_table():
    st.markdown('<div class="panel-title"><b>📊 Portas OFF por região</b></div>',unsafe_allow_html=True)
    if region_summary.empty: st.success("Nenhuma porta OFF na leitura atual."); return
    rows=[]; total=0
    for r in region_summary.itertuples():
        total+=int(r.Portas_OFF); rows.append(f"<tr><td><b>{esc(r.Regiao)}</b></td><td>{int(r.Portas_OFF)}</td><td><b>{r.Pct_crise:.1f}%</b></td></tr>")
    rows.append(f"<tr><td><b>TOTAL</b></td><td><b>{total}</b></td><td><b>100,0%</b></td></tr>")
    st.markdown(f'<div class="panel"><table class="rank"><tr><th>Região</th><th>Portas OFF</th><th>% da crise</th></tr>{"".join(rows)}</table></div>',unsafe_allow_html=True)


def render_neighborhood_table(limit=6):
    st.markdown('<div class="panel-title"><b>🏘️ Bairros mais impactados</b><span>portas OFF</span></div>',unsafe_allow_html=True)
    if neighborhood_summary.empty: st.success("Nenhuma porta OFF na leitura atual."); return
    rows=[]
    for r in neighborhood_summary.head(limit).itertuples(): rows.append(f"<tr><td><b>{esc(r.Bairro)}</b></td><td>{int(r.Portas_OFF)}</td><td><b>{r.Pct_crise:.1f}%</b></td></tr>")
    st.markdown(f'<div class="panel"><table class="rank"><tr><th>Bairro</th><th>Portas OFF</th><th>% da crise</th></tr>{"".join(rows)}</table></div>',unsafe_allow_html=True)


def render_top_nodes(show_treatment=False,limit=10):
    st.markdown('<div class="panel-title"><b>⚠️ Top Nodes críticos</b><span>sem sinal</span></div>',unsafe_allow_html=True)
    crit=logical[logical["Portas_OFF"]>0].copy() if not logical.empty else pd.DataFrame()
    if crit.empty: st.info("Nenhum node com porta OFF na leitura atual."); return
    sev={"OFF TOTAL":3,"PARCIAL CRÍTICO":2,"PARCIAL":2}; crit["_sev"]=crit["Status"].map(sev).fillna(0); crit=crit.sort_values(["_sev","Portas_OFF"],ascending=False).head(limit); rows=[]
    for i,r in enumerate(crit.itertuples()):
        cells=f"<td>{i+1}</td><td><b>{esc(r.Node)}</b></td><td>{int(r.Portas_OFF)} / {int(r.Total_portas) if r.Total_portas else '—'}</td><td>{badge(r.Status)}</td>"
        if show_treatment: cells+=f"<td>{treatment_badge(r.Tratativa)}</td>"
        rows.append(f"<tr>{cells}</tr>")
    head="<th>#</th><th>Node</th><th>Portas</th><th>Status</th>"+("<th>Atendimento</th>" if show_treatment else "")
    st.markdown(f'<div class="panel"><table class="rank"><tr>{head}</tr>{"".join(rows)}</table></div>',unsafe_allow_html=True)


def render_legend():
    st.markdown('''<div class="legend-card"><span class="legend-title">Legenda</span><span class="legend-line"><span class="sw" style="background:#10a55a"></span>Online</span><span class="legend-line"><span class="sw" style="background:#f5b700"></span>SS Parcial</span><span class="legend-line"><span class="ring" style="background:#f5b700;border-color:#b88700"></span>Porta crítica</span><span class="legend-line"><span class="sw" style="background:#ef3340"></span>SS Total</span><span class="legend-line"><span class="ring" style="border-color:#0091e6"></span>Em atendimento</span></div>''',unsafe_allow_html=True)


def operational_priority():
    if logical.empty: return pd.DataFrame()
    d=logical[logical["Portas_OFF"]>0].copy()
    if d.empty: return d
    d["Em_atendimento"]=(d["Tratativa"].map(norm_txt)=="EM ATENDIMENTO").astype(int); d["Sem_atendimento"]=(d["Tratativa"].map(norm_txt)=="SEM TRATATIVA").astype(int); d["Critico"]=d["Status"].isin(["OFF TOTAL","PARCIAL","PARCIAL CRÍTICO"]).astype(int); d["Regiao"]=d["Regiao"].replace("","SEM REGIÃO"); d["Bairro"]=d["Bairro"].replace("","SEM BAIRRO")
    rows=[]
    for (reg,b),g in d.groupby(["Regiao","Bairro"]):
        off=int(g["Portas_OFF"].sum()); crit=int(g["Critico"].sum()); atend=int(g["Em_atendimento"].sum()); sem=int(g["Sem_atendimento"].sum()); score=off*10+crit*15+sem*8
        acao="REFORÇAR EQUIPE" if sem>0 and (off>=3 or crit>0) else "DIRECIONAR EQUIPE" if sem>0 else "MANTER ATUAÇÃO"
        rows.append({"Região":reg,"Bairro":b,"Portas OFF":off,"Nodes críticos":crit,"Em atendimento":atend,"Sem atendimento":sem,"Ação":acao,"Score":score})
    return pd.DataFrame(rows).sort_values(["Score","Portas OFF"],ascending=False).reset_index(drop=True)


def haversine_km(lat1,lon1,lat2,lon2):
    try:
        r=6371.0; a1,a2=math.radians(float(lat1)),math.radians(float(lat2)); dp=math.radians(float(lat2)-float(lat1)); dl=math.radians(float(lon2)-float(lon1)); a=math.sin(dp/2)**2+math.cos(a1)*math.cos(a2)*math.sin(dl/2)**2; return 2*r*math.asin(math.sqrt(a))
    except Exception: return None


def node_priority_score(r):
    sev={"OFF TOTAL":100,"PARCIAL CRÍTICO":70,"PARCIAL":70}.get(r.get("Status"),0); return sev+int(r.get("Portas_OFF",0))*15+(25 if norm_txt(r.get("Tratativa"))=="SEM TRATATIVA" else 0)


def suggestions_for_tech(tech):
    if logical.empty: return pd.DataFrame()
    cand=logical[(logical["Portas_OFF"]>0)&(logical["Tratativa"].map(norm_txt)=="SEM TRATATIVA")].copy()
    if cand.empty: return cand
    cand["Prioridade_score"]=cand.apply(node_priority_score,axis=1); ref=None
    if not assignments_latest.empty:
        aa=assignments_latest[(assignments_latest["Tecnico"].map(norm_txt)==norm_txt(tech)) & (~assignments_latest["Status"].map(norm_txt).isin(["NORMALIZADO","CANCELADO"]))]
        if not aa.empty:
            last=aa.iloc[-1]; rr=logical[logical["Node"]==norm_txt(last.Node)]
            if not rr.empty and pd.notna(rr.iloc[0]["Latitude_Base"]): ref=(rr.iloc[0]["Latitude_Base"],rr.iloc[0]["Longitude_Base"])
    if ref:
        cand["Dist_km"]=cand.apply(lambda r:haversine_km(ref[0],ref[1],r["Latitude_Base"],r["Longitude_Base"]) if pd.notna(r["Latitude_Base"]) else None,axis=1); cand["Sugestao_score"]=cand["Prioridade_score"]-cand["Dist_km"].fillna(50)*8; cand=cand.sort_values(["Sugestao_score","Prioridade_score"],ascending=False)
    else:
        cand["Dist_km"]=pd.NA; cand=cand.sort_values("Prioridade_score",ascending=False)
    return cand.head(12)


# -----------------------------
# VISÃO EXECUTIVA
# -----------------------------
if view=="Executiva":
    render_kpis(); render_evolution()
    st.markdown('<div class="mobile-hint">📱 Toque em um node para centralizar e ver os detalhes. O mapa permanece em modo claro.</div>',unsafe_allow_html=True)

    # Resumo operacional compacto: três leituras lado a lado no PC e empilhadas no celular.
    c_reg,c_bairro,c_nodes=st.columns(3,gap="small")
    with c_reg:
        render_region_table()
    with c_bairro:
        render_neighborhood_table(limit=5)
    with c_nodes:
        render_top_nodes(show_treatment=False,limit=6)

    # Mapa ocupa toda a largura disponível.
    st.markdown('<div class="panel-title" style="margin-top:10px"><b>🗺️ Porto Alegre – RS</b><span>Rede HFC • todos os nodes</span></div>',unsafe_allow_html=True)
    render_map_search(); render_legend(); render_map(height=650)
    if crisis_mode: st.error(f"⚡ MODO CRISE ATIVO — {fmt_int(ports_off or 0)} portas OFF.")

    st.markdown('<div class="footer-version" style="text-align:center;color:#7a879c;font-size:10px;margin-top:12px">v10.1.4</div>', unsafe_allow_html=True)

# -----------------------------
# VISÃO SUPERVISOR
# -----------------------------
elif view=="Supervisor":
    if not check_admin_access(): st.stop()
    render_kpis(); render_evolution()
    op=operational_priority()
    st.markdown("### 🎯 Prioridade operacional")
    if not op.empty:
        rows=[]
        for i,r in op.head(12).iterrows():
            cls="action-high" if r["Ação"]=="REFORÇAR EQUIPE" else "action-mid" if r["Ação"]=="DIRECIONAR EQUIPE" else "action-ok"
            rows.append(f"<tr><td>{i+1}</td><td><b>{esc(r['Região'])}</b></td><td>{esc(r['Bairro'])}</td><td>{int(r['Portas OFF'])}</td><td>{int(r['Nodes críticos'])}</td><td>{int(r['Em atendimento'])}</td><td>{int(r['Sem atendimento'])}</td><td class='{cls}'>{esc(r['Ação'])}</td></tr>")
        st.markdown(f'<div class="panel"><table class="rank"><tr><th>#</th><th>Região</th><th>Bairro</th><th>Portas OFF</th><th>Críticos</th><th>Em atendimento</th><th>Sem atendimento</th><th>Ação</th></tr>{"".join(rows)}</table></div>',unsafe_allow_html=True)
    else: st.success("Nenhuma prioridade operacional aberta.")

    c1,c2=st.columns([1.35,1])
    with c1:
        st.markdown("### 👥 Equipe do dia")
        team=load_team(); today=now_poa().strftime("%Y-%m-%d"); current=team[team["Data"].astype(str)==today].copy()
        if current.empty: current=pd.DataFrame([{"Data":today,"Tecnico":"","Equipe":"","Turno":"","Regiao_Preferencial":"","Ativo":True,"Atualizado_em":""}])
        edit=current[["Tecnico","Equipe","Turno","Regiao_Preferencial","Ativo"]].copy(); edit["Ativo"]=edit["Ativo"].astype(str).str.upper().map({"TRUE":True,"1":True,"SIM":True,"YES":True,"FALSE":False,"0":False,"NAO":False,"NÃO":False}).fillna(True)
        edited=st.data_editor(edit,num_rows="dynamic",use_container_width=True,hide_index=True,column_config={"Ativo":st.column_config.CheckboxColumn("Ativo")},key="team_editor")
        if st.button("💾 Salvar equipe do dia",use_container_width=True):
            keep=team[team["Data"].astype(str)!=today].copy(); new=edited.copy(); new["Data"]=today; new["Tecnico"]=new["Tecnico"].map(norm_txt); new["Equipe"]=new["Equipe"].map(norm_txt); new["Turno"]=new["Turno"].map(norm_txt); new["Regiao_Preferencial"]=new["Regiao_Preferencial"].map(norm_txt); new["Atualizado_em"]=now_poa().strftime("%d/%m/%Y %H:%M"); new=new[new["Tecnico"].astype(str).str.len()>0]; ok,msg=save_team(pd.concat([keep,new],ignore_index=True)); st.success(msg) if ok else st.error(msg); st.rerun() if ok else None
    with c2:
        st.markdown("### 📌 Resumo de despacho")
        active=assignments_latest[~assignments_latest["Status"].map(norm_txt).isin(["NORMALIZADO","CANCELADO"])] if not assignments_latest.empty else assignments_latest
        team_today=load_team(); team_today=team_today[(team_today["Data"].astype(str)==today)&(team_today["Ativo"].astype(str).str.upper().isin(["TRUE","1","SIM","YES"]))]
        cards=[mini_html("Técnicos disponíveis",len(team_today["Tecnico"].dropna().unique()),"equipe ativa hoje"),mini_html("Nodes atribuídos",len(active),"fila atual"),mini_html("Nodes sem atendimento",sem_atendimento,"porta OFF sem equipe"),mini_html("Em atendimento",em_atendimento,"atuação registrada")]
        st.markdown('<div class="mini-grid">'+''.join(cards)+'</div>',unsafe_allow_html=True)

    st.markdown("### 🚚 Distribuir nodes")
    team_today=load_team(); team_today=team_today[(team_today["Data"].astype(str)==today)&(team_today["Ativo"].astype(str).str.upper().isin(["TRUE","1","SIM","YES"]))].copy(); techs=sorted(team_today["Tecnico"].map(norm_txt).dropna().unique().tolist())
    if not techs: st.info("Cadastre a Equipe do dia antes de distribuir atendimentos.")
    else:
        tech=st.selectbox("Técnico / equipe",techs,key="dispatch_tech"); sug=suggestions_for_tech(tech)
        if not sug.empty:
            show=sug[["Node","Regiao","Bairro","Portas_OFF","Status","Dist_km"]].copy(); show.columns=["Node","Região","Bairro","Portas OFF","Status","Distância km"]; show["Distância km"]=pd.to_numeric(show["Distância km"],errors="coerce").round(1); st.caption("Sugestões combinam criticidade, portas OFF e proximidade de onde o técnico já está atuando."); st.dataframe(show.head(8),use_container_width=True,hide_index=True)
            options=sug["Node"].tolist()+[n for n in logical[(logical["Portas_OFF"]>0)&(logical["Tratativa"].map(norm_txt)=="SEM TRATATIVA")]["Node"].tolist() if n not in set(sug["Node"])]
            selected=st.multiselect("Nodes para atribuir",options,default=options[:1] if options else [])
            if st.button("➡️ Atribuir ao técnico",use_container_width=True,disabled=not selected):
                all_a=load_assignments(); eqrow=team_today[team_today["Tecnico"].map(norm_txt)==norm_txt(tech)]; equipe=norm_txt(eqrow.iloc[0]["Equipe"]) if not eqrow.empty else ""; now=now_poa().strftime("%d/%m/%Y %H:%M"); out=all_a.copy()
                for order,node in enumerate(selected,1):
                    out=out[out["Node"].map(norm_txt)!=norm_txt(node)].copy(); rr=logical[logical["Node"]==norm_txt(node)]; pri=int(node_priority_score(rr.iloc[0])) if not rr.empty else 0; out=pd.concat([out,pd.DataFrame([{"Node":norm_txt(node),"Tecnico":norm_txt(tech),"Equipe":equipe,"Prioridade":pri,"Ordem":order,"Status":"ATRIBUÍDO","Situacao":"ATRIBUÍDO","Observacao":"","Atribuido_em":now,"Atualizado_em":now,"Foto_Path":""}])],ignore_index=True)
                ok,msg=save_assignments(out); st.success(msg) if ok else st.error(msg); st.rerun() if ok else None
        else: st.success("Não há nodes sem atendimento para distribuir.")

    st.markdown("### 🗺️ Operação no mapa")
    render_map(height=500); render_region_table(); render_neighborhood_table(); render_top_nodes(show_treatment=True,limit=12)

    active_t=logical[logical["Tratativa"].map(norm_txt)=="EM ATENDIMENTO"].copy() if not logical.empty else pd.DataFrame()
    st.markdown("### 🛠️ Nodes em atendimento")
    if not active_t.empty:
        active_t=active_t.sort_values(["Portas_OFF","Status"],ascending=False); rows=[]
        for i,r in enumerate(active_t.itertuples()): rows.append(f"<tr><td>{i+1}</td><td><b>{esc(r.Node)}</b></td><td>{esc(getattr(r,'Responsavel','') or '—')}</td><td>{esc(getattr(r,'Situacao','') or '—')}</td><td>{int(r.Portas_OFF)}</td></tr>")
        st.markdown(f'<div class="panel"><table class="rank"><tr><th>#</th><th>Node</th><th>Técnico</th><th>Situação</th><th>OFF</th></tr>{"".join(rows)}</table></div>',unsafe_allow_html=True)
    else: st.info("Nenhum node em atendimento no momento.")

    with st.expander("⏱️ Registrar atualização da crise / outages",expanded=False):
        with st.form("snapshot_form"):
            outages_input=st.number_input("Outages sem sinal",min_value=0,step=1,value=int(outages_current or 0)); obs=st.text_input("Observação (opcional)"); save=st.form_submit_button("💾 Registrar atualização",use_container_width=True)
        if save:
            now=now_poa(); row=_snapshot_row(now,int(outages_input),"MANUAL",source_key,obs.strip()); hd=history_df.drop(columns=["DataHora_dt"],errors="ignore") if not history_df.empty else pd.DataFrame(columns=HISTORY_COLUMNS); hd=pd.concat([hd,pd.DataFrame([row])],ignore_index=True); ok,msg=save_history(hd); st.success(msg) if ok else st.error(msg); st.rerun() if ok else None

# -----------------------------
# VISÃO TÉCNICO
# -----------------------------
else:
    technician=check_tech_access()
    if not technician: st.stop()
    st.markdown(f"### 👷 {esc(technician)} — meus atendimentos")
    a=load_assignments(); a["Node"]=a["Node"].map(norm_txt); a["Tecnico"]=a["Tecnico"].map(norm_txt); a["Status"]=a["Status"].map(norm_txt)
    mine=a[(a["Tecnico"]==norm_txt(technician)) & (~a["Status"].isin(["CANCELADO"]))].copy()
    if mine.empty: st.info("Você não possui nodes atribuídos no momento."); st.stop()
    if logical.empty:
        st.warning("Aguardando a leitura do XPERTrack para carregar os dados técnicos dos nodes atribuídos.")
        st.stop()
    mine_latest=mine.drop_duplicates("Node",keep="last"); merged=mine_latest.merge(logical[["Node","Regiao","Bairro","Portas_OFF","Total_portas","Status","Latitude","Longitude","Latitude_Base","Longitude_Base","Origem","Portas_popup","color","line_color","Status_exibicao","Regiao_exibicao","Bairro_exibicao","Tratativa_txt"]],on="Node",how="left",suffixes=("_Atend",""))
    active_mine=merged[~merged["Status_Atend"].map(norm_txt).isin(["NORMALIZADO","CONCLUÍDO","CONCLUIDO"])]
    cards=[mini_html("Na fila",len(active_mine),"nodes atribuídos"),mini_html("Portas OFF",int(active_mine["Portas_OFF"].fillna(0).sum()),"na sua fila"),mini_html("Normalizados",int(merged["Status_Atend"].map(norm_txt).isin(["NORMALIZADO","CONCLUÍDO","CONCLUIDO"]).sum()),"atendimentos concluídos")]
    st.markdown('<div class="mini-grid">'+''.join(cards)+'</div>',unsafe_allow_html=True)
    if not active_mine.empty:
        st.markdown("#### 🗺️ Minha rota / nodes atribuídos"); render_map(active_mine,height=360,assigned_only=True)
        show=active_mine[["Node","Ordem","Regiao","Bairro","Portas_OFF","Status_Atend"]].copy(); show.columns=["Node","Ordem","Região","Bairro","Portas OFF","Atendimento"]; show=show.sort_values(["Ordem","Portas OFF"],ascending=[True,False]); st.dataframe(show,use_container_width=True,hide_index=True)
    choices=merged["Node"].tolist(); selected=st.selectbox("Node para atualizar",choices); cur=mine_latest[mine_latest["Node"]==selected].iloc[-1]
    statuses=["ATRIBUÍDO","EM DESLOCAMENTO","NO LOCAL","EM ATUAÇÃO","ROMPIMENTO LOCALIZADO","AGUARDANDO ENERGIA","AGUARDANDO ACESSO","NORMALIZADO"]
    current_status=norm_txt(cur.get("Status","ATRIBUÍDO")); idx=statuses.index(current_status) if current_status in statuses else 0
    with st.form("tech_update",clear_on_submit=False):
        status=st.selectbox("Status do atendimento",statuses,index=idx); obs=st.text_area("Observação",value=str(cur.get("Observacao","") or ""),placeholder="Ex.: rompimento localizado; sem energia; aguardando acesso..."); photo=st.file_uploader("Foto (opcional)",type=["jpg","jpeg","png","webp"],accept_multiple_files=False); submit=st.form_submit_button("💾 Atualizar atendimento",use_container_width=True)
    if submit:
        photo_path=str(cur.get("Foto_Path","") or "")
        if photo is not None:
            okp,pth,msgp=save_photo(photo,selected,technician)
            if not okp: st.error(msgp); st.stop()
            photo_path=pth
        all_a=load_assignments(); mask=all_a["Node"].map(norm_txt)==norm_txt(selected); base_row=cur.to_dict(); base_row.update({"Node":norm_txt(selected),"Tecnico":norm_txt(technician),"Status":status,"Situacao":status,"Observacao":obs.strip(),"Atualizado_em":now_poa().strftime("%d/%m/%Y %H:%M"),"Foto_Path":photo_path}); all_a=all_a[~mask].copy(); all_a=pd.concat([all_a,pd.DataFrame([base_row])],ignore_index=True); ok,msg=save_assignments(all_a); st.success("Atendimento atualizado.") if ok else st.error(msg); st.rerun() if ok else None
