import io
import time
import requests
import pandas as pd
import streamlit as st
import folium
from streamlit_folium import st_folium

st.set_page_config(page_title="Rio Grande • Nodes HFC", page_icon="📡", layout="wide")

DRIVE_FILE_ID = "1khmqXK9bpeTub9eXgluJrkCYm8491sCY"
DRIVE_CSV_URL = f"https://drive.google.com/uc?export=download&id={DRIVE_FILE_ID}"

ALIASES = {
    "TRVABA3": "TRVABA",
    "TRVACD": "TRVACC",
}

def norm(s):
    return str(s).strip().upper().replace("\ufeff", "")

def base_node(port):
    p = norm(port)
    if "-" in p and p.rsplit("-", 1)[-1].isdigit():
        p = p.rsplit("-", 1)[0]
    return ALIASES.get(p, p)

@st.cache_data(ttl=60, show_spinner=False)
def load_geo():
    df = pd.read_csv("base_nodes_rio_grande.csv", encoding="utf-8-sig")
    df["Node"] = df["Node"].map(norm)
    df["Geo_origem"] = df["Geo_origem"].map(norm)
    return df

@st.cache_data(ttl=60, show_spinner=False)
def load_drive_csv(_tick):
    # Cache-busting + multiple decoders because XPERTrack exports may vary.
    url = DRIVE_CSV_URL + f"&v={int(time.time())}"
    r = requests.get(url, timeout=20, allow_redirects=True)
    r.raise_for_status()
    raw = r.content

    # If Drive returns an HTML login/permission page, fail clearly.
    ct = (r.headers.get("content-type") or "").lower()
    if b"<html" in raw[:500].lower() or "text/html" in ct:
        raise RuntimeError(
            "O Google Drive não liberou o CSV para leitura pública. "
            "No Drive, deixe RIO GRANDE.csv como 'Qualquer pessoa com o link — Leitor'."
        )

    last = None
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            txt = raw.decode(enc)
            df = pd.read_csv(io.StringIO(txt))
            df.columns = [norm(c) for c in df.columns]
            needed = {"NODE","PONTUAÇÃO","IMPACTADO","ESTRESSADO","TOTAL"}
            if needed.issubset(df.columns):
                return df
        except Exception as e:
            last = e
    raise RuntimeError(f"CSV recebido, mas o formato não foi reconhecido: {last}")

geo = load_geo()

st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 1rem;}
[data-testid="stMetric"] {background:#0e1a2b; border:1px solid #263a56; padding:10px 14px; border-radius:12px;}
.small {color:#8fa4bd; font-size:.88rem}
</style>
""", unsafe_allow_html=True)

c1, c2 = st.columns([5,1])
with c1:
    st.title("RIO GRANDE • MAPA DE NODES HFC")
    st.caption("Base geográfica fixa • coleta operacional atualizada pelo arquivo RIO GRANDE.csv no Google Drive")
with c2:
    if st.button("🔄 Atualizar agora", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

try:
    coleta = load_drive_csv(int(time.time() // 60))
    coleta["NODE_ORIG"] = coleta["NODE"].map(norm)
    coleta["BASE"] = coleta["NODE_ORIG"].map(base_node)
    for c in ["PONTUAÇÃO","IMPACTADO","ESTRESSADO","TOTAL"]:
        coleta[c] = pd.to_numeric(coleta[c], errors="coerce").fillna(0)

    groups = {}
    for base, g in coleta.groupby("BASE"):
        groups[base] = {
            "score": round(float(g["PONTUAÇÃO"].mean()), 1),
            "impactado": int(g["IMPACTADO"].sum()),
            "estressado": int(g["ESTRESSADO"].sum()),
            "total": int(g["TOTAL"].sum()),
            "ports": g[["NODE_ORIG","PONTUAÇÃO","IMPACTADO","ESTRESSADO","TOTAL"]].to_dict("records"),
        }

    def find_group(row):
        for key in (row["Node"], row["Geo_origem"]):
            if key in groups:
                return groups[key]
        return None

    geo["coleta"] = geo.apply(find_group, axis=1)
    geo["com_coleta"] = geo["coleta"].notna()
    geo["impactado"] = geo["coleta"].apply(lambda x: x["impactado"] if isinstance(x, dict) else 0)

    report_bases = set(coleta["BASE"])
    geo_keys = set(geo["Node"]) | set(geo["Geo_origem"])
    unmatched = sorted(report_bases - geo_keys - set(ALIASES.keys()))

    m1,m2,m3,m4 = st.columns(4)
    m1.metric("Nodes no mapa", len(geo))
    m2.metric("Com coleta", int(geo["com_coleta"].sum()))
    m3.metric("Impactados", int(((geo["com_coleta"]) & (geo["impactado"] > 0)).sum()))
    m4.metric("Sem geolocalização", len(unmatched))

    st.success("Coleta carregada diretamente do Google Drive.", icon="✅")

    f1,f2 = st.columns([2,1])
    busca = f1.text_input("Pesquisar node ou bairro", placeholder="Ex.: CNTAD ou Cassino")
    bairros = ["Todos"] + sorted([x for x in geo["Bairro"].dropna().unique()])
    bairro = f2.selectbox("Bairro", bairros)

    view = geo.copy()
    if busca:
        q = busca.upper()
        view = view[
            view["Node"].str.contains(q, case=False, na=False) |
            view["Bairro"].astype(str).str.contains(q, case=False, na=False)
        ]
    if bairro != "Todos":
        view = view[view["Bairro"] == bairro]

    center = [float(geo["Latitude"].mean()), float(geo["Longitude"].mean())]
    fmap = folium.Map(location=center, zoom_start=12, tiles="CartoDB dark_matter", control_scale=True)

    for _, r in view.iterrows():
        c = r["coleta"] if isinstance(r["coleta"], dict) else None
        if c:
            color = "#f1c40f" if c["impactado"] > 0 else "#2ecc71"
            port_lines = "".join(
                f"<div>{p['NODE_ORIG']} • Pont. {p['PONTUAÇÃO']:.0f} • "
                f"Impact. {int(p['IMPACTADO'])} • Estr. {int(p['ESTRESSADO'])}</div>"
                for p in c["ports"]
            )
            detail = (
                f"<b>Pontuação média:</b> {c['score']}<br>"
                f"<b>Impactado:</b> {c['impactado']}<br>"
                f"<b>Estressado:</b> {c['estressado']}<br>"
                f"<b>Total:</b> {c['total']}<hr>{port_lines}"
            )
        else:
            color = "#8291a5"
            detail = "Sem coleta correspondente nesta extração."

        route = f"https://www.google.com/maps/dir/?api=1&destination={r['Latitude']},{r['Longitude']}"
        popup = f"""
        <div style="font-family:Arial;min-width:280px">
          <h3 style="margin:0 0 8px">{r['Node']}</h3>
          <b>Bairro:</b> {r['Bairro']}<br>{detail}<br><br>
          <a href="{route}" target="_blank"
             style="background:#1473e6;color:white;padding:7px 10px;border-radius:6px;text-decoration:none">
             Traçar rota no Google Maps
          </a>
        </div>"""
        folium.CircleMarker(
            location=[r["Latitude"], r["Longitude"]],
            radius=7, color="#06101c", weight=2,
            fill=True, fill_color=color, fill_opacity=.95,
            popup=folium.Popup(popup, max_width=380),
            tooltip=r["Node"],
        ).add_to(fmap)

    st_folium(fmap, height=650, use_container_width=True, returned_objects=[])

    if unmatched:
        with st.expander(f"⚠️ Códigos da coleta sem geolocalização ({len(unmatched)})"):
            st.write(", ".join(unmatched))

except Exception as e:
    st.error("Não foi possível carregar a coleta do Google Drive.")
    st.warning(str(e))
    st.info(
        "A base geográfica está preservada. Para a atualização automática funcionar no Streamlit, "
        "o arquivo RIO GRANDE.csv precisa estar compartilhado como “Qualquer pessoa com o link — Leitor”."
    )
