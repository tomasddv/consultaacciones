from __future__ import annotations

import json
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st


BASE = Path(__file__).resolve().parent
DATA = BASE / "data"


def fold(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    return "".join(ch for ch in text if not unicodedata.combining(ch)).upper().strip()


def load_data() -> tuple[pd.DataFrame, dict]:
    frame = pd.read_csv(DATA / "client_discounts.csv", dtype={"codigo": str}, encoding="utf-8-sig").fillna("")
    meta = json.loads((DATA / "snapshot_meta.json").read_text(encoding="utf-8"))
    frame["busqueda"] = (
        frame["codigo"].map(fold) + " " + frame["fantasia"].map(fold) + " " + frame["razon_social"].map(fold)
    )
    return frame, meta


st.set_page_config(page_title="Descuentos DDV", page_icon="%", layout="wide", initial_sidebar_state="collapsed")
st.markdown(
    """
    <style>
    :root { --ink:#17212b; --muted:#667085; --line:#d7dde5; --paper:#ffffff; --soft:#f4f7f9; --blue:#0866c6; --green:#087f5b; --amber:#9a6700; }
    .stApp { background:#f5f7f9; color:var(--ink); }
    .block-container { max-width:980px; padding-top:1rem; padding-bottom:3rem; }
    h1,h2,h3,p { letter-spacing:0 !important; }
    h1 { font-size:1.65rem !important; margin:0 0 .15rem 0 !important; }
    .eyebrow { color:var(--blue); font-weight:800; font-size:.75rem; text-transform:uppercase; margin-bottom:.2rem; }
    .subtle { color:var(--muted); font-size:.9rem; margin-bottom:1rem; }
    .client-head { background:var(--paper); border:1px solid var(--line); border-radius:6px; padding:1rem; margin:.75rem 0; }
    .client-name { font-size:1.22rem; font-weight:800; }
    .client-meta { color:var(--muted); font-size:.88rem; margin-top:.3rem; }
    .discount-card { background:var(--paper); border:1px solid var(--line); border-left:5px solid var(--blue); border-radius:6px; padding:.9rem 1rem; min-height:138px; margin:.35rem 0; }
    .family { font-size:.78rem; font-weight:800; color:var(--muted); text-transform:uppercase; }
    .pct { font-size:2.05rem; line-height:1.1; font-weight:850; margin:.28rem 0; }
    .threshold { color:var(--green); font-size:.92rem; font-weight:750; }
    .warning { color:var(--amber); font-size:.78rem; margin-top:.35rem; }
    div[data-testid="stTextInput"] input { min-height:48px; font-size:1rem; }
    div[data-testid="stButton"] button { min-height:46px; border-radius:6px; font-weight:750; }
    div[data-testid="stExpander"] { border-color:var(--line); border-radius:6px; background:var(--paper); }
    @media (max-width: 640px) {
      .block-container { padding: .75rem .75rem 2rem; }
      h1 { font-size:1.42rem !important; }
      .discount-card { min-height:0; }
      .pct { font-size:1.8rem; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

df, meta = load_data()
st.markdown('<div class="eyebrow">Distribuidora del Valle</div>', unsafe_allow_html=True)
st.title("Consulta de descuentos")
st.markdown('<div class="subtle">Octubre 2026 · buscá por código o nombre de fantasía</div>', unsafe_allow_html=True)

query = st.text_input("Cliente", placeholder="Ejemplo: 4020 o Gardelitos", label_visibility="collapsed")
matches = pd.DataFrame()
if query.strip():
    terms = [fold(term) for term in query.split() if term.strip()]
    mask = df["busqueda"].map(lambda value: all(term in value for term in terms))
    matches = df[mask]

if not query.strip():
    st.info("Ingresá un código o nombre para ver sus descuentos máximos.")
elif matches.empty:
    st.warning("No encontré ese cliente entre los grupos relevados.")
else:
    choices = (
        matches[["codigo", "fantasia"]]
        .drop_duplicates()
        .assign(label=lambda x: x["codigo"] + " · " + x["fantasia"])
    )
    selected_label = st.selectbox("Coincidencias", choices["label"].tolist(), label_visibility="collapsed")
    selected_code = selected_label.split(" · ", 1)[0]
    client = df[df["codigo"] == selected_code].copy()
    first = client.iloc[0]

    st.markdown(
        f"""
        <div class="client-head">
          <div class="client-name">{first['codigo']} · {first['fantasia']}</div>
          <div class="client-meta">Promotor: {first['promotor'] or 'Sin dato'}<br>Ruta: {first['ruta'] or 'Sin dato'}<br>Subcanal: {first['subcanal'] or 'Sin dato'}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    family_priority = {
        "VALUE LATA 473": 10,
        "VALUE LITRO 1000": 11,
        "VALUE LATON 710": 12,
        "CORE RGB / 473": 20,
        "CORE 269": 21,
        "CORE LATA": 22,
        "CORE SABORIZADAS 473": 23,
        "LATONES": 30,
        "PREMIUM LATON": 31,
        "CERVEZA 0.0": 40,
    }
    client["orden"] = client["familia"].map(family_priority).fillna(50)
    rows = client.sort_values(["orden", "familia", "porcentaje_maximo"], ascending=[True, True, False]).to_dict("records")
    for start in range(0, len(rows), 2):
        columns = st.columns(2)
        for column, row in zip(columns, rows[start:start + 2]):
            inferred = bool(row["tiene_inferidos"])
            warning = '<div class="warning">Grupo pendiente de validación final en ERP</div>' if inferred else ""
            with column:
                st.markdown(
                    f"""
                    <div class="discount-card">
                      <div class="family">{row['familia']}</div>
                      <div class="pct">{row['porcentaje_maximo']:.2f}%</div>
                      <div class="threshold">Máximo desde {int(row['desde_bultos'])} bulto{'s' if int(row['desde_bultos']) != 1 else ''}</div>
                      {warning}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                details = json.loads(row["detalle"])
                detail_label = f"Ver {len(details)} {'acciones' if len(details) != 1 else 'acción'}"
                with st.expander(detail_label):
                    for detail in details:
                        action = detail["accion_id"] or "s/n"
                        st.markdown(f"**{detail['promo']} / acción {action}**")
                        st.caption(f"{detail['nombre']} · Grupo {detail['grupo']} {detail['grupo_nombre']}")
                        st.write(detail["descuento_regla"])

with st.sidebar:
    st.subheader("Estado del relevamiento")
    st.metric("Acciones del Excel", meta["actions_total"])
    st.metric("Acciones confirmadas", meta["actions_confirmed"])
    st.metric("Excluidas por no vigencia", meta["actions_excluded"])
    st.metric("Pendientes", meta["actions_pending"])
    st.caption("Prototipo local. No modifica Drive, GitHub ni el dashboard vigente.")

