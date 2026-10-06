import io

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

st.set_page_config(page_title="Dashboard OTD", page_icon="📦", layout="wide")

NINGUNA = "(ninguna)"
CATS_DEFECTO = {"racks", "shelves", "peripherals", "tray"}
COLORES = ["#1f4e79", "#e07b39", "#2e6b2e", "#2a9bd1", "#8e44ad", "#c0392b"]

# ---------- Contraseña opcional ----------
def clave_configurada():
    try:
        return st.secrets.get("password")
    except Exception:
        return None

clave = clave_configurada()
if clave:
    if st.text_input("Contraseña", type="password") != clave:
        st.stop()

# ---------- Datos ----------
def datos_ejemplo():
    rng = np.random.default_rng(7)
    fechas = pd.date_range("2026-01-05", periods=38, freq="W-MON")
    df = pd.DataFrame({"Date": fechas,
                       "CW": fechas.isocalendar().week.astype(int).values,
                       "Year": fechas.year})
    cats = ["Racks", "Shelves", "Peripherals", "Tray"]
    for c, base in zip(cats, [0.99, 0.975, 0.97, 0.985]):
        df[c] = np.clip(base + rng.normal(0, 0.012, len(df)), 0.85, 1.0)
    df["Vol"] = rng.integers(40000, 180000, len(df))
    df["Total"] = df[cats].mean(axis=1)
    df["Target"] = 0.95
    df["Notas"] = ""
    df.loc[30, "Notas"] = "shutdown week"
    return df

def a_fraccion(serie):
    s = pd.to_numeric(serie, errors="coerce")
    if s.notna().any() and s.median() > 1.5:   # vienen como 98.5 en vez de 0.985
        s = s / 100
    return s

st.title("📦 Dashboard OTD (on time delivery)")

fuente = st.radio("Fuente de datos", ["Datos de ejemplo", "Subir mi Excel"], horizontal=True)
if fuente == "Datos de ejemplo":
    df_raw = datos_ejemplo()
else:
    archivo = st.file_uploader("Sube tu archivo (.xlsx o .csv)", type=["xlsx", "csv"])
    if archivo is None:
        st.info("Sube tu archivo para ver el dashboard. La tabla debe empezar en la celda A1 de su hoja.")
        st.stop()
    if archivo.name.lower().endswith(".csv"):
        df_raw = pd.read_csv(archivo)
    else:
        xls = pd.ExcelFile(archivo)
        hoja = st.selectbox("Hoja", xls.sheet_names)
        df_raw = xls.parse(hoja)

cols = list(df_raw.columns)
opc = [NINGUNA] + cols

def pick(candidatos, opcional=False):
    for i, c in enumerate(cols):
        if str(c).strip().lower() in candidatos:
            return i + 1 if opcional else i
    return 0

# ---------- Panel lateral: columnas y opciones ----------
with st.sidebar:
    st.subheader("Columnas de tu tabla")
    c_fecha = st.selectbox("Fecha", cols, index=pick({"date", "fecha"}))
    c_total = st.selectbox("OTD total", cols, index=pick({"total res gdl", "total", "total otd", "otd"}))
    c_vol = st.selectbox("Volumen", opc, index=pick({"vol", "volume", "volumen"}, True))
    c_sem = st.selectbox("Semana (CW)", opc, index=pick({"cw", "semana", "week", "wk"}, True))
    c_anio = st.selectbox("Año", opc, index=pick({"year", "año", "anio"}, True))
    c_target = st.selectbox("Meta (target)", opc, index=pick({"target", "meta"}, True))
    c_nota = st.selectbox("Notas", opc, index=pick({"column1", "notas", "nota", "notes", "comentarios"}, True))
    defecto_cats = [c for c in cols if str(c).strip().lower() in CATS_DEFECTO]
    categorias = st.multiselect("Categorías (columnas de OTD)", cols, default=defecto_cats)
    meta_fija = st.number_input("Meta fija (%) si no hay columna de meta", value=95.0, step=0.5) / 100
    st.subheader("Opciones")
    eje = st.radio("Eje X", ["Semana (CW)", "Fecha"])
    ponderado = st.checkbox("Promedio ponderado por volumen", value=False)
    excluir = st.checkbox("Excluir semanas con nota (ej. paro)", value=False)
    etiquetas = st.checkbox("Mostrar etiquetas de datos", value=True)

# ---------- Limpieza ----------
d = pd.DataFrame(index=df_raw.index)
d["fecha"] = pd.to_datetime(df_raw[c_fecha], errors="coerce")
d["total"] = a_fraccion(df_raw[c_total])
d["vol"] = pd.to_numeric(df_raw[c_vol], errors="coerce") if c_vol != NINGUNA else np.nan
d["target"] = a_fraccion(df_raw[c_target]) if c_target != NINGUNA else meta_fija
d["target"] = d["target"].fillna(meta_fija)
d["nota"] = df_raw[c_nota].fillna("").astype(str).str.strip() if c_nota != NINGUNA else ""
d["año"] = pd.to_numeric(df_raw[c_anio], errors="coerce") if c_anio != NINGUNA else d["fecha"].dt.year
d["sem"] = (pd.to_numeric(df_raw[c_sem], errors="coerce") if c_sem != NINGUNA
            else d["fecha"].dt.isocalendar().week.astype(float))
for c in categorias:
    d[c] = a_fraccion(df_raw[c])
d = d.dropna(subset=["total", "año", "sem"]).copy()
if d.empty:
    st.warning("No encontré filas válidas. Revisa que las columnas del panel lateral sean las correctas.")
    st.stop()
d["año"] = d["año"].astype(int)
d["sem"] = d["sem"].astype(int)
d = d.sort_values(["año", "sem"]).reset_index(drop=True)

# ---------- Filtros ----------
años = sorted(d["año"].unique())
sel_años = st.multiselect("Año", años, default=años)
d = d[d["año"].isin(sel_años)].reset_index(drop=True)
if excluir:
    d = d[d["nota"] == ""].reset_index(drop=True)
if d.empty:
    st.warning("No hay datos con esos filtros.")
    st.stop()

if eje == "Fecha" and d["fecha"].notna().all():
    d["x"] = d["fecha"].dt.strftime("%d-%b")
elif len(sel_años) > 1:
    d["x"] = d["año"].astype(str) + "-S" + d["sem"].astype(str).str.zfill(2)
else:
    d["x"] = d["sem"].astype(str)

if len(d) > 1:
    i0, i1 = st.select_slider("Rango de periodos", options=list(range(len(d))),
                              value=(0, len(d) - 1), format_func=lambda i: d["x"].iloc[i])
    d = d.iloc[i0:i1 + 1]

sel_cats = st.multiselect("Categorías a mostrar", categorias, default=categorias)

def prom(serie):
    s = serie.dropna()
    if s.empty:
        return float("nan")
    if ponderado and d["vol"].notna().any():
        w = d.loc[s.index, "vol"].fillna(0)
        if w.sum() > 0:
            return float((s * w).sum() / w.sum())
    return float(s.mean())

# ---------- Tarjetas ----------
meta = float(d["target"].iloc[-1])
otd = prom(d["total"])
k1, k2, k3, k4 = st.columns(4)
k1.metric("OTD promedio", f"{otd:.2%}", f"{(otd - meta) * 100:+.1f} pts vs meta")
k2.metric("Volumen total", f"{d['vol'].sum():,.0f}" if d["vol"].notna().any() else "—")
k3.metric("Meta", f"{meta:.0%}")
k4.metric("Periodos bajo la meta", f"{int((d['total'] < d['target']).sum())} de {len(d)}")

# ---------- Gráfico 1: OTD total, volumen y meta ----------
fig = make_subplots(specs=[[{"secondary_y": True}]])
hay_vol = d["vol"].notna().any()
if hay_vol:
    barras = dict(x=d["x"], y=d["vol"], name="Volumen", marker_color="#fbe0d3")
    if etiquetas:
        barras.update(text=d["vol"].map(lambda v: f"{v:,.0f}"), textposition="inside",
                      textangle=-90, textfont=dict(size=9))
    fig.add_trace(go.Bar(**barras), secondary_y=False)
linea = dict(x=d["x"], y=d["total"], name="OTD total", line=dict(color="#1f4e79", width=2),
             mode="lines+markers+text" if etiquetas else "lines+markers")
if etiquetas:
    linea.update(text=d["total"].map(lambda v: f"{v:.0%}"), textposition="top center")
fig.add_trace(go.Scatter(**linea), secondary_y=True)
bajo = d[d["total"] < d["target"]]
if not bajo.empty:
    fig.add_trace(go.Scatter(x=bajo["x"], y=bajo["total"], mode="markers", name="Bajo la meta",
                             marker=dict(color="red", size=10)), secondary_y=True)
fig.add_trace(go.Scatter(x=d["x"], y=d["target"], name="Meta", mode="lines",
                         line=dict(color="green", dash="dash")), secondary_y=True)
for r in d[d["nota"] != ""].itertuples():
    fig.add_annotation(x=r.x, y=r.total, yref="y2", text=r.nota, showarrow=True,
                       arrowhead=2, ay=-45, font=dict(size=10))
piso = max(0.0, min(d["total"].min(), d["target"].min()) - 0.04)
fig.update_yaxes(tickformat=".0%", range=[piso, 1.03], secondary_y=True)
fig.update_yaxes(showgrid=False, title_text="Volumen", secondary_y=False,
                 range=[0, d["vol"].max() * 1.25] if hay_vol else None)
fig.update_layout(title="OTD, volumen y meta", height=480, xaxis=dict(type="category"),
                  legend=dict(orientation="h", y=-0.2), margin=dict(l=10, r=10, t=50, b=10))
st.plotly_chart(fig, use_container_width=True)

# ---------- Gráfico 2: OTD por categoría ----------
if sel_cats:
    cols_k = st.columns(len(sel_cats))
    for col, c in zip(cols_k, sel_cats):
        col.metric(f"OTD {c}", f"{prom(d[c]):.1%}")
    fig2 = go.Figure()
    for i, c in enumerate(sel_cats):
        fig2.add_trace(go.Scatter(x=d["x"], y=d[c], name=c, mode="lines+markers",
                                  line=dict(color=COLORES[i % len(COLORES)], width=2)))
    fig2.add_trace(go.Scatter(x=d["x"], y=d["target"], name="Meta", mode="lines",
                              line=dict(color="purple", dash="dash")))
    minimo = min(d[c].min() for c in sel_cats)
    fig2.update_layout(title="OTD por categoría", height=420, xaxis=dict(type="category"),
                       yaxis=dict(tickformat=".0%", range=[max(0.0, minimo - 0.03), 1.02]),
                       legend=dict(orientation="h", y=-0.2), margin=dict(l=10, r=10, t=50, b=10))
    st.plotly_chart(fig2, use_container_width=True)

# ---------- Tabla y descarga ----------
st.subheader("Detalle")
tabla = d[["fecha", "año", "sem", "total"] + sel_cats + ["vol", "target", "nota"]].rename(
    columns={"fecha": "Fecha", "año": "Año", "sem": "Semana", "total": "OTD total",
             "vol": "Volumen", "target": "Meta", "nota": "Nota"})
config = {c: st.column_config.NumberColumn(format="percent") for c in ["OTD total", "Meta"] + sel_cats}
st.dataframe(tabla, hide_index=True, column_config=config)
buf = io.BytesIO()
tabla.to_excel(buf, index=False)
st.download_button("⬇️ Descargar a Excel", buf.getvalue(), file_name="otd_filtrado.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
