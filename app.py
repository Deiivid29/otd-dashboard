import datetime as dt
import io
import os
import re

import gspread
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from google.oauth2.service_account import Credentials
from plotly.subplots import make_subplots

LOGO = "logo.png" if os.path.exists("logo.png") else None
st.set_page_config(page_title='Dashboard OTD "Axiom" RES GDL', page_icon=LOGO or "📦", layout="wide")

NINGUNA = "(ninguna)"
CATS_DEFECTO = {"racks", "shelves", "peripherals", "tray"}
FIJAS = {"Date", "CW", "Year", "Total", "Vol", "Target", "Nota"}
COLORES = ["#1f4e79", "#e07b39", "#2e6b2e", "#2a9bd1", "#8e44ad", "#c0392b"]
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]
F_PUB = "Datos publicados (equipo)"
F_UP = "Subir mi Excel (administrador)"
F_EJ = "Datos de ejemplo"
REGLA_PROM = "Promedio del OTD por línea"
REGLA_PCT = "% de líneas que cumplen la meta"
COLS_CUENTAS = ["Cuenta", "Date", "PM", "LOB", "Item", "Familia", "Commit", "Shipment", "Target"]


def secreto(nombre):
    try:
        return st.secrets.get(nombre)
    except Exception:
        return None


# ---------- Contraseña opcional ----------
clave = secreto("password")
if clave:
    if st.text_input("Contraseña", type="password") != clave:
        st.stop()


# ======================================================================
# Utilidades de datos
# ======================================================================
def a_fraccion(serie):
    texto = serie.astype(str).str.replace("%", "", regex=False)
    s = pd.to_numeric(texto.str.replace(",", ".", regex=False).str.strip(), errors="coerce")
    return s.where(s <= 1.5, s / 100)   # un 98.5 pasa a 0.985, valor por valor


def semana_excel16(fechas):
    """Misma numeración que WEEKNUM(fecha, 16) de Excel: semanas sábado a viernes."""
    f = pd.to_datetime(fechas)
    ene1 = pd.to_datetime(f.dt.year.astype(str) + "-01-01")
    desfase = (ene1.dt.weekday - 5) % 7
    return ((f.dt.dayofyear - 1 + desfase) // 7 + 1).astype(int)


def unificar_lob(serie):
    """Junta variantes de escritura (mayúsculas, espacios dobles) de un mismo LOB."""
    s = serie.fillna("").astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
    clave_lob = s.str.lower()
    mapa = {}
    for k, grupo in s.groupby(clave_lob):
        conteo = grupo.value_counts()
        preferidos = [v for v in conteo.index if not (v.isupper() and len(v) > 4)]
        mapa[k] = (preferidos or list(conteo.index))[0]
    return clave_lob.map(mapa)


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
    df["Nota"] = ""
    df.loc[30, "Nota"] = "shutdown week"
    return df


def cuentas_ejemplo():
    rng = np.random.default_rng(11)
    filas = []
    for cuenta, lobs in [("CIS", ["Mckenzie", "Churchill", "N9300"]), ("IBM", ["Proyecto A", "Proyecto B"])]:
        items = [(f"{cuenta}H-{int(rng.integers(100000, 999999))}-FG", lobs[i % len(lobs)]) for i in range(8)]
        for f in pd.date_range("2026-01-05", periods=38, freq="W-MON"):
            for item, lob in items:
                if rng.random() < 0.7:
                    commit = int(rng.choice([50, 100, 200, 500, 1000]))
                    ship = int(commit * rng.choice([0, 0.5, 0.9, 1, 1, 1, 1.1, 1.3]))
                    filas.append({"Cuenta": cuenta, "Date": f, "PM": "PM 1", "LOB": lob, "Item": item,
                                  "Familia": "Familia " + lob[:1], "Commit": commit, "Shipment": ship,
                                  "Target": 0.95})
    return pd.DataFrame(filas)


def leer_cuentas_xls(xls, fam_map):
    """Detecta las pestañas de cuenta (nombre con 'OTD' y columnas Commit y Flex Item FG)."""
    partes, resumen = [], []
    for nombre in xls.sheet_names:
        if "otd" not in nombre.lower():
            continue
        df = xls.parse(nombre)
        df.columns = [str(c).strip() for c in df.columns]
        mapa = {c.lower(): c for c in df.columns}
        if not {"commit", "flex item fg"} <= set(mapa):
            continue
        cuenta = re.sub(r"^\s*otd\s+", "", nombre.strip(), flags=re.I).strip() or nombre.strip()

        def col(clave, defecto=None):
            return df[mapa[clave]] if clave in mapa else pd.Series(defecto, index=df.index)

        t = pd.DataFrame({
            "Cuenta": cuenta,
            "Date": pd.to_datetime(col("date"), errors="coerce"),
            "PM": col("pm", "").fillna("").astype(str).str.strip(),
            "LOB": col("lob", "").fillna("").astype(str),
            "Item": col("flex item fg", "").fillna("").astype(str).str.strip(),
            "Commit": pd.to_numeric(col("commit"), errors="coerce"),
            "Shipment": pd.to_numeric(col("shipment"), errors="coerce"),
            "Target": a_fraccion(col("otd target", 0.95)),
        })
        t = t[(t["Item"] != "") & t["Date"].notna() & (t["Commit"].notna() | t["Shipment"].notna())].copy()
        t["Target"] = t["Target"].fillna(0.95)
        t["Familia"] = t["Item"].map(fam_map).fillna("")
        resumen.append((cuenta, nombre, len(t)))
        if len(t):
            partes.append(t)
    if not partes:
        return None, resumen
    C = pd.concat(partes, ignore_index=True)
    C["LOB"] = unificar_lob(C["LOB"])
    return C[COLS_CUENTAS], resumen


def preparar_lineas(C):
    """Agrega topado, OTD por línea y semana. Replica las fórmulas de tus columnas G, H y J."""
    C = C.copy()
    C["Date"] = pd.to_datetime(C["Date"], errors="coerce")
    C = C.dropna(subset=["Date"]).copy()
    for c in ["Commit", "Shipment", "Target"]:
        C[c] = pd.to_numeric(C[c], errors="coerce")
    C["Commit"] = C["Commit"].fillna(0)
    C["Shipment"] = C["Shipment"].fillna(0)
    C["Target"] = C["Target"].fillna(0.95)
    for c in ["Cuenta", "PM", "LOB", "Item", "Familia"]:
        C[c] = C[c].fillna("").astype(str) if c in C.columns else ""
    C["topado"] = np.minimum(C["Commit"], C["Shipment"])
    C["otd_linea"] = np.where(C["Commit"] > 0, C["topado"] / C["Commit"].where(C["Commit"] > 0, 1), np.nan)
    C["sem"] = semana_excel16(C["Date"])
    C["año"] = C["Date"].dt.year
    return C


def cuentas_publicadas(df):
    if df is None or df.empty or "Cuenta" not in df.columns:
        return None
    t = df.copy()
    for c in COLS_CUENTAS:
        if c not in t.columns:
            t[c] = ""
    return t[COLS_CUENTAS]


# ======================================================================
# Datos publicados en Google Sheets (se guardan como texto)
# ======================================================================
@st.cache_resource
def libro_otd():
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=SCOPES)
    return gspread.authorize(creds).open_by_key(st.secrets["otd_sheet_id"])


def hoja(nombre):
    lb = libro_otd()
    try:
        return lb.worksheet(nombre)
    except gspread.WorksheetNotFound:
        return lb.add_worksheet(nombre, rows=1000, cols=20)


def leer_hoja(nombre, tolerante=False):
    try:
        registros = hoja(nombre).get_all_records(value_render_option="UNFORMATTED_VALUE",
                                                 numericise_ignore=["all"])
        return pd.DataFrame(registros)
    except Exception:
        if tolerante:      # la pestaña de cuentas puede no existir o estar vacía todavía
            return pd.DataFrame()
        raise


@st.cache_data(ttl=60, show_spinner="Cargando datos publicados...")
def leer_publicados():
    datos = leer_hoja("datos")
    cuentas = leer_hoja("cuentas", tolerante=True)
    publicado = hoja("meta").acell("B1").value
    return datos, cuentas, publicado


def tabla_publicable(dfull, categorias):
    t = pd.DataFrame({"Date": dfull["fecha"].dt.strftime("%Y-%m-%d"),
                      "CW": dfull["sem"], "Year": dfull["año"], "Total": dfull["total"]})
    for c in categorias:
        t[c] = dfull[c]
    t["Vol"] = dfull["vol"]
    t["Target"] = dfull["target"]
    t["Nota"] = dfull["nota"]
    return t


def tabla_cuentas_publicable(C):
    t = C.copy()
    t["Date"] = pd.to_datetime(t["Date"]).dt.strftime("%Y-%m-%d")
    return t[COLS_CUENTAS]


def a_texto(t, no_numericas):
    t = t.copy()
    for c in t.columns:
        if c not in no_numericas:
            t[c] = t[c].map(lambda v: "" if pd.isna(v) else repr(float(v)))
    return t.fillna("").astype(str)


def escribir(nombre, t):
    valores = [list(t.columns)] + t.values.tolist()
    ws = hoja(nombre)
    ws.clear()
    ws.resize(rows=len(valores) + 20, cols=len(valores[0]) + 2)
    ws.update(values=valores, range_name="A1", value_input_option="RAW")


def publicar(t_datos, t_cuentas):
    escribir("datos", a_texto(t_datos, {"Date", "Nota"}))
    if t_cuentas is not None and len(t_cuentas):
        escribir("cuentas", a_texto(t_cuentas, {"Cuenta", "Date", "PM", "LOB", "Item", "Familia"}))
    ahora = pd.Timestamp.now(tz="America/Mexico_City").strftime("%d/%m/%Y %H:%M")
    hoja("meta").update(values=[["publicado", ahora]], range_name="A1", value_input_option="RAW")
    leer_publicados.clear()


# ======================================================================
# Título
# ======================================================================
if LOGO:
    c_logo, c_tit = st.columns([1, 8], vertical_alignment="center")
    c_logo.image(LOGO, width=90)
    c_tit.title('Dashboard OTD "Axiom" RES GDL')
else:
    st.title('Dashboard OTD "Axiom" RES GDL')

# ======================================================================
# Fuente de datos
# ======================================================================
sheets_cfg = bool(secreto("otd_sheet_id")) and bool(secreto("gcp_service_account"))
opciones_fuente = ([F_PUB] if sheets_cfg else []) + [F_UP, F_EJ]
fuente = st.radio("Fuente de datos", opciones_fuente, horizontal=True)

C_raw = None
if fuente == F_EJ:
    df_raw = datos_ejemplo()
    C_raw = cuentas_ejemplo()
elif fuente == F_PUB:
    try:
        df_raw, cuentas_pub, publicado = leer_publicados()
    except Exception as e:
        st.error(f"No pude leer los datos publicados: {e}")
        st.stop()
    if df_raw.empty:
        st.info("Todavía no hay datos publicados. Pide al administrador que publique el Excel.")
        st.stop()
    C_raw = cuentas_publicadas(cuentas_pub)
    st.caption(f"Datos publicados el {publicado or 'fecha desconocida'}")
    if st.button("🔄 Actualizar datos ahora"):
        leer_publicados.clear()
        st.rerun()
else:
    admin = secreto("admin_password")
    if admin and st.text_input("Contraseña de administrador", type="password", key="adm") != admin:
        st.info("Escribe la contraseña de administrador para subir y publicar datos.")
        st.stop()
    archivo = st.file_uploader("Sube tu archivo (.xlsx)", type=["xlsx"])
    if archivo is None:
        st.info("Sube tu Excel. Puede traer la hoja del OTD acumulado y una hoja por cuenta "
                "(con 'OTD' en el nombre y las columnas Date, PM, LOB, Flex Item FG, Commit y Shipment).")
        st.stop()
    xls = pd.ExcelFile(archivo)
    nombres = xls.sheet_names
    idx_acum = next((i for i, n in enumerate(nombres)
                     if any(k in n.lower() for k in ("comulativo", "acumulado", "cumulative"))), 0)
    nombre_hoja = st.selectbox("Hoja del OTD acumulado", nombres, index=idx_acum)
    df_raw = xls.parse(nombre_hoja)

    fam_map = {}
    for n in nombres:
        if "matrix" in n.lower():
            mt = xls.parse(n)
            mt.columns = [str(c).strip() for c in mt.columns]
            if {"Flex ITEM", "Family"} <= set(mt.columns):
                mt = mt.dropna(subset=["Flex ITEM"]).copy()
                mt["Flex ITEM"] = mt["Flex ITEM"].astype(str).str.strip()
                fam_map = dict(mt.drop_duplicates("Flex ITEM")[["Flex ITEM", "Family"]].values)
    C_raw, resumen_cuentas = leer_cuentas_xls(xls, fam_map)
    if resumen_cuentas:
        st.caption("Cuentas detectadas: " + " · ".join(
            f"{c} ({n} líneas)" if n else f"{c} (sin datos)" for c, _, n in resumen_cuentas))
    else:
        st.caption("No detecté pestañas de cuenta en este archivo.")

with st.expander("🔍 Ver cómo leí tu archivo"):
    st.write(f"{len(df_raw)} filas y {len(df_raw.columns)} columnas (hoja del OTD acumulado)")
    st.write("Primeras filas")
    st.dataframe(df_raw.head(5))
    st.write("Últimas filas")
    st.dataframe(df_raw.tail(8))

cols = list(df_raw.columns)
opc = [NINGUNA] + cols


def pick(candidatos, opcional=False):
    for i, c in enumerate(cols):
        if str(c).strip().lower() in candidatos:
            return i + 1 if opcional else i
    return 0


# ======================================================================
# Panel lateral
# ======================================================================
with st.sidebar:
    st.subheader("OTD por cuenta")
    incluir_curso = st.checkbox("Incluir la semana en curso", value=False)
    sin_topar = st.checkbox("Volumétrico sin topar (embarcado ÷ comprometido)", value=False)
    regla_bin = st.radio("OTD binario cuenta como", [REGLA_PROM, REGLA_PCT])
    st.subheader("Columnas del OTD acumulado")
    c_fecha = st.selectbox("Fecha", cols, index=pick({"date", "fecha"}))
    c_total = st.selectbox("OTD total", cols, index=pick({"total res gdl", "total", "total otd", "otd"}))
    c_vol = st.selectbox("Volumen", opc, index=pick({"vol", "volume", "volumen"}, True))
    c_sem = st.selectbox("Semana (CW)", opc, index=pick({"cw", "semana", "week", "wk"}, True))
    c_anio = st.selectbox("Año", opc, index=pick({"year", "año", "anio"}, True))
    c_target = st.selectbox("Meta (target)", opc, index=pick({"target", "meta"}, True))
    c_nota = st.selectbox("Notas", opc, index=pick({"column1", "notas", "nota", "notes", "comentarios"}, True))
    if fuente == F_PUB:
        defecto_cats = [c for c in cols if c not in FIJAS]
    else:
        defecto_cats = [c for c in cols if str(c).strip().lower() in CATS_DEFECTO]
    categorias = st.multiselect("Categorías (columnas de OTD)", cols, default=defecto_cats)
    meta_fija = st.number_input("Meta fija (%) si no hay columna de meta", value=95.0, step=0.5) / 100
    st.subheader("Opciones")
    eje = st.radio("Eje X", ["Semana (CW)", "Fecha"])
    ponderado = st.checkbox("Promedio ponderado por volumen", value=False)
    excluir = st.checkbox("Excluir semanas con nota (ej. paro)", value=False)
    etiquetas = st.checkbox("Mostrar etiquetas de datos", value=True)

# ======================================================================
# Limpieza del OTD acumulado
# ======================================================================
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
    d = None
else:
    d["año"] = d["año"].astype(int)
    d["sem"] = d["sem"].astype(int)
    d = d.sort_values(["año", "sem"]).reset_index(drop=True)

C = preparar_lineas(C_raw) if C_raw is not None and len(C_raw) else None

# ======================================================================
# Revisión de calidad de datos
# ======================================================================
problemas = []
if d is not None:
    for i, v in d["total"][(d["total"] < 0.5) | (d["total"] > 1.005)].items():
        problemas.append(f"OTD acumulado, CW{d.loc[i, 'sem']} ({d.loc[i, 'año']}): OTD total = {v:.1%}")
    for c in categorias:
        for i, v in d[c][(d[c] < 0.5) | (d[c] > 1.005)].items():
            problemas.append(f"OTD acumulado, CW{d.loc[i, 'sem']} ({d.loc[i, 'año']}): {c} = {v:.1%}")
    for r in d[d["fecha"].notna() & (d["fecha"].dt.weekday != 0)].itertuples():
        problemas.append(f"OTD acumulado, CW{r.sem} ({r.año}): la fecha {r.fecha:%d/%m/%Y} no es lunes")
    for a, s in d[d.duplicated(["año", "sem"], keep=False)][["año", "sem"]].drop_duplicates().itertuples(index=False):
        problemas.append(f"OTD acumulado, semana repetida: CW{s} ({a})")
    for a, g in d.groupby("año"):
        faltan = sorted(set(range(int(g["sem"].min()), int(g["sem"].max()) + 1)) - set(g["sem"]))
        if faltan:
            problemas.append(f"OTD acumulado, faltan semanas en {a}: " + ", ".join(f"CW{s}" for s in faltan))
if C is not None:
    for cuenta, g in C.groupby("Cuenta"):
        sin_commit = int((g["Commit"] <= 0).sum())
        if sin_commit:
            problemas.append(f"{cuenta}: {sin_commit} líneas sin Commit (no cuentan en el OTD binario)")
        no_lunes = int((g["Date"].dt.weekday != 0).sum())
        if no_lunes:
            problemas.append(f"{cuenta}: {no_lunes} líneas con fecha que no es lunes")
if problemas:
    with st.expander(f"⚠️ {len(problemas)} posibles problemas en los datos", expanded=False):
        for p in problemas[:40]:
            st.write("• " + p)

# ---------- Publicar (solo administrador) ----------
if fuente == F_UP and sheets_cfg and d is not None:
    if st.button("📤 Publicar estos datos para el equipo"):
        try:
            t_c = tabla_cuentas_publicable(C_raw) if C_raw is not None and len(C_raw) else None
            publicar(tabla_publicable(d, categorias), t_c)
            n_c = 0 if t_c is None else t_c["Cuenta"].nunique()
            st.success(f"Publicado: {len(d)} semanas del OTD acumulado y {n_c} cuentas con datos. "
                       f"El equipo ya las ve en '{F_PUB}'.")
        except Exception as e:
            st.error(f"No pude publicar: {e}")

# ---------- Semana en curso ----------
if d is not None:
    try:
        iso = pd.Timestamp.now(tz="America/Mexico_City").isocalendar()
        sem_hoy = int(iso[1])
        lunes_hoy = dt.date.fromisocalendar(int(iso[0]), sem_hoy, 1)
        ult_fila = d.iloc[-1]
        lunes_ult = dt.date.fromisocalendar(int(ult_fila["año"]), int(ult_fila["sem"]), 1)
        atraso = (lunes_hoy - lunes_ult).days // 7
        resumen = (f"Semana en curso: CW{sem_hoy} · última semana con datos del OTD acumulado: "
                   f"CW{int(ult_fila['sem'])} ({int(ult_fila['año'])})")
        if atraso <= 1:
            st.success("✅ Datos al día. " + resumen)
        else:
            st.warning(f"⚠️ Faltan datos ({atraso} semanas de atraso). " + resumen)
    except Exception:
        pass


# ======================================================================
# Vista 1: OTD acumulado (la que ya tenías)
# ======================================================================
def mostrar_acumulado(d):
    if d is None:
        st.warning("No encontré filas válidas en la hoja del OTD acumulado. "
                   "Revisa las columnas del panel lateral.")
        return
    años = sorted(d["año"].unique())
    sel_años = st.multiselect("Año", años, default=años, key="ac_años")
    d = d[d["año"].isin(sel_años)].reset_index(drop=True)
    if excluir:
        d = d[d["nota"] == ""].reset_index(drop=True)
    if d.empty:
        st.warning("No hay datos con esos filtros.")
        return

    if eje == "Fecha" and d["fecha"].notna().all():
        d["x"] = d["fecha"].dt.strftime("%d-%b")
    elif len(sel_años) > 1:
        d["x"] = d["año"].astype(str) + "-S" + d["sem"].astype(str).str.zfill(2)
    else:
        d["x"] = d["sem"].astype(str)

    if len(d) > 1:
        i0, i1 = st.select_slider("Rango de periodos", options=list(range(len(d))),
                                  value=(0, len(d) - 1), format_func=lambda i: d["x"].iloc[i],
                                  key="ac_rango")
        d = d.iloc[i0:i1 + 1]

    sel_cats = st.multiselect("Categorías a mostrar", categorias, default=categorias, key="ac_cats")

    def prom(serie):
        s = serie.dropna()
        if s.empty:
            return float("nan")
        if ponderado and d["vol"].notna().any():
            w = d.loc[s.index, "vol"].fillna(0)
            if w.sum() > 0:
                return float((s * w).sum() / w.sum())
        return float(s.mean())

    meta = float(d["target"].iloc[-1])
    otd = prom(d["total"])
    ult = d.iloc[-1]
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("OTD promedio", f"{otd:.2%}", f"{(otd - meta) * 100:+.1f} pts vs meta")
    k2.metric("Volumen total", f"{d['vol'].sum():,.0f}" if d["vol"].notna().any() else "—")
    k3.metric("Meta", f"{meta:.0%}")
    k4.metric("Periodos bajo la meta", f"{int((d['total'] < d['target']).sum())} de {len(d)}")
    k5.metric(f"Última ({ult['x']})", f"{ult['total']:.1%}",
              f"{(ult['total'] - ult['target']) * 100:+.1f} pts vs meta")

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
    st.plotly_chart(fig, use_container_width=True, key="ac_fig1")

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
        st.plotly_chart(fig2, use_container_width=True, key="ac_fig2")

    st.subheader("Detalle")
    tabla = d[["fecha", "año", "sem", "total"] + sel_cats + ["vol", "target", "nota"]].rename(
        columns={"fecha": "Fecha", "año": "Año", "sem": "Semana", "total": "OTD total",
                 "vol": "Volumen", "target": "Meta", "nota": "Nota"})
    config = {c: st.column_config.NumberColumn(format="percent") for c in ["OTD total", "Meta"] + sel_cats}
    st.dataframe(tabla, hide_index=True, column_config=config)
    buf = io.BytesIO()
    tabla.to_excel(buf, index=False)
    st.download_button("⬇️ Descargar a Excel", buf.getvalue(), file_name="otd_filtrado.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="ac_dl")


# ======================================================================
# Vista 2: OTD por cuenta (binario y volumétrico)
# ======================================================================
def metricas(L):
    com = float(L["Commit"].sum())
    ship = float(L["Shipment"].sum())
    top = float(L["topado"].sum())
    act = ship if sin_topar else top
    vol = act / com if com > 0 else float("nan")
    lin = L[L["Commit"] > 0]
    if len(lin) == 0:
        binario = float("nan")
    elif regla_bin == REGLA_PROM:
        binario = float(lin["otd_linea"].mean())
    else:
        binario = float((lin["otd_linea"] >= lin["Target"]).mean())
    bajo = int((lin["otd_linea"] < lin["Target"]).sum())
    return dict(com=com, act=act, vol=vol, binario=binario, lineas=len(lin), bajo=bajo)


def semanal(L):
    filas = []
    for (a, s), g in L.groupby(["año", "sem"]):
        m = metricas(g)
        m.update({"año": a, "sem": s})
        filas.append(m)
    W = pd.DataFrame(filas)
    if W["año"].nunique() == 1:
        W["x"] = W["sem"].astype(str)
    else:
        W["x"] = W["año"].astype(str) + "-S" + W["sem"].astype(str).str.zfill(2)
    return W


def anillo(valor, meta, titulo):
    sin_dato = pd.isna(valor)
    v = 0.0 if sin_dato else max(0.0, min(float(valor), 1.0))
    color = "#2e9e4f" if (not sin_dato and valor >= meta) else "#d64545"
    fig = go.Figure(go.Pie(values=[v, 1 - v], hole=0.72, sort=False, direction="clockwise",
                           marker=dict(colors=[color, "#e6e6e6"]), textinfo="none", hoverinfo="skip"))
    fig.update_layout(
        title=dict(text=titulo, x=0.5), showlegend=False, height=250,
        margin=dict(l=10, r=10, t=50, b=10),
        annotations=[dict(text="<b>—</b>" if sin_dato else f"<b>{valor:.1%}</b>", x=0.5, y=0.53,
                          font=dict(size=26), showarrow=False),
                     dict(text=f"meta {meta:.0%}", x=0.5, y=0.38, font=dict(size=12, color="gray"),
                          showarrow=False)])
    return fig


def barras_por(L, campo, titulo, meta, key):
    filas = []
    for k, g in L.groupby(campo):
        m = metricas(g)
        filas.append({"nombre": k if k != "" else "(sin dato)", "vol": m["vol"], "com": m["com"]})
    B = pd.DataFrame(filas).dropna(subset=["vol"])
    if B.empty:
        return
    B = B.sort_values("vol", ascending=False)   # lo peor queda arriba
    fig = go.Figure(go.Bar(
        x=B["vol"], y=B["nombre"], orientation="h",
        marker_color=["#2e9e4f" if v >= meta else "#d64545" for v in B["vol"]],
        text=[f"{v:.1%}" for v in B["vol"]], textposition="outside"))
    fig.add_vline(x=meta, line_dash="dash", line_color="black")
    fig.update_layout(title=titulo, height=min(700, 130 + 34 * len(B)),
                      xaxis=dict(tickformat=".0%", range=[0, max(1.1, float(B["vol"].max()) + 0.1)]),
                      margin=dict(l=10, r=30, t=50, b=10))
    st.plotly_chart(fig, use_container_width=True, key=key)


def tabla_items(L, meta):
    filas = []
    for item, g in L.groupby("Item"):
        m = metricas(g)
        lin = g[g["Commit"] > 0]
        filas.append({"Número de parte": item, "LOB": g["LOB"].iloc[-1], "Familia": g["Familia"].iloc[-1],
                      "Commit": m["com"], "Actuals": m["act"], "OTD volumétrico": m["vol"],
                      "OTD binario": m["binario"], "Semanas bajo la meta": int((lin["otd_linea"] < meta).sum())})
    T = pd.DataFrame(filas)
    return T.sort_values("OTD volumétrico", na_position="last").reset_index(drop=True)


def mostrar_cuentas(C):
    if C is None or len(C) == 0:
        st.info("Todavía no hay datos por cuenta. El administrador debe subir el Excel con las "
                "pestañas de cada cuenta y publicarlo.")
        return
    cuenta = st.selectbox("Cuenta", sorted(C["Cuenta"].unique()), key="ct_cuenta")
    L = C[C["Cuenta"] == cuenta].copy()

    hoy = pd.Timestamp.now(tz="America/Mexico_City").tz_localize(None).normalize()
    lunes_actual = hoy - pd.Timedelta(days=int(hoy.weekday()))
    en_curso = L["Date"] >= lunes_actual
    if not incluir_curso and en_curso.any():
        st.caption(f"Semana en curso (CW{int(L.loc[en_curso, 'sem'].max())}) excluida del cálculo. "
                   "Actívala en el panel lateral para incluirla.")
        L = L[~en_curso].copy()
    if L.empty:
        st.warning("No hay datos de esta cuenta en semanas cerradas.")
        return

    with st.expander("Filtros", expanded=True):
        años = sorted(L["año"].unique())
        sel_años = st.multiselect("Año", años, default=años, key=f"ct_años_{cuenta}")
        if sel_años:
            L = L[L["año"].isin(sel_años)]
        for campo, etiqueta in [("LOB", "Proyecto (LOB)"), ("PM", "PM"), ("Familia", "Familia")]:
            opciones = sorted(x for x in L[campo].unique() if x != "")
            if opciones:
                sel = st.multiselect(f"{etiqueta} (vacío = todos)", opciones, key=f"ct_{campo}_{cuenta}")
                if sel:
                    L = L[L[campo].isin(sel)]
        sel_items = st.multiselect("Número de parte (vacío = todos)", sorted(L["Item"].unique()),
                                   key=f"ct_item_{cuenta}")
        if sel_items:
            L = L[L["Item"].isin(sel_items)]
        L = L.copy()
        L["pk"] = L["año"] * 100 + L["sem"]
        claves = sorted(L["pk"].unique())
        if len(claves) > 1:
            def etiqueta_sem(k):
                return f"CW{k % 100}" if len(sel_años) <= 1 else f"{k // 100}-S{k % 100:02d}"
            i0, i1 = st.select_slider("Rango de semanas", options=list(range(len(claves))),
                                      value=(0, len(claves) - 1),
                                      format_func=lambda i: etiqueta_sem(claves[i]),
                                      key=f"ct_rango_{cuenta}_{len(claves)}")
            L = L[(L["pk"] >= claves[i0]) & (L["pk"] <= claves[i1])]
    if L.empty:
        st.warning("No hay datos con esos filtros.")
        return

    meta = float(L["Target"].median())
    tot = metricas(L)

    a, b = st.columns(2)
    a.plotly_chart(anillo(tot["binario"], meta, "OTD binario"), use_container_width=True,
                   key=f"ct_an1_{cuenta}")
    b.plotly_chart(anillo(tot["vol"], meta, "OTD volumétrico"), use_container_width=True,
                   key=f"ct_an2_{cuenta}")
    st.caption(f"Binario: {regla_bin.lower()}. Volumétrico: "
               f"{'embarcado' if sin_topar else 'embarcado topado al commit'} ÷ comprometido.")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Commit", f"{tot['com']:,.0f}")
    k2.metric("Actuals" + (" (sin topar)" if sin_topar else " (topado)"), f"{tot['act']:,.0f}")
    k3.metric("Líneas con commit", f"{tot['lineas']:,}")
    k4.metric("Líneas bajo la meta", f"{tot['bajo']:,}")

    W = semanal(L)
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(x=W["x"], y=W["com"], name="Commit", marker_color="#1f4e79"), secondary_y=False)
    fig.add_trace(go.Bar(x=W["x"], y=W["act"], name="Actuals" + (" (sin topar)" if sin_topar else " (topado)"),
                         marker_color="#90d36b"), secondary_y=False)
    modo = "lines+markers+text" if etiquetas else "lines+markers"
    fig.add_trace(go.Scatter(x=W["x"], y=W["binario"], name="OTD binario", mode=modo,
                             line=dict(color="#e07b39", width=2),
                             text=W["binario"].map(lambda v: "" if pd.isna(v) else f"{v:.0%}"),
                             textposition="top center"), secondary_y=True)
    fig.add_trace(go.Scatter(x=W["x"], y=W["vol"], name="OTD volumétrico", mode=modo,
                             line=dict(color="#2a9bd1", width=2),
                             text=W["vol"].map(lambda v: "" if pd.isna(v) else f"{v:.0%}"),
                             textposition="bottom center"), secondary_y=True)
    fig.add_trace(go.Scatter(x=W["x"], y=[meta] * len(W), name="Meta", mode="lines",
                             line=dict(color="black", dash="dash")), secondary_y=True)
    valores = W[["binario", "vol"]].to_numpy(dtype=float)
    piso = max(0.0, float(np.nanmin(np.append(valores, meta))) - 0.05)
    techo = max(1.05, float(np.nanmax(valores)) + 0.05)
    fig.update_yaxes(tickformat=".0%", range=[piso, techo], secondary_y=True)
    fig.update_yaxes(showgrid=False, title_text="Piezas", secondary_y=False,
                     range=[0, float(max(W["com"].max(), W["act"].max())) * 1.25])
    fig.update_layout(title=f"{cuenta}: Commit, Actuals y OTD por semana", height=500, barmode="group",
                      xaxis=dict(type="category"), legend=dict(orientation="h", y=-0.2),
                      margin=dict(l=10, r=10, t=50, b=10))
    st.plotly_chart(fig, use_container_width=True, key=f"ct_fig1_{cuenta}")

    barras_por(L, "LOB", "OTD volumétrico por proyecto (LOB)", meta, f"ct_lob_fig_{cuenta}")
    if (L["Familia"] != "").any():
        barras_por(L, "Familia", "OTD volumétrico por familia", meta, f"ct_fam_fig_{cuenta}")

    st.subheader("Detalle por número de parte")
    T = tabla_items(L, meta)
    solo_bajo = st.checkbox("Mostrar solo los que están bajo la meta (volumétrico)", key=f"ct_bajo_{cuenta}")
    if solo_bajo:
        T = T[T["OTD volumétrico"] < meta]
    cfg = {"OTD volumétrico": st.column_config.NumberColumn(format="percent"),
           "OTD binario": st.column_config.NumberColumn(format="percent"),
           "Commit": st.column_config.NumberColumn(format="%d"),
           "Actuals": st.column_config.NumberColumn(format="%d")}
    st.dataframe(T, hide_index=True, column_config=cfg)
    buf = io.BytesIO()
    T.to_excel(buf, index=False)
    st.download_button("⬇️ Descargar a Excel", buf.getvalue(), file_name=f"otd_{cuenta}_por_parte.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key=f"ct_dl_{cuenta}")


# ======================================================================
# Pestañas
# ======================================================================
tab_acum, tab_cta = st.tabs(["📈 OTD acumulado", "🏢 OTD por cuenta"])
with tab_acum:
    mostrar_acumulado(d)
with tab_cta:
    mostrar_cuentas(C)
