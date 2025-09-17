# =========================
# Pareto Comunidad – MSP (mínima, rápida y funcional)
# =========================

import re
import unicodedata
from io import BytesIO
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Pareto Comunidad – MSP", layout="wide", initial_sidebar_state="collapsed")

# ---------- Intentamos importar plotly; si falla, usamos fallback ----------
try:
    import plotly.graph_objects as go
    HAS_PLOTLY = True
except Exception:
    HAS_PLOTLY = False

TOP_N_GRAFICO = 30  # para que el gráfico no pese; la descarga incluye todo

# ---------- Catálogo EMBEBIDO (puedes ampliarlo luego) ----------
CATALOGO_BASE = [
    ("HURTO", "DELITOS CONTRA LA PROPIEDAD", "HURTO"),
    ("ROBO", "DELITOS CONTRA LA PROPIEDAD", "ROBO"),
    ("DAÑOS A LA PROPIEDAD", "DELITOS CONTRA LA PROPIEDAD", "DAÑOS A LA PROPIEDAD"),
    ("ASALTO", "DELITOS CONTRA LA PROPIEDAD", "ASALTO"),
    ("VENTA DE DROGAS", "DROGAS", "VENTA DE DROGAS"),
    ("TRAFICO DE DROGAS", "DROGAS", "TRÁFICO DE DROGAS"),
    ("MICROTRAFICO", "DROGAS", "MICROTRÁFICO"),
    ("CONSUMO DE DROGAS", "DROGAS", "CONSUMO DE DROGAS"),
    ("BUNKER", "DROGAS", "BÚNKER"),
    ("PUNTO DE VENTA", "DROGAS", "PUNTO DE VENTA"),
    ("CONSUMO DE ALCOHOL", "ALCOHOL", "CONSUMO DE ALCOHOL EN VÍA PÚBLICA"),
    ("HOMICIDIOS", "DELITOS CONTRA LA VIDA", "HOMICIDIOS"),
    ("HERIDOS", "DELITOS CONTRA LA VIDA", "HERIDOS"),
    ("TENTATIVA DE HOMICIDIO", "DELITOS CONTRA LA VIDA", "TENTATIVA DE HOMICIDIO"),
    ("VIOLENCIA DOMESTICA", "VIOLENCIA", "VIOLENCIA DOMÉSTICA"),
    ("AGRESION", "VIOLENCIA", "AGRESIÓN"),
    ("ABUSO SEXUAL", "VIOLENCIA", "ABUSO SEXUAL"),
    ("VIOLACION", "VIOLENCIA", "VIOLACIÓN"),
    ("ACOSO SEXUAL CALLEJERO", "RIESGO SOCIAL", "ACOSO SEXUAL CALLEJERO"),
    ("ACOSO ESCOLAR", "RIESGO SOCIAL", "ACOSO ESCOLAR (BULLYING)"),
    ("ACTOS OBSCENOS", "RIESGO SOCIAL", "ACTOS OBSCENOS EN VIA PUBLICA"),
    ("PANDILLAS", "ORDEN PÚBLICO", "PANDILLAS"),
    ("INDIGENCIA", "ORDEN PÚBLICO", "INDIGENCIA"),
    ("VAGANCIA", "ORDEN PÚBLICO", "VAGANCIA"),
    ("RUIDO", "ORDEN PÚBLICO", "CONTAMINACIÓN SONORA"),
    ("CARRERAS ILEGALES", "ORDEN PÚBLICO", "CARRERAS ILEGALES"),
    ("ARMAS BLANCAS", "ORDEN PÚBLICO", "PORTACIÓN DE ARMA BLANCA"),
]

# ---------- Utils ----------
def strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in s if not unicodedata.combining(ch))

def norm(s: Optional[str]) -> str:
    if s is None:
        return ""
    if not isinstance(s, str):
        s = str(s)
    s = strip_accents(s).strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df2 = df.copy()
    df2.columns = [norm(c) for c in df2.columns]
    return df2

@st.cache_data(show_spinner=False)
def read_plantilla_matriz(file) -> pd.DataFrame:
    # lee solo la hoja 'matriz'
    return pd.read_excel(file, sheet_name="matriz", engine="openpyxl")

def build_catalogo_df() -> pd.DataFrame:
    return pd.DataFrame(CATALOGO_BASE, columns=["NOMBRE CORTO", "CATEGORÍA", "DESCRIPTOR"])

def build_keyword_maps(desc_df: pd.DataFrame) -> Tuple[Dict[str, str], Dict[str, str]]:
    # mapeos normalizados para HEADERS
    by_desc_norm = {}
    by_nc_norm = {}
    for _, r in desc_df.iterrows():
        desc = str(r["DESCRIPTOR"]).strip()
        nc = str(r["NOMBRE CORTO"]).strip()
        if desc:
            by_desc_norm[norm(desc)] = desc
        if nc:
            by_nc_norm[norm(nc)] = desc
    return by_desc_norm, by_nc_norm

def header_marked_series(s: pd.Series) -> pd.Series:
    # marcado rápido: num != 0 o texto no vacío/ni "no"/"0"
    s2 = s.copy()
    num = pd.to_numeric(s2, errors="coerce").fillna(0) != 0
    txt = s2.astype(str).str.strip().str.lower().apply(strip_accents)
    txtmask = ~txt.isin(["", "no", "0", "nan", "none", "false"])
    return num | txtmask

def detect_by_headers(df_raw: pd.DataFrame, by_desc_norm: Dict[str, str], by_nc_norm: Dict[str, str]) -> List[str]:
    df = normalize_columns(df_raw)
    hits: List[str] = []

    # mapear encabezados → descriptor (match por igualdad o substring)
    def header_to_descriptor(ncol: str) -> Optional[str]:
        for k in by_desc_norm:
            if k == ncol or k in ncol or ncol in k:
                return by_desc_norm[k]
        for k in by_nc_norm:
            if k == ncol or k in ncol or ncol in k:
                return by_nc_norm[k]
        return None

    header_map = {}
    for col in df.columns:
        d = header_to_descriptor(col)
        if d:
            header_map[col] = d

    # contar por columnas mapeadas
    for col, desc in header_map.items():
        try:
            m = header_marked_series(df[col])
            c = int(m.sum())
            if c > 0:
                hits.extend([desc] * c)
        except Exception:
            continue

    return hits

def make_copilado(hits: List[str]) -> pd.DataFrame:
    if not hits:
        return pd.DataFrame({"Descriptor": [], "Frecuencia": []})
    s = pd.Series(hits, name="Descriptor")
    df = s.value_counts(dropna=False).rename_axis("Descriptor").reset_index(name="Frecuencia")
    return df.sort_values(["Frecuencia", "Descriptor"], ascending=[False, True], ignore_index=True)

def make_pareto(copilado_df: pd.DataFrame, cat_df: pd.DataFrame) -> pd.DataFrame:
    if copilado_df.empty:
        return pd.DataFrame(columns=["Categoría", "Descriptor", "Frecuencia", "Porcentaje", "% Acumulado", "Acumulado", "80/20"])
    # unir categoría
    df = copilado_df.merge(cat_df[["DESCRIPTOR", "CATEGORÍA"]], left_on="Descriptor", right_on="DESCRIPTOR", how="left")
    df["Categoría"] = df["CATEGORÍA"].fillna("")
    df = df.drop(columns=["DESCRIPTOR", "CATEGORÍA"])
    total = df["Frecuencia"].sum()
    df["Porcentaje"] = (df["Frecuencia"] / total) * 100.0
    df = df.sort_values(["Frecuencia", "Descriptor"], ascending=[False, True], ignore_index=True)
    df["% Acumulado"] = df["Porcentaje"].cumsum()
    df["Acumulado"] = df["Frecuencia"].cumsum()
    df["80/20"] = np.where(df["% Acumulado"] <= 80.0, "≤80%", ">80%")
    return df[["Categoría", "Descriptor", "Frecuencia", "Porcentaje", "% Acumulado", "Acumulado", "80/20"]]

def plot_pareto_plotly(df_pareto: pd.DataFrame):
    import plotly.graph_objects as go  # import local por seguridad
    x = df_pareto["Descriptor"].astype(str).tolist()
    y = df_pareto["Frecuencia"].tolist()
    cum = df_pareto["% Acumulado"].tolist()
    fig = go.Figure()
    fig.add_bar(x=x, y=y, name="Frecuencia", yaxis="y1")
    fig.add_trace(go.Scatter(x=x, y=cum, name="% Acumulado", yaxis="y2", mode="lines+markers"))
    fig.add_trace(go.Scatter(x=x, y=[80.0]*len(x), name="Corte 80%", yaxis="y2", mode="lines", line=dict(dash="dash")))
    fig.update_layout(
        title="Pareto Comunidad",
        xaxis_title="Descriptor",
        yaxis=dict(title="Frecuencia"),
        yaxis2=dict(title="% Acumulado", overlaying="y", side="right", range=[0, 100]),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=20, r=20, t=50, b=80),
        hovermode="x unified"
    )
    return fig

def plot_pareto_fallback(df_pareto: pd.DataFrame):
    # Fallback básico usando dos charts nativos de Streamlit
    st.write("⚠️ Plotly no disponible: mostrando barras y línea en dos gráficos simples.")
    st.bar_chart(df_pareto.set_index("Descriptor")["Frecuencia"])
    st.line_chart(df_pareto.set_index("Descriptor")["% Acumulado"])

def build_excel_bytes(copilado_df: pd.DataFrame, pareto_df: pd.DataFrame) -> bytes:
    from pandas import ExcelWriter
    import xlsxwriter  # noqa: F401
    output = BytesIO()
    with ExcelWriter(output, engine="xlsxwriter") as writer:
        copilado_df.to_excel(writer, index=False, sheet_name="Copilado Comunidad")
        pareto_df.to_excel(writer, index=False, sheet_name="Pareto Comunidad")
        # Gráfico embebido (si xlsxwriter disponible)
        wb = writer.book
        wsP = writer.sheets["Pareto Comunidad"]
        n = len(pareto_df)
        if n >= 1:
            chart = wb.add_chart({'type': 'column'})
            chart.add_series({
                'name': 'Frecuencia',
                'categories': ['Pareto Comunidad', 1, 1, n, 1],  # B
                'values':     ['Pareto Comunidad', 1, 2, n, 2],  # C
            })
            line = wb.add_chart({'type': 'line'})
            line.add_series({
                'name': '% Acumulado',
                'categories': ['Pareto Comunidad', 1, 1, n, 1],  # B
                'values':     ['Pareto Comunidad', 1, 4, n, 4],  # E
                'y2_axis': True,
            })
            chart.combine(line)
            chart.set_title({'name': 'Pareto Comunidad'})
            chart.set_x_axis({'name': 'Descriptor'})
            chart.set_y_axis({'name': 'Frecuencia'})
            chart.set_y2_axis({'name': '% Acumulado', 'min': 0, 'max': 100})
            wsP.insert_chart(1, 9, chart, {'x_scale': 1.2, 'y_scale': 1.2})
    return output.getvalue()

# ---------- UI ----------
st.title("Pareto Comunidad (MSP) – mínima y funcional")

c1, c2 = st.columns([2,1])
with c1:
    plantilla = st.file_uploader("📄 Subí la Plantilla (XLSX, hoja 'matriz')", type=["xlsx"])
with c2:
    demo = st.button("Probar DEMO (sin subir archivo)")

cat_df = build_catalogo_df()

st.divider()

try:
    if demo:
        # DEMO mínima para comprobar que la app abre y calcula
        st.info("Demo cargada: generando datos de ejemplo…")
        demo_df = pd.DataFrame({
            "hurto": [1,0,1,0,1,1,0],
            "robo": [0,1,0,1,0,0,1],
            "venta de drogas": [1,1,0,0,1,0,0],
            "observaciones": ["", "hurto", "asaltos y hurto", "", "búnker", "", ""],
        })
        by_desc, by_nc = build_keyword_maps(cat_df)
        hits = detect_by_headers(demo_df, by_desc, by_nc)
        copilado = make_copilado(hits)
        pareto = make_pareto(copilado, cat_df)
        st.subheader("Copilado (DEMO)")
        st.dataframe(copilado, use_container_width=True)
        st.subheader("Pareto (DEMO)")
        st.dataframe(pareto, use_container_width=True)
        st.subheader("Gráfico (DEMO)")
        plot_df = pareto.head(TOP_N_GRAFICO).copy()
        if HAS_PLOTLY:
            st.plotly_chart(plot_pareto_plotly(plot_df), use_container_width=True)
        else:
            plot_pareto_fallback(plot_df)
        st.download_button(
            "⬇️ Descargar Excel (DEMO)",
            data=build_excel_bytes(copilado, pareto),
            file_name="Pareto_Comunidad_DEMO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    if plantilla is not None:
        with st.spinner("Leyendo hoja 'matriz'…"):
            df_matriz = read_plantilla_matriz(plantilla)

        st.subheader("Previsualización (primeras filas)")
        st.dataframe(df_matriz.head(200), use_container_width=True)

        by_desc, by_nc = build_keyword_maps(cat_df)

        st.subheader("Copilado Comunidad")
        with st.spinner("Procesando por encabezados…"):
            hits = detect_by_headers(df_matriz, by_desc, by_nc)
            copilado = make_copilado(hits)

        if copilado.empty:
            st.warning("No se detectaron descriptores con el catálogo actual (por encabezados). Amplía el catálogo embebido si es necesario.")
        else:
            st.dataframe(copilado, use_container_width=True)

            st.subheader("Pareto Comunidad")
            pareto = make_pareto(copilado, cat_df)
            st.dataframe(pareto, use_container_width=True)

            st.subheader("Gráfico Pareto")
            plot_df = pareto.head(TOP_N_GRAFICO).copy()
            if HAS_PLOTLY:
                st.plotly_chart(plot_pareto_plotly(plot_df), use_container_width=True)
            else:
                plot_pareto_fallback(plot_df)

            st.subheader("Descarga")
            st.download_button(
                "⬇️ Descargar Excel (Copilado + Pareto + Gráfico)",
                data=build_excel_bytes(copilado, pareto),
                file_name="Pareto_Comunidad.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    if (plantilla is None) and (not demo):
        st.info("Subí la **Plantilla** o usa el botón **DEMO** para comprobar que la app funciona.")

except Exception as e:
    st.error(f"⚠️ Error de ejecución: {e}")
