# app.py — Pareto 80/20 + Portafolio + Unificado + Google Sheets + Informe PDF (verde/azul)
# -----------------------------------------------------------------------------------------
# Requisitos:
#   pip install -r requirements.txt
#   streamlit run app.py
# -----------------------------------------------------------------------------------------

import io
import os
from typing import List, Dict, Tuple

import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# ====== Google Sheets (DB) ======
import gspread
from google.oauth2.service_account import Credentials

# ====== PDF ======
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from PIL import Image  # noqa: F401 (usada por reportlab internamente)

# ----------------- CONFIG -----------------
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1cf-avzRjtBXcqr69WfrrsTAegm0PMAe8LgjeLpfcS5g/edit?usp=sharing"
WS_PARETOS = "paretos"  # hoja donde se guardan los paretos (nombre, descriptor, frecuencia)

st.set_page_config(page_title="Pareto de Descriptores", layout="wide")

# Paleta (verde/azul)
VERDE = "#1B9E77"
AZUL  = "#2C7FB8"
NEGRO = "#000000"

# Íconos/portadas esperados en el directorio de la app
IMG_ICONO_PEQUENO     = "Iconos pequeños.png"         # esquina sup-izq en páginas intermedias
IMG_PORTADA_GRANDE    = "Icono grande portada.png"    # portada
IMG_PORTADA_MEDIANO   = "Iconos medianos portada.png" # portada (fallback)

# ==========================================
# 1) CATÁLOGO EMBEBIDO
# ==========================================
CATALOGO: List[Dict[str, str]] = [
    {"categoria": "Delito", "descriptor": "Abandono de personas (menor de edad, adulto mayor o con capacidades diferentes)"},
    {"categoria": "Delito", "descriptor": "Abigeato (robo y destace de ganado)"},
    {"categoria": "Delito", "descriptor": "Aborto"},
    {"categoria": "Delito", "descriptor": "Abuso de autoridad"},
    {"categoria": "Riesgo social", "descriptor": "Accidentes de tránsito"},
    {"categoria": "Delito", "descriptor": "Accionamiento de arma de fuego (balaceras)"},
    {"categoria": "Riesgo social", "descriptor": "Acoso escolar (bullying)"},
    {"categoria": "Riesgo social", "descriptor": "Acoso laboral (mobbing)"},
    {"categoria": "Riesgo social", "descriptor": "Acoso sexual callejero"},
    {"categoria": "Riesgo social", "descriptor": "Actos obscenos en vía pública"},
    {"categoria": "Delito", "descriptor": "Administración fraudulenta, apropiaciones indebidas o enriquecimiento ilícito"},
    {"categoria": "Delito", "descriptor": "Agresión con armas"},
    {"categoria": "Riesgo social", "descriptor": "Agrupaciones delincuenciales no organizadas"},
    {"categoria": "Delito", "descriptor": "Alteración de datos y sabotaje informático"},
    {"categoria": "Otros factores", "descriptor": "Ambiente laboral inadecuado"},
    {"categoria": "Delito", "descriptor": "Amenazas"},
    {"categoria": "Riesgo social", "descriptor": "Analfabetismo"},
    {"categoria": "Riesgo social", "descriptor": "Bajos salarios"},
    {"categoria": "Riesgo social", "descriptor": "Barras de fútbol"},
    {"categoria": "Riesgo social", "descriptor": "Búnker (eje de expendio de drogas)"},
    {"categoria": "Delito", "descriptor": "Calumnia"},
    {"categoria": "Delito", "descriptor": "Caza ilegal"},
    {"categoria": "Delito", "descriptor": "Conducción temeraria"},
    {"categoria": "Riesgo social", "descriptor": "Consumo de alcohol en vía pública"},
    {"categoria": "Riesgo social", "descriptor": "Consumo de drogas"},
    {"categoria": "Riesgo social", "descriptor": "Contaminación sónica"},
    {"categoria": "Delito", "descriptor": "Contrabando"},
    {"categoria": "Delito", "descriptor": "Corrupción"},
    {"categoria": "Delito", "descriptor": "Corrupción policial"},
    {"categoria": "Delito", "descriptor": "Cultivo de droga (marihuana)"},
    {"categoria": "Delito", "descriptor": "Daño ambiental"},
    {"categoria": "Delito", "descriptor": "Daños/vandalismo"},
    {"categoria": "Riesgo social", "descriptor": "Deficiencia en la infraestructura vial"},
    {"categoria": "Otros factores", "descriptor": "Deficiencia en la línea 9-1-1"},
    {"categoria": "Riesgo social", "descriptor": "Deficiencias en el alumbrado público"},
    {"categoria": "Delito", "descriptor": "Delincuencia organizada"},
    {"categoria": "Delito", "descriptor": "Delitos contra el ámbito de intimidad (violación de secretos, correspondencia y comunicaciones electrónicas)"},
    {"categoria": "Delito", "descriptor": "Delitos sexuales"},
    {"categoria": "Riesgo social", "descriptor": "Desaparición de personas"},
    {"categoria": "Riesgo social", "descriptor": "Desarticulación interinstitucional"},
    {"categoria": "Riesgo social", "descriptor": "Desempleo"},
    {"categoria": "Riesgo social", "descriptor": "Desvinculación estudiantil"},
    {"categoria": "Delito", "descriptor": "Desobediencia"},
    {"categoria": "Delito", "descriptor": "Desórdenes en vía pública"},
    {"categoria": "Delito", "descriptor": "Disturbios (riñas)"},
    {"categoria": "Riesgo social", "descriptor": "Enfrentamientos estudiantiles"},
    {"categoria": "Delito", "descriptor": "Estafa o defraudación"},
    {"categoria": "Delito", "descriptor": "Estupro (delitos sexuales contra menor de edad)"},
    {"categoria": "Delito", "descriptor": "Evasión y quebrantamiento de pena"},
    {"categoria": "Delito", "descriptor": "Explosivos"},
    {"categoria": "Delito", "descriptor": "Extorsión"},
    {"categoria": "Delito", "descriptor": "Fabricación, producción o reproducción de pornografía"},
    {"categoria": "Riesgo social", "descriptor": "Facilismo económico"},
    {"categoria": "Delito", "descriptor": "Falsificación de moneda y otros valores"},
    {"categoria": "Riesgo social", "descriptor": "Falta de cámaras de seguridad"},
    {"categoria": "Otros factores", "descriptor": "Falta de capacitación policial"},
    {"categoria": "Riesgo social", "descriptor": "Falta de control a patentes"},
    {"categoria": "Riesgo social", "descriptor": "Falta de control fronterizo"},
    {"categoria": "Riesgo social", "descriptor": "Falta de corresponsabilidad en seguridad"},
    {"categoria": "Riesgo social", "descriptor": "Falta de cultura vial"},
    {"categoria": "Riesgo social", "descriptor": "Falta de cultura y compromiso ciudadano"},
    {"categoria": "Riesgo social", "descriptor": "Falta de educación familiar"},
    {"categoria": "Otros factores", "descriptor": "Falta de incentivos"},
    {"categoria": "Riesgo social", "descriptor": "Falta de inversión social"},
    {"categoria": "Riesgo social", "descriptor": "Falta de legislación de extinción de dominio"},
    {"categoria": "Otros factores", "descriptor": "Falta de personal administrativo"},
    {"categoria": "Otros factores", "descriptor": "Falta de personal policial"},
    {"categoria": "Otros factores", "descriptor": "Falta de policías de tránsito"},
    {"categoria": "Riesgo social", "descriptor": "Falta de políticas públicas en seguridad"},
    {"categoria": "Riesgo social", "descriptor": "Falta de presencia policial"},
    {"categoria": "Riesgo social", "descriptor": "Falta de salubridad pública"},
    {"categoria": "Riesgo social", "descriptor": "Familias disfuncionales"},
    {"categoria": "Delito", "descriptor": "Fraude informático"},
    {"categoria": "Delito", "descriptor": "Grooming"},
    {"categoria": "Riesgo social", "descriptor": "Hacinamiento carcelario"},
    {"categoria": "Riesgo social", "descriptor": "Hacinamiento policial"},
    {"categoria": "Delito", "descriptor": "Homicidio"},
    {"categoria": "Riesgo social", "descriptor": "Hospedajes ilegales (cuarterías)"},
    {"categoria": "Delito", "descriptor": "Hurto"},
    {"categoria": "Otros factores", "descriptor": "Inadecuado uso del recurso policial"},
    {"categoria": "Riesgo social", "descriptor": "Incumplimiento al plan regulador de la municipalidad"},
    {"categoria": "Delito", "descriptor": "Incumplimiento del deber alimentario"},
    {"categoria": "Riesgo social", "descriptor": "Indiferencia social"},
    {"categoria": "Otros factores", "descriptor": "Inefectividad en el servicio de policía"},
    {"categoria": "Riesgo social", "descriptor": "Ineficiencia en la administración de justicia"},
    {"categoria": "Otros factores", "descriptor": "Infraestructura inadecuada"},
    {"categoria": "Riesgo social", "descriptor": "Intolerancia social"},
    {"categoria": "Otros factores", "descriptor": "Irrespeto a la jefatura"},
    {"categoria": "Otros factores", "descriptor": "Irrespeto al subalterno"},
    {"categoria": "Otros factores", "descriptor": "Jornadas laborales extensas"},
    {"categoria": "Delito", "descriptor": "Lavado de activos"},
    {"categoria": "Delito", "descriptor": "Lesiones"},
    {"categoria": "Delito", "descriptor": "Ley de armas y explosivos N° 7530"},
    {"categoria": "Riesgo social", "descriptor": "Ley de control de tabaco (Ley 9028)"},
    {"categoria": "Riesgo social", "descriptor": "Lotes baldíos"},
    {"categoria": "Delito", "descriptor": "Maltrato animal"},
    {"categoria": "Delito", "descriptor": "Narcotráfico"},
    {"categoria": "Riesgo social", "descriptor": "Necesidades básicas insatisfechas"},
    {"categoria": "Riesgo social", "descriptor": "Percepción de inseguridad"},
    {"categoria": "Riesgo social", "descriptor": "Pérdida de espacios públicos"},
    {"categoria": "Riesgo social", "descriptor": "Personas con exceso de tiempo de ocio"},
    {"categoria": "Riesgo social", "descriptor": "Personas en estado migratorio irregular"},
    {"categoria": "Riesgo social", "descriptor": "Personas en situación de calle"},
    {"categoria": "Delito", "descriptor": "Menores en vulnerabilidad"},
    {"categoria": "Delito", "descriptor": "Pesca ilegal"},
    {"categoria": "Delito", "descriptor": "Portación ilegal de armas"},
    {"categoria": "Riesgo social", "descriptor": "Presencia multicultural"},
    {"categoria": "Otros factores", "descriptor": "Presión por resultados operativos"},
    {"categoria": "Delito", "descriptor": "Privación de libertad sin ánimo de lucro"},
    {"categoria": "Riesgo social", "descriptor": "Problemas vecinales"},
    {"categoria": "Delito", "descriptor": "Receptación"},
    {"categoria": "Delito", "descriptor": "Relaciones impropias"},
    {"categoria": "Delito", "descriptor": "Resistencia (irrespeto a la autoridad)"},
    {"categoria": "Delito", "descriptor": "Robo a comercio (intimidación)"},
    {"categoria": "Delito", "descriptor": "Robo a comercio (tacha)"},
    {"categoria": "Delito", "descriptor": "Robo a edificación (tacha)"},
    {"categoria": "Delito", "descriptor": "Robo a personas"},
    {"categoria": "Delito", "descriptor": "Robo a transporte comercial"},
    {"categoria": "Delito", "descriptor": "Robo a vehículos (tacha)"},
    {"categoria": "Delito", "descriptor": "Robo a vivienda (intimidación)"},
    {"categoria": "Delito", "descriptor": "Robo a vivienda (tacha)"},
    {"categoria": "Delito", "descriptor": "Robo de bicicleta"},
    {"categoria": "Delito", "descriptor": "Robo de cultivos"},
    {"categoria": "Delito", "descriptor": "Robo de motocicletas/vehículos (bajonazo)"},
    {"categoria": "Delito", "descriptor": "Robo de vehículos"},
    {"categoria": "Delito", "descriptor": "Secuestro"},
    {"categoria": "Delito", "descriptor": "Simulación de delito"},
    {"categoria": "Riesgo social", "descriptor": "Sistema jurídico desactualizado"},
    {"categoria": "Riesgo social", "descriptor": "Suicidio"},
    {"categoria": "Delito", "descriptor": "Sustracción de una persona menor de edad o incapaz"},
    {"categoria": "Delito", "descriptor": "Tala ilegal"},
    {"categoria": "Riesgo social", "descriptor": "Tendencia social hacia el delito (pautas de crianza violenta)"},
    {"categoria": "Riesgo social", "descriptor": "Tenencia de droga"},
    {"categoria": "Delito", "descriptor": "Tentativa de homicidio"},
    {"categoria": "Delito", "descriptor": "Terrorismo"},
    {"categoria": "Riesgo social", "descriptor": "Trabajo informal"},
    {"categoria": "Delito", "descriptor": "Tráfico de armas"},
    {"categoria": "Delito", "descriptor": "Tráfico de influencias"},
    {"categoria": "Riesgo social", "descriptor": "Transporte informal (Uber, porteadores, piratas)"},
    {"categoria": "Delito", "descriptor": "Trata de personas"},
    {"categoria": "Delito", "descriptor": "Turbación de actos religiosos y profanaciones"},
    {"categoria": "Delito", "descriptor": "Uso ilegal de uniformes, insignias o dispositivos policiales"},
    {"categoria": "Delito", "descriptor": "Usurpación de terrenos (precarios)"},
    {"categoria": "Delito", "descriptor": "Venta de drogas"},
    {"categoria": "Riesgo social", "descriptor": "Ventas informales (ambulantes)"},
    {"categoria": "Riesgo social", "descriptor": "Vigilancia informal"},
    {"categoria": "Delito", "descriptor": "Violación de domicilio"},
    {"categoria": "Delito", "descriptor": "Violación de la custodia de las cosas"},
    {"categoria": "Delito", "descriptor": "Violación de sellos"},
    {"categoria": "Delito", "descriptor": "Violencia de género"},
    {"categoria": "Delito", "descriptor": "Violencia intrafamiliar"},
    {"categoria": "Riesgo social", "descriptor": "Xenofobia"},
    {"categoria": "Riesgo social", "descriptor": "Zonas de prostitución"},
    {"categoria": "Riesgo social", "descriptor": "Zonas vulnerables"},
    {"categoria": "Delito", "descriptor": "Robo a transporte público con intimidación"},
    {"categoria": "Delito", "descriptor": "Robo de cable"},
    {"categoria": "Delito", "descriptor": "Explotación sexual infantil"},
    {"categoria": "Delito", "descriptor": "Explotación laboral infantil"},
    {"categoria": "Delito", "descriptor": "Tráfico ilegal de personas"},
    {"categoria": "Riesgo social", "descriptor": "Bares clandestinos"},
    {"categoria": "Delito", "descriptor": "Robo de combustible"},
    {"categoria": "Delito", "descriptor": "Femicidio"},
    {"categoria": "Delito", "descriptor": "Delitos contra la vida (homicidios, heridos)"},
    {"categoria": "Delito", "descriptor": "Venta y consumo de drogas en vía pública"},
    {"categoria": "Delito", "descriptor": "Asalto (a personas, comercio, vivienda, transporte público)"},
    {"categoria": "Delito", "descriptor": "Robo de ganado y agrícola"},
    {"categoria": "Delito", "descriptor": "Robo de equipo agrícola"},
]

# ==========================================
# 2) UTILIDADES BASE
# ==========================================
def _map_descriptor_a_categoria() -> Dict[str, str]:
    df = pd.DataFrame(CATALOGO); return dict(zip(df["descriptor"], df["categoria"]))
DESC2CAT = _map_descriptor_a_categoria()

def normalizar_freq_map(freq_map: Dict[str, int]) -> Dict[str, int]:
    out = {}
    for d, v in (freq_map or {}).items():
        try:
            vv = int(pd.to_numeric(v, errors="coerce"))
            if vv > 0: out[d] = vv
        except Exception:
            continue
    return out

def df_desde_freq_map(freq_map: Dict[str, int]) -> pd.DataFrame:
    items = []
    for d, f in normalizar_freq_map(freq_map).items():
        items.append({"descriptor": d, "categoria": DESC2CAT.get(d, "—"), "frecuencia": int(f)})
    df = pd.DataFrame(items)
    if df.empty: return pd.DataFrame(columns=["descriptor", "categoria", "frecuencia"])
    return df

def combinar_maps(maps: List[Dict[str, int]]) -> Dict[str, int]:
    total = {}
    for m in maps:
        for d, f in normalizar_freq_map(m).items():
            total[d] = total.get(d, 0) + int(f)
    return total

def info_pareto(freq_map: Dict[str, int]) -> Dict[str, int]:
    d = normalizar_freq_map(freq_map); return {"descriptores": len(d), "total": int(sum(d.values()))}

# --- Cálculo Pareto ---
def calcular_pareto(df_in: pd.DataFrame) -> pd.DataFrame:
    df = df_in.copy()
    df["frecuencia"] = pd.to_numeric(df["frecuencia"], errors="coerce").fillna(0).astype(int)
    df = df[df["frecuencia"] > 0]
    if df.empty:
        return df.assign(porcentaje=0.0, acumulado=0, pct_acum=0.0,
                         segmento_real="20%", segmento="80%")
    df = df.sort_values("frecuencia", ascending=False)
    total = int(df["frecuencia"].sum())
    df["porcentaje"] = (df["frecuencia"] / total * 100).round(2)
    df["acumulado"]  = df["frecuencia"].cumsum()
    df["pct_acum"]   = (df["acumulado"] / total * 100).round(2)
    df["segmento_real"] = np.where(df["pct_acum"] <= 80.00, "80%", "20%")
    df["segmento"] = "80%"
    return df.reset_index(drop=True)

def _colors_for_segments(segments: List[str]) -> List[str]:
    return [VERDE if s == "80%" else AZUL for s in segments]

def dibujar_pareto(df_par: pd.DataFrame, titulo: str):
    if df_par.empty:
        st.info("Ingresa frecuencias (>0) para ver el gráfico.")
        return
    x        = np.arange(len(df_par))
    freqs    = df_par["frecuencia"].to_numpy()
    pct_acum = df_par["pct_acum"].to_numpy()
    colors_b = _colors_for_segments(df_par["segmento_real"].tolist())

    fig, ax1 = plt.subplots(figsize=(14, 5))
    ax1.bar(x, freqs, color=colors_b)
    ax1.set_ylabel("Frecuencia")
    ax1.set_xticks(x)
    ax1.set_xticklabels(df_par["descriptor"].tolist(), rotation=75, ha="right")
    ax1.set_title(titulo if titulo.strip() else "Pareto — Frecuencia y % acumulado", color="#124559")
    ax2 = ax1.twinx()
    ax2.plot(x, pct_acum, marker="o", linewidth=2)
    ax2.set_ylabel("% acumulado")
    ax2.set_ylim(0, 110)
    if (df_par["segmento_real"] == "80%").any():
        cut_idx = np.where(df_par["segmento_real"].to_numpy() == "80%")[0].max()
        ax1.axvline(cut_idx, linestyle=":", color="k")
    ax2.axhline(80, linestyle="--", linewidth=1)
    st.pyplot(fig)

# --- Excel export ---
def exportar_excel_con_grafico(df_par: pd.DataFrame, titulo: str) -> bytes:
    import xlsxwriter
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        hoja = "Pareto"
        df_x = df_par.copy()
        df_x["porcentaje"] = (df_x["porcentaje"] / 100.0).round(4)
        df_x["pct_acum"]   = (df_x["pct_acum"] / 100.0).round(4)
        df_x = df_x[["categoria", "descriptor", "frecuencia",
                     "porcentaje", "pct_acum", "acumulado", "segmento"]]
        df_x.to_excel(writer, sheet_name=hoja, index=False, startrow=0, startcol=0)
        wb = writer.book; ws = writer.sheets[hoja]
        pct_fmt = wb.add_format({"num_format": "0.00%"})
        total_fmt = wb.add_format({"bold": True})
        ws.set_column("A:A", 18); ws.set_column("B:B", 55); ws.set_column("C:C", 12)
        ws.set_column("D:D", 12, pct_fmt); ws.set_column("E:E", 18, pct_fmt)
        ws.set_column("F:F", 12); ws.set_column("G:G", 10)
        n = len(df_x)
        cats = f"=Pareto!$B$2:$B${n+1}"; vals = f"=Pareto!$C$2:$C${n+1}"; pcts = f"=Pareto!$E$2:$E${n+1}"
        total = int(df_par["frecuencia"].sum())
        ws.write(n + 2, 1, "TOTAL:", total_fmt); ws.write(n + 2, 2, total, total_fmt)
        try:
            idxs = np.where(df_par["segmento_real"].to_numpy() == "80%")[0]
            if len(idxs) > 0:
                last = int(idxs.max())
                green_bg = wb.add_format({"bg_color": VERDE, "font_color": "#FFFFFF"})
                ws.conditional_format(1, 0, 1 + last, 6, {"type": "no_blanks", "format": green_bg})
        except Exception:
            pass
        chart = wb.add_chart({"type": "column"})
        points = [{"fill": {"color": (VERDE if s == "80%" else AZUL)}} for s in df_par["segmento_real"]]
        chart.add_series({"name": "Frecuencia", "categories": cats, "values": vals, "points": points})
        line = wb.add_chart({"type": "line"})
        line.add_series({"name": "% acumulado", "categories": cats, "values": pcts,
                         "y2_axis": True, "marker": {"type": "circle"}})
        chart.combine(line)
        chart.set_y_axis({"name": "Frecuencia"})
        chart.set_y2_axis({"name": "Porcentaje acumulado",
                           "min": 0, "max": 1.10, "major_unit": 0.10, "num_format": "0%"})
        chart.set_title({"name": titulo if titulo.strip() else "PARETO – Frecuencia y % acumulado"})
        chart.set_legend({"position": "bottom"}); chart.set_size({"width": 1180, "height": 420})
        ws.insert_chart("I2", chart)
    return output.getvalue()

# ==========================================
# 3) GOOGLE SHEETS HELPERS
# ==========================================
SCOPES = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]

def _gc():
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=SCOPES)
    return gspread.authorize(creds)

def _open_sheet():
    gc = _gc(); return gc.open_by_url(SPREADSHEET_URL)

def _ensure_ws(sh, title: str, header: List[str]):
    try:
        ws = sh.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=1000, cols=10)
        ws.append_row(header); return ws
    values = ws.get_all_values()
    if not values:
        ws.append_row(header)
    else:
        first = values[0]
        if [c.strip().lower() for c in first] != [c.strip().lower() for c in header]:
            ws.clear(); ws.append_row(header)
    return ws

def sheets_cargar_portafolio() -> Dict[str, Dict[str, int]]:
    try:
        sh = _open_sheet(); ws = _ensure_ws(sh, WS_PARETOS, ["nombre","descriptor","frecuencia"])
        rows = ws.get_all_records()
        port: Dict[str, Dict[str, int]] = {}
        for r in rows:
            nom = str(r.get("nombre","")).strip()
            desc = str(r.get("descriptor","")).strip()
            freq = int(pd.to_numeric(r.get("frecuencia",0), errors="coerce") or 0)
            if not nom or not desc or freq <= 0: continue
            bucket = port.setdefault(nom, {}); bucket[desc] = bucket.get(desc, 0) + freq
        return port
    except Exception:
        return {}

def sheets_guardar_pareto(nombre: str, freq_map: Dict[str, int], sobrescribir: bool = True):
    sh = _open_sheet()
    ws = _ensure_ws(sh, WS_PARETOS, ["nombre","descriptor","frecuencia"])
    if sobrescribir:
        vals = ws.get_all_values()
        header = vals[0] if vals else ["nombre","descriptor","frecuencia"]
        others = [r for r in vals[1:] if (len(r)>0 and r[0].strip().lower()!=nombre.strip().lower())]
        ws.clear(); ws.update("A1", [header])
        if others: ws.append_rows(others, value_input_option="RAW")
    rows_new = [[nombre, d, int(f)] for d, f in normalizar_freq_map(freq_map).items()]
    if rows_new: ws.append_rows(rows_new, value_input_option="RAW")

# ==========================================
# 4) ESTADO DE SESIÓN
# ==========================================
st.session_state.setdefault("freq_map", {})
st.session_state.setdefault("portafolio", {})
st.session_state.setdefault("msel", [])
st.session_state.setdefault("reset_after_save", False)

if not st.session_state["portafolio"]:
    loaded = sheets_cargar_portafolio()
    if loaded: st.session_state["portafolio"].update(loaded)

if st.session_state.get("reset_after_save", False):
    st.session_state["freq_map"] = {}
    st.session_state["msel"] = []
    st.session_state.pop("editor_freq", None)
    st.session_state["reset_after_save"] = False

# ==========================================
# 5) MAPA DE IMÁGENES POR DESCRIPTOR
# ==========================================
IMG_BY_KEYWORD = {
    "búnker": "Bunker.png",
    "bunker": "Bunker.png",
    "consumo de drogas": "Consumo de drogas.png",
    "deficiencia en la infraestructura vial": "Deficiencia en la infraestructura Vial.png",
    "estafa": "Estafa o defraudacion.png",
    "falta de inversión social": "Falta de Inversion social.png",
    "venta de drogas": "Venta de drogas.png",
    "violencia intrafamiliar": "Violencia intrafamiliar.png",
}
def _find_image_for_descriptor(descriptor: str) -> str | None:
    dnorm = descriptor.lower()
    for k, fname in IMG_BY_KEYWORD.items():
        if k in dnorm and os.path.exists(fname):
            return fname
    return None

def _safe_image_reader(path: str | None) -> ImageReader | None:
    try:
        if path and os.path.exists(path):
            return ImageReader(path)
    except Exception:
        return None
    return None

# ==========================================
# 6) PDF BUILDER
# ==========================================
PAGE_W, PAGE_H = A4

def _draw_header_icon_if_needed(c: canvas.Canvas, is_first: bool, is_last: bool):
    if is_first or is_last:
        return
    ir = _safe_image_reader(IMG_ICONO_PEQUENO)
    if ir:
        c.drawImage(ir, 1.2*cm, PAGE_H - 2.4*cm, width=1.1*cm, height=1.1*cm, mask='auto')

def _title_portada(c: canvas.Canvas, titulo: str):
    c.setFillColor(colors.HexColor("#124559"))
    ir = _safe_image_reader(IMG_PORTADA_GRANDE) or _safe_image_reader(IMG_PORTADA_MEDIANO)
    if ir:
        iw, ih = ir.getSize()
        ratio = (8*cm) / max(iw, ih)
        w = iw*ratio; h = ih*ratio
        c.drawImage(ir, (PAGE_W - w)/2, PAGE_H - h - 4*cm, width=w, height=h, mask='auto')
    c.setFont("Helvetica-Bold", 24)
    c.drawCentredString(PAGE_W/2, PAGE_H - 12*cm, titulo)
    c.setFont("Helvetica", 12)
    c.setFillColor(colors.black)
    c.drawCentredString(PAGE_W/2, PAGE_H - 13.3*cm, "Estrategia Integral de Prevención — Sembremos Seguridad")
    c.showPage()

def _table_results_image(df: pd.DataFrame) -> bytes:
    fig, ax = plt.subplots(figsize=(8.4, 6.2))
    ax.axis('off')
    show = df[["categoria","descriptor","frecuencia"]].rename(columns={
        "categoria":"Categoría","descriptor":"Descriptor","frecuencia":"Frecuencia"
    })
    the_table = ax.table(cellText=show.values, colLabels=show.columns, loc='center')
    the_table.auto_set_font_size(False); the_table.set_fontsize(9)
    the_table.scale(1, 1.2)
    ax.set_title("Resultados", pad=18, color="#124559")
    buf = io.BytesIO(); fig.tight_layout(); fig.savefig(buf, format="PNG", dpi=200); plt.close(fig)
    return buf.getvalue()

def _pareto_fig_image(df_par: pd.DataFrame, titulo: str) -> bytes:
    x        = np.arange(len(df_par))
    freqs    = df_par["frecuencia"].to_numpy()
    pct_acum = df_par["pct_acum"].to_numpy()
    colors_b = _colors_for_segments(df_par["segmento_real"].tolist())
    fig, ax1 = plt.subplots(figsize=(10, 4))
    ax1.bar(x, freqs, color=colors_b)
    ax1.set_ylabel("Frecuencia")
    ax1.set_xticks(x); ax1.set_xticklabels(df_par["descriptor"].tolist(), rotation=75, ha="right", fontsize=8)
    ax1.set_title(titulo if titulo.strip() else "Pareto — Frecuencia y % acumulado", color="#124559")
    ax2 = ax1.twinx()
    ax2.plot(x, pct_acum, marker="o", linewidth=2)
    ax2.set_ylabel("% acumulado"); ax2.set_ylim(0, 110)
    if (df_par["segmento_real"] == "80%").any():
        cut_idx = np.where(df_par["segmento_real"].to_numpy() == "80%")[0].max()
        ax1.axvline(cut_idx, linestyle=":", color="k")
    ax2.axhline(80, linestyle="--", linewidth=1)
    buf = io.BytesIO(); fig.tight_layout(); fig.savefig(buf, format="PNG", dpi=200); plt.close(fig)
    return buf.getvalue()

def _pie_image_from_breakdown(title: str, data_pairs: List[Tuple[str, float]]) -> bytes:
    labels = [l for l, _ in data_pairs if l]
    sizes  = [float(p or 0) for _, p in data_pairs if _]
    if not labels or sum(sizes) <= 0:
        labels, sizes = ["Sin datos"], [100]
    fig, ax = plt.subplots(figsize=(7.5, 5))
    ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=140)
    ax.axis('equal'); ax.set_title(title, color="#124559")
    buf = io.BytesIO(); fig.tight_layout(); fig.savefig(buf, format="PNG", dpi=200); plt.close(fig)
    return buf.getvalue()

def generar_pdf_informe(nombre_informe: str,
                        df_par: pd.DataFrame,
                        desgloses: List[Dict]) -> bytes:
    """
    desgloses: lista de secciones. Cada dict:
        {"descriptor": str, "rows": [{"Etiqueta": str, "%": float}, ...]}
    """
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    # 1) Portada
    _title_portada(c, nombre_informe)

    # 2) Resultados + imagen asociada al TOP 1 (si existe)
    _draw_header_icon_if_needed(c, is_first=False, is_last=False)
    top_img = _find_image_for_descriptor(df_par.iloc[0]["descriptor"]) if not df_par.empty else None
    ir = _safe_image_reader(top_img)
    tbl_png = _table_results_image(df_par)
    c.drawImage(ImageReader(io.BytesIO(tbl_png)), 1.5*cm, 8.0*cm, width=18*cm-3*cm, height=11*cm, mask='auto')
    if ir:
        c.drawImage(ir, PAGE_W-6.5*cm, PAGE_H-7.5*cm, width=5.5*cm, height=5.5*cm, mask='auto')
    c.showPage()

    # 3) Diagrama de Pareto
    _draw_header_icon_if_needed(c, is_first=False, is_last=False)
    pareto_png = _pareto_fig_image(df_par, "Diagrama de Pareto")
    c.drawImage(ImageReader(io.BytesIO(pareto_png)), 1.5*cm, 5.5*cm, width=18*cm-3*cm, height=11*cm, mask='auto')
    c.showPage()

    # 4) Modalidades
    for sec in desgloses:
        _draw_header_icon_if_needed(c, is_first=False, is_last=False)
        descriptor = sec.get("descriptor","").strip()
        pares = [(p.get("Etiqueta",""), float(p.get("%", 0) or 0)) for p in sec.get("rows", [])]
        c.setFont("Helvetica-Bold", 16); c.setFillColor(colors.HexColor("#124559"))
        c.drawString(2*cm, PAGE_H - 3*cm, f"Modalidades de la problemática — {descriptor}")
        dimg = _safe_image_reader(_find_image_for_descriptor(descriptor))
        if dimg:
            c.drawImage(dimg, PAGE_W-6.0*cm, PAGE_H-6.5*cm, width=4.8*cm, height=4.8*cm, mask='auto')
        pie_png = _pie_image_from_breakdown(descriptor, pares)
        c.drawImage(ImageReader(io.BytesIO(pie_png)), 2*cm, 4.0*cm, width=16*cm, height=11*cm, mask='auto')
        c.showPage()

    # 5) Cierre (última)
    c.setFont("Helvetica-Bold", 18); c.setFillColor(colors.HexColor("#124559"))
    c.drawCentredString(PAGE_W/2, PAGE_H - 6*cm, "Sembremos Seguridad – Resultados")
    c.setFont("Helvetica", 11); c.setFillColor(colors.black)
    c.drawCentredString(PAGE_W/2, PAGE_H - 7.2*cm, "Dirección de Programas Policiales Preventivos – MSP")
    c.save()
    return buf.getvalue()

# === Helpers UI para formulario de desgloses (reutilizable) ===
def ui_desgloses(descriptor_list: List[str], key_prefix: str) -> List[Dict]:
    st.caption("Opcional: agrega secciones de ‘Modalidades’ (hasta 3). Cada sección admite hasta 10 filas (Etiqueta + %).")
    n_secs = st.number_input("Cantidad de secciones de Modalidades",
                             min_value=0, max_value=3, value=1, step=1, key=f"{key_prefix}_nsecs")
    desgloses: List[Dict] = []
    for i in range(n_secs):
        with st.expander(f"Sección Modalidades #{i+1}", expanded=(i == 0)):
            dsel = st.selectbox(f"Descriptor para la sección #{i+1}",
                                options=["(elegir)"] + descriptor_list, index=0, key=f"{key_prefix}_desc_{i}")
            rows = [{"Etiqueta":"", "%":0.0} for _ in range(10)]
            df_rows = pd.DataFrame(rows)
            de = st.data_editor(
                df_rows, key=f"{key_prefix}_rows_{i}", use_container_width=True,
                column_config={
                    "Etiqueta": st.column_config.TextColumn("Etiqueta / Modalidad", width="large"),
                    "%": st.column_config.NumberColumn("Porcentaje", min_value=0.0, max_value=100.0, step=0.1)
                },
                num_rows="fixed"
            )
            total_pct = float(pd.to_numeric(de["%"], errors="coerce").fillna(0).sum())
            st.caption(f"Suma actual: **{total_pct:.1f}%** (se recomienda ~100%)")
            if dsel != "(elegir)":
                desgloses.append({"descriptor": dsel, "rows": de.to_dict(orient="records")})
    return desgloses

# ==========================================
# 7) UI PRINCIPAL (Editor Pareto)
# ==========================================
st.title("Pareto de Descriptores")

c_t1, c_t2, c_t3 = st.columns([2,1,1])
with c_t1:
    titulo = st.text_input("Título del Pareto (opcional)", value="Pareto Comunidad")
with c_t2:
    nombre_para_guardar = st.text_input("Nombre para guardar este Pareto", value="Comunidad")
with c_t3:
    if st.button("🔄 Recargar portafolio desde Sheets"):
        st.session_state["portafolio"] = sheets_cargar_portafolio()
        st.success("Portafolio recargado desde Google Sheets.")
        st.rerun()

cat_df = pd.DataFrame(CATALOGO).sort_values(["categoria","descriptor"]).reset_index(drop=True)
opciones = cat_df["descriptor"].tolist()
seleccion = st.multiselect("1) Escoge uno o varios descriptores", options=opciones,
                           default=st.session_state["msel"], key="msel")

st.subheader("2) Asigna la frecuencia")
if seleccion:
    base = cat_df[cat_df["descriptor"].isin(seleccion)].copy()
    base["frecuencia"] = [st.session_state["freq_map"].get(d, 0) for d in base["descriptor"]]

    edit = st.data_editor(
        base, key="editor_freq", num_rows="fixed", use_container_width=True,
        column_config={
            "descriptor": st.column_config.TextColumn("DESCRIPTOR", width="large"),
            "categoria": st.column_config.TextColumn("CATEGORÍA", width="small"),
            "frecuencia": st.column_config.NumberColumn("Frecuencia", min_value=0, step=1),
        },
    )
    for _, row in edit.iterrows():
        st.session_state["freq_map"][row["descriptor"]] = int(row["frecuencia"])

    df_in = edit[["descriptor","categoria"]].copy()
    df_in["frecuencia"] = df_in["descriptor"].map(st.session_state["freq_map"]).fillna(0).astype(int)

    st.subheader("3) Pareto (en edición)")
    tabla = calcular_pareto(df_in)

    mostrar = tabla.copy()[["categoria","descriptor","frecuencia","porcentaje","pct_acum","acumulado","segmento"]]
    mostrar = mostrar.rename(columns={"pct_acum": "porcentaje acumulado"})
    mostrar["porcentaje"] = mostrar["porcentaje"].map(lambda x: f"{x:.2f}%")
    mostrar["porcentaje acumulado"] = mostrar["porcentaje acumulado"].map(lambda x: f"{x:.2f}%")

    c1, c2 = st.columns([1,1], gap="large")
    with c1:
        st.markdown("**Tabla de Pareto**")
        if not tabla.empty:
            st.dataframe(mostrar, use_container_width=True, hide_index=True)
        else:
            st.info("Ingresa frecuencias (>0) para ver la tabla.")
    with c2:
        st.markdown("**Gráfico de Pareto**"); dibujar_pareto(tabla, titulo)

    st.subheader("4) Guardar / Descargar")
    col_g1, col_g2, _ = st.columns([1,1,2])
    with col_g1:
        sobrescribir = st.checkbox("Sobrescribir si existe", value=True)
        if st.button("💾 Guardar este Pareto"):
            nombre = nombre_para_guardar.strip()
            if not nombre:
                st.warning("Indica un nombre para guardar el Pareto.")
            else:
                st.session_state["portafolio"][nombre] = normalizar_freq_map(st.session_state["freq_map"])
                try:
                    sheets_guardar_pareto(nombre, st.session_state["freq_map"], sobrescribir=sobrescribir)
                    st.success(f"Pareto '{nombre}' guardado en Google Sheets y en la sesión.")
                except Exception as e:
                    st.warning(f"Se guardó en la sesión, pero hubo un problema con Sheets: {e}")
                st.session_state["reset_after_save"] = True
                st.rerun()
    with col_g2:
        if not tabla.empty:
            st.download_button(
                "⬇️ Excel del Pareto (edición)",
                data=exportar_excel_con_grafico(tabla, titulo),
                file_name=f"pareto_{(nombre_para_guardar or 'edicion').lower().replace(' ','_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # ===== Elaborador de INFORME PDF (desde el editor) =====
    st.markdown("---")
    st.header("🧾 Elaborar Informe PDF (desde el editor)")
    nombre_informe = st.text_input(
        "Nombre del informe (portada)", 
        value=(nombre_para_guardar.strip() or "Pareto Comunidad"),
        key="inf_editor_nombre"
    )
    desgloses = ui_desgloses(tabla["descriptor"].tolist(), key_prefix="editor")
    col_inf1, col_inf2 = st.columns([1,3])
    with col_inf1:
        gen = st.button("📄 Generar Informe PDF (editor)", type="primary", use_container_width=True, key="btn_inf_editor")
    with col_inf2:
        st.caption("El PDF incluirá: Portada, Resultados, Diagrama de Pareto y cada sección de Modalidades.")

    if gen:
        if tabla.empty:
            st.warning("No hay datos para el informe. Asigna frecuencias primero.")
        else:
            pdf_bytes = generar_pdf_informe(nombre_informe, tabla, desgloses)
            st.success("Informe generado.")
            st.download_button(
                "⬇️ Descargar Informe PDF",
                data=pdf_bytes,
                file_name=f"informe_{nombre_informe.lower().replace(' ','_')}.pdf",
                mime="application/pdf",
                key="dl_inf_editor",
            )

else:
    st.info("Selecciona al menos un descriptor para continuar. Tus frecuencias se conservarán si luego agregas más descriptores.")

# ==========================================
# 8) PORTAFOLIO, UNIFICADO Y DESCARGAS
# ==========================================
st.markdown("---")
st.header("📁 Portafolio de Paretos (guardados)")

if not st.session_state["portafolio"]:
    st.info("Aún no hay paretos guardados. Guarda el primero desde la sección anterior.")
else:
    st.subheader("Selecciona paretos para Unificar")
    nombres = sorted(st.session_state["portafolio"].keys())
    sel_unif = st.multiselect("Elige 2 o más paretos para combinar (o usa el botón de 'Unificar todos')",
                              options=nombres, default=[], key="sel_unif")

    c_unif1, c_unif2 = st.columns([1,1])
    with c_unif1: unificar_todos = st.button("🔗 Unificar TODOS los paretos guardados")
    with c_unif2: st.caption(f"Total de paretos guardados: **{len(nombres)}**")

    st.markdown("### Paretos guardados")
    for nom in nombres:
        freq_map = st.session_state["portafolio"][nom]
        meta = info_pareto(freq_map)
        with st.expander(f"🔹 {nom} — {meta['descriptores']} descriptores | Total: {meta['total']}"):
            df_base = df_desde_freq_map(freq_map)
            tabla_g = calcular_pareto(df_base)

            mostrar_g = tabla_g.copy()[["categoria","descriptor","frecuencia","porcentaje","pct_acum","acumulado","segmento"]]
            mostrar_g = mostrar_g.rename(columns={"pct_acum":"porcentaje acumulado"})
            if not mostrar_g.empty:
                mostrar_g["porcentaje"] = mostrar_g["porcentaje"].map(lambda x: f"{x:.2f}%")
                mostrar_g["porcentaje acumulado"] = mostrar_g["porcentaje acumulado"].map(lambda x: f"{x:.2f}%")


            cc1, cc2, cc3 = st.columns([1,1,1])
            with cc1:
                if not mostrar_g.empty:
                    st.dataframe(mostrar_g, use_container_width=True, hide_index=True)
                else:
                    st.info("Este pareto no tiene frecuencias > 0.")
            with cc2:
                st.markdown("**Gráfico**"); dibujar_pareto(tabla_g, f"Pareto — {nom}")
            with cc3:
                st.markdown("**Acciones**")
                if not tabla_g.empty:
                    st.download_button(
                        "⬇️ Excel de este Pareto",
                        data=exportar_excel_con_grafico(tabla_g, f"Pareto — {nom}"),
                        file_name=f"pareto_{nom.lower().replace(' ','_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{nom}",
                    )
                if st.button("📥 Cargar este Pareto al editor", key=f"load_{nom}"):
                    st.session_state["freq_map"] = dict(freq_map)
                    st.session_state["msel"] = list(freq_map.keys())
                    st.success(f"Pareto '{nom}' cargado al editor (arriba). Desplázate para editar.")
                # Informe directo de este pareto guardado
                with st.popover("📄 Informe PDF de este Pareto"):
                    nombre_inf_ind = st.text_input("Nombre del informe", value=f"Pareto — {nom}", key=f"inf_nom_{nom}")
                    desgloses_ind = ui_desgloses(tabla_g["descriptor"].tolist(), key_prefix=f"inf_{nom}")
                    if st.button("Generar PDF", key=f"btn_inf_{nom}"):
                        pdf_bytes = generar_pdf_informe(nombre_inf_ind, tabla_g, desgloses_ind)
                        st.download_button(
                            "⬇️ Descargar PDF",
                            data=pdf_bytes,
                            file_name=f"informe_{nom.lower().replace(' ','_')}.pdf",
                            mime="application/pdf",
                            key=f"dl_inf_{nom}",
                        )
                if st.button("🗑️ Eliminar de la sesión", key=f"del_{nom}"):
                    try:
                        del st.session_state["portafolio"][nom]
                        st.warning(f"Pareto '{nom}' eliminado del portafolio de la sesión.")
                        st.rerun()
                    except Exception:
                        st.error("No se pudo eliminar. Intenta de nuevo.")

    # ---------- UNIFICADO ----------
    st.markdown("---"); st.header("🔗 Pareto Unificado (por filtro o general)")
    maps_a_unir = []; titulo_unif = ""
    if unificar_todos and nombres:
        maps_a_unir = [st.session_state["portafolio"][n] for n in nombres]
        titulo_unif = "Pareto General (todos los paretos)"
    elif len(st.session_state.get("sel_unif", [])) >= 2:
        maps_a_unir = [st.session_state["portafolio"][n] for n in st.session_state["sel_unif"]]
        titulo_unif = f"Unificado: {', '.join(st.session_state['sel_unif'])}"
    if maps_a_unir:
        combinado = combinar_maps(maps_a_unir)
        df_unif = df_desde_freq_map(combinado)
        tabla_unif = calcular_pareto(df_unif)
        mostrar_u = tabla_unif.copy()[["categoria","descriptor","frecuencia","porcentaje","pct_acum","acumulado","segmento"]]
        mostrar_u = mostrar_u.rename(columns={"pct_acum":"porcentaje acumulado"})
        if not mostrar_u.empty:
            mostrar_u["porcentaje"] = mostrar_u["porcentaje"].map(lambda x: f"{x:.2f}%")
            mostrar_u["porcentaje acumulado"] = mostrar_u["porcentaje acumulado"].map(lambda x: f"{x:.2f}%")
        cu1, cu2 = st.columns([1,1], gap="large")
        with cu1:
            st.markdown("**Tabla Unificada**")
            if not mostrar_u.empty:
                st.dataframe(mostrar_u, use_container_width=True, hide_index=True)
            else:
                st.info("Sin datos > 0 en la combinación seleccionada.")
        with cu2:
            st.markdown("**Gráfico Unificado**"); dibujar_pareto(tabla_unif, titulo_unif or "Pareto Unificado")
        if not tabla_unif.empty:
            st.download_button(
                "⬇️ Descargar Excel del Pareto Unificado",
                data=exportar_excel_con_grafico(tabla_unif, titulo_unif or "Pareto Unificado"),
                file_name="pareto_unificado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_unificado",
            )

        # ===== Informe PDF Unificado / General =====
        st.markdown("### 🧾 Elaborar Informe PDF (Unificado / General)")
        nombre_inf_unif = st.text_input("Nombre del informe",
                                        value=(titulo_unif or "Pareto Unificado"),
                                        key="inf_unif_nombre")
        desgloses_unif = ui_desgloses(tabla_unif["descriptor"].tolist(), key_prefix="unif")
        if st.button("📄 Generar Informe PDF (unificado)", key="btn_inf_unif"):
            pdf_bytes = generar_pdf_informe(nombre_inf_unif, tabla_unif, desgloses_unif)
            st.download_button(
                "⬇️ Descargar Informe PDF",
                data=pdf_bytes,
                file_name=f"informe_{(titulo_unif or 'Pareto Unificado').lower().replace(' ','_')}.pdf",
                mime="application/pdf",
                key="dl_inf_unif",
            )
    else:
        st.info("Selecciona 2+ paretos en el multiselect o usa el botón 'Unificar TODOS' para habilitar el unificado.")








