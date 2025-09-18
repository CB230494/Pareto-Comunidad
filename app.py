# app.py — Pareto con columnas reordenadas y segmento fijo "80%"
# ----------------------------------------------------------------
# Requisitos:
#   pip install streamlit pandas matplotlib xlsxwriter
#   streamlit run app.py
# ----------------------------------------------------------------

import io
from typing import List, Dict

import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

st.set_page_config(page_title="Pareto de Descriptores", layout="wide")

# =========================
# 1) Catálogo embebido (normalizado)
# =========================
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

# =========================
# 2) Utilidades
# =========================
ORANGE = "#FF8C00"  # naranja vivo

def calcular_pareto(df_in: pd.DataFrame) -> pd.DataFrame:
    """Calcula Pareto y deja 'segmento' fijo en '80%' para todas las filas."""
    df = df_in.copy()
    df["frecuencia"] = pd.to_numeric(df["frecuencia"], errors="coerce").fillna(0).astype(int)
    df = df[df["frecuencia"] > 0]
    if df.empty:
        return df.assign(porcentaje=0.0, acumulado=0, pct_acum=0.0, segmento="80%")

    df = df.sort_values("frecuencia", ascending=False)
    total = int(df["frecuencia"].sum())
    df["porcentaje"] = (df["frecuencia"] / total * 100).round(2)
    df["acumulado"]  = df["frecuencia"].cumsum()
    df["pct_acum"]   = (df["acumulado"] / total * 100).round(2)
    df["segmento"]   = "80%"  # fijo según lo solicitado
    return df.reset_index(drop=True)

def dibujar_pareto(df_par: pd.DataFrame, titulo: str):
    if df_par.empty:
        st.info("Ingresa frecuencias (>0) para ver el gráfico.")
        return

    x        = np.arange(len(df_par))
    freqs    = df_par["frecuencia"].to_numpy()
    pct_acum = df_par["pct_acum"].to_numpy()

    fig, ax1 = plt.subplots(figsize=(14, 5))
    ax1.bar(x, freqs, color=ORANGE)  # todo naranja (segmento 80%)
    ax1.set_ylabel("Frecuencia")
    ax1.set_xticks(x)
    ax1.set_xticklabels(df_par["descriptor"].tolist(), rotation=75, ha="right")
    ax1.set_title(titulo if titulo.strip() else "Pareto — Frecuencia y % acumulado")

    ax2 = ax1.twinx()
    ax2.plot(x, pct_acum, marker="o")
    ax2.set_ylabel("% acumulado")
    ax2.set_ylim(0, 110)
    ax2.axhline(80, linestyle="--")  # referencia 80%

    st.pyplot(fig)

def exportar_excel_con_grafico(df_par: pd.DataFrame, titulo: str) -> bytes:
    """XLSX con columnas reordenadas y gráfico con barras naranjas."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        hoja = "Pareto"

        # Reordenar y formatear antes de exportar
        df_x = df_par.copy()
        # Para Excel: % como fracción 0-1 para aplicar formato 0.00%
        df_x["porcentaje"] = (df_x["porcentaje"] / 100.0).round(4)
        df_x["pct_acum"]   = (df_x["pct_acum"] / 100.0).round(4)

        # Reordenar columnas: categoría, descriptor, frecuencia, porcentaje, porcentaje acumulado, acumulado, segmento
        df_x = df_x[["categoria", "descriptor", "frecuencia", "porcentaje", "pct_acum", "acumulado", "segmento"]]

        df_x.to_excel(writer, sheet_name=hoja, index=False, startrow=0, startcol=0)
        wb = writer.book
        ws = writer.sheets[hoja]

        # Formatos de columnas
        pct_fmt = wb.add_format({"num_format": "0.00%"})
        ws.set_column("A:A", 18)   # categoria
        ws.set_column("B:B", 55)   # descriptor
        ws.set_column("C:C", 12)   # frecuencia
        ws.set_column("D:D", 12, pct_fmt)  # porcentaje
        ws.set_column("E:E", 16, pct_fmt)  # porcentaje acumulado
        ws.set_column("F:F", 12)   # acumulado
        ws.set_column("G:G", 10)   # segmento

        n = len(df_x)
        # Rangos para el gráfico con el nuevo orden:
        cats = f"=Pareto!$B$2:$B${n+1}"   # descriptor (columna B)
        vals = f"=Pareto!$C$2:$C${n+1}"   # frecuencia (columna C)
        pcts = f"=Pareto!$E$2:$E${n+1}"   # porcentaje acumulado (columna E)

        # Total
        total = int(df_par["frecuencia"].sum())
        ws.write(n+2, 1, "TOTAL:")
        ws.write(n+2, 2, total)

        # Gráfico
        chart = wb.add_chart({"type": "column"})
        chart.add_series({
            "name": "Frecuencia",
            "categories": cats,
            "values": vals,
            "fill": {"color": ORANGE},
        })

        line = wb.add_chart({"type": "line"})
        line.add_series({
            "name": "% acumulado",
            "categories": cats,
            "values": pcts,
            "y2_axis": True,
            "marker": {"type": "circle"},
        })

        chart.combine(line)
        chart.set_y_axis({"name": "Frecuencia"})
        chart.set_y2_axis({"name": "% acumulado", "min": 0, "max": 1.10, "major_unit": 0.10, "num_format": "0%"})
        chart.set_title({"name": titulo if titulo.strip() else "PARETO – Frecuencia y % acumulado"})
        chart.set_legend({"position": "bottom"})
        chart.set_size({"width": 1180, "height": 420})
        ws.insert_chart("I2", chart)

    return output.getvalue()

# =========================
# 3) Estado (persistencia) y UI
# =========================
if "freq_map" not in st.session_state:
    st.session_state.freq_map = {}  # {descriptor: frecuencia}

st.title("Pareto de Descriptores")
st.caption("Columnas ordenadas y segmento fijado a '80%' según lo solicitado.")

titulo = st.text_input("Título del Pareto (opcional)", value="Pareto Comunidad")

cat_df = pd.DataFrame(CATALOGO).sort_values(["categoria", "descriptor"]).reset_index(drop=True)
opciones = cat_df["descriptor"].tolist()
seleccion = st.multiselect("1) Escoge uno o varios descriptores", options=opciones, default=[])

st.subheader("2) Asigna la frecuencia")
if seleccion:
    base = cat_df[cat_df["descriptor"].isin(seleccion)].copy()
    base["frecuencia"] = [st.session_state.freq_map.get(d, 0) for d in base["descriptor"]]

    edit = st.data_editor(
        base,
        key="editor_freq",
        num_rows="fixed",
        use_container_width=True,
        column_config={
            "descriptor": st.column_config.TextColumn("DESCRIPTOR", width="large"),
            "categoria": st.column_config.TextColumn("CATEGORÍA", width="small"),
            "frecuencia": st.column_config.NumberColumn("Frecuencia", min_value=0, step=1),
        },
    )

    # Persistir cambios
    for _, row in edit.iterrows():
        st.session_state.freq_map[row["descriptor"]] = int(row["frecuencia"])

    df_in = edit[["descriptor", "categoria"]].copy()
    df_in["frecuencia"] = df_in["descriptor"].map(st.session_state.freq_map).fillna(0).astype(int)

    st.subheader("3) Pareto")
    tabla = calcular_pareto(df_in)

    # ---- Reordenar/renombrar para mostrar en la app ----
    mostrar = tabla.copy()
    mostrar = mostrar[["categoria", "descriptor", "frecuencia", "porcentaje", "pct_acum", "acumulado", "segmento"]]
    mostrar = mostrar.rename(columns={"pct_acum": "porcentaje acumulado"})
    # % con 2 decimales
    mostrar["porcentaje"] = mostrar["porcentaje"].map(lambda x: f"{x:.2f}%")
    mostrar["porcentaje acumulado"] = mostrar["porcentaje acumulado"].map(lambda x: f"{x:.2f}%")

    c1, c2 = st.columns([1, 1], gap="large")
    with c1:
        st.markdown("**Tabla de Pareto**")
        if tabla.empty:
            st.info("Ingresa frecuencias (>0) para ver la tabla.")
        else:
            st.dataframe(mostrar, use_container_width=True, hide_index=True)

    with c2:
        st.markdown("**Gráfico de Pareto**")
        dibujar_pareto(tabla, titulo)

    st.subheader("4) Exportar")
    if not tabla.empty:
        # CSV mantiene nombres internos; si quieres los mismos nombres visibles, exporta 'mostrar'
        st.download_button(
            "⬇️ Descargar CSV",
            data=mostrar.to_csv(index=False).encode("utf-8"),
            file_name="pareto_descriptores.csv",
            mime="text/csv",
        )
        st.download_button(
            "⬇️ Descargar Excel con gráfico",
            data=exportar_excel_con_grafico(tabla, titulo),
            file_name="pareto_descriptores.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Selecciona al menos un descriptor para continuar. Tus frecuencias se conservarán si luego agregas más descriptores.")



