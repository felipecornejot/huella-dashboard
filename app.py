import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Huella de Carbono - Dashboard", layout="wide")

st.title("📊 Dashboard de Huella de Carbono Organizacional - SustREND")
st.markdown("Este dashboard interactivo muestra las **emisiones de GEI** y métricas asociadas, basado en los datos de la huella de carbono organizacional calculada para el Comité de Desarrollo Productivo Regional Bío-Bío (2024).")

st.sidebar.header("📁 Cargar datos")
uploaded_files = st.sidebar.file_uploader(
    "Sube la planilla Excel de huella de carbono (*.xlsx)", 
    type=["xlsx"], accept_multiple_files=True,
    help="Cargue uno o varios archivos Excel con el cálculo de huella de carbono. Si no se cargan archivos, se usan datos de ejemplo."
)

def cargar_datos_desde_excel(file):
    xls = pd.ExcelFile(file)
    resumen_df = pd.read_excel(xls, sheet_name="Resumen Cálculo", skiprows=2, usecols="A:C")
    resumen_df.columns = ["Categoria", "Emision_tCO2e", "Porcentaje"]
    resumen_df.dropna(inplace=True)
    resumen_df["Emision_tCO2e"] = pd.to_numeric(resumen_df["Emision_tCO2e"], errors='coerce')

    total_emisiones = resumen_df["Emision_tCO2e"].sum()
    alcances_df = pd.read_excel(xls, sheet_name="Alcances", usecols="A:B", nrows=30).fillna("")
    ano, empleados, superficie = None, None, None
    for _, row in alcances_df.iterrows():
        celda = str(row[0]).strip().lower()
        if "año inventario" in celda:
            try: ano = int(float(row[1]))
            except: ano = str(row[1]).strip()
        if "trabajadores" in celda:
            try: empleados = float(row[1])
            except: empleados = None
        if "superficie" in celda:
            try: superficie = float(row[1])
            except: superficie = None
    emisiones_df = pd.read_excel(xls, sheet_name="Resumen Emisiones", header=1)
    emisiones_df.rename(columns={"Emisión CO2e\n[t]": "Emisiones (tCO2e)", "% del total": "% del total"}, inplace=True)
    emisiones_df["Emisiones (tCO2e)"] = pd.to_numeric(emisiones_df["Emisiones (tCO2e)"], errors='coerce')
    emisiones_df["Año"] = ano
    return {"ano": ano, "total": total_emisiones, "empleados": empleados, "superficie": superficie,
            "resumen": resumen_df, "emisiones_detalle": emisiones_df}

data_list = []
if uploaded_files:
    for file in uploaded_files:
        try:
            data = cargar_datos_desde_excel(file)
            data_list.append(data)
        except Exception as e:
            st.error(f"⚠️ Error al leer {file.name}: {e}")
else:
    try:
        data = cargar_datos_desde_excel("Calculadora_Memoria_Huella.xlsx")
        data_list.append(data)
        st.sidebar.info("Usando datos de ejemplo 2024 (CDPR Bío-Bío).")
    except Exception as e:
        st.sidebar.error(f"No se pudo cargar archivo de ejemplo: {e}")

if not data_list:
    st.stop()

if len(data_list) > 1:
    st.header("📈 Comparación entre Años")
    comp_df = pd.DataFrame([{"Año": d["ano"], "Emisiones tCO2e": d["total"],
                             "tCO2e/empleado": (d["total"]/d["empleados"]) if d["empleados"] else None,
                             "tCO2e/m2": (d["total"]/d["superficie"]) if d["superficie"] else None}
                            for d in data_list]).sort_values("Año")
    st.dataframe(comp_df.set_index("Año"), use_container_width=True)
    st.bar_chart(comp_df.set_index("Año")["Emisiones tCO2e"])
    st.markdown("Tendencia de emisiones totales por año. Abajo, el detalle del año más reciente.")

data = data_list[-1]
ano, total_emisiones, empleados, superficie = data["ano"], data["total"], data["empleados"], data["superficie"]
resumen_df, emisiones_df = data["resumen"], data["emisiones_detalle"]

st.header(f"Resultados del Inventario {ano}")
st.markdown(f"**Emisiones GEI totales:** {total_emisiones:.2f} tCO2e en el año {ano}.")

st.subheader("Indicadores de Intensidad")
cols = st.columns(3)
cols[0].metric("Emisiones por trabajador", f"{(total_emisiones/empleados):.2f} tCO2e/persona" if empleados else "N/A")
cols[1].metric("Emisiones por superficie", f"{(total_emisiones/superficie):.2f} tCO2e/m²" if superficie else "N/A")
cols[2].metric("Emisiones totales", f"{total_emisiones:.2f} tCO2e")

st.subheader("Distribución de Emisiones por Categoría")
alcance1 = resumen_df[resumen_df["Categoria"].str.startswith("Alcance 1")]["Emision_tCO2e"].sum()
alcance2 = resumen_df[resumen_df["Categoria"].str.startswith("Alcance 2")]["Emision_tCO2e"].sum()
alcance3 = resumen_df[resumen_df["Categoria"].str.startswith("Alcance 3")]["Emision_tCO2e"].sum()
alcances_data = pd.DataFrame({"Alcance": ["Alcance 1", "Alcance 2", "Alcance 3"],
                              "Emisiones (tCO2e)": [alcance1, alcance2, alcance3]})
st.plotly_chart(px.bar(alcances_data, x="Alcance", y="Emisiones (tCO2e)",
                       color="Alcance", color_discrete_sequence=["#4CAF50", "#2196F3", "#FFC107"],
                       title="Emisiones por Alcance (tCO2e)"), use_container_width=True)

categorias_plot = resumen_df[~resumen_df["Categoria"].str.contains("Alcance") & (resumen_df["Emision_tCO2e"] > 0)]
categorias_plot = categorias_plot.sort_values("Emision_tCO2e", ascending=False)
fig_categorias = px.bar(categorias_plot, x="Categoria", y="Emision_tCO2e",
                        title="Emisiones por Categoría/Subcategoría",
                        labels={"Emision_tCO2e": "Emisiones (tCO2e)"},
                        color="Emision_tCO2e", color_continuous_scale="Blues")
fig_categorias.update_xaxes(categoryorder='total descending')
st.plotly_chart(fig_categorias, use_container_width=True)

st.subheader("Emisiones por Fuente Específica")
top_fuentes = emisiones_df.groupby("Fuente de Emisión")["Emisiones (tCO2e)"].sum().sort_values(ascending=False)
top5 = top_fuentes.head(5)
fig_fuentes = px.bar(x=top5.values, y=top5.index, orientation='h',
                     labels={'x': 'Emisiones (tCO2e)', 'y': 'Fuente de Emisión'},
                     title="Top 5 Fuentes de Emisión de GEI")
fig_fuentes.update_yaxes(categoryorder="total ascending")
st.plotly_chart(fig_fuentes, use_container_width=True)
st.write(f"Las cinco fuentes más importantes de {ano} explican gran parte de las emisiones totales. En este caso, se observa que **{top5.index[0]}** y **{top5.index[1]}** son las fuentes principales, contribuyendo con aproximadamente {top5.iloc[0]:.1f} y {top5.iloc[1]:.1f} tCO2e respectivamente.")

st.subheader("📋 Detalle de Emisiones por Fuente (Tabla Interactiva)")
alcances_opt = sorted(emisiones_df["Alcance"].unique())
alcance_sel = st.multiselect("Filtrar por alcance:", options=alcances_opt, default=alcances_opt)
categorias_opt = sorted(emisiones_df["Categoría"].unique())
cat_sel = st.multiselect("Filtrar por categoría específica:", options=categorias_opt, default=categorias_opt)
df_filtrado = emisiones_df[(emisiones_df["Alcance"].isin(alcance_sel)) & (emisiones_df["Categoría"].isin(cat_sel))]
columnas_mostrar = ["Alcance", "Categoría", "Fuente de Emisión", "Dato de Actividad\n(DA)", "Unidad DA", "Emisiones (tCO2e)", "% del total"]
tabla = df_filtrado[columnas_mostrar]
st.dataframe(tabla, use_container_width=True)
csv_data = tabla.to_csv(index=False).encode('utf-8')
st.download_button("💾 Exportar resultados filtrados a CSV", data=csv_data, file_name=f"emisiones_filtradas_{ano}.csv", mime="text/csv")

st.markdown("---")
col1, col2, col3 = st.columns([1,1,1])
with col2:
    st.image("sustrend_logo.png", use_column_width=True)
st.write("<center><em>Dashboard elaborado por SustREND - Datos Programa HuellaChile, {}</em></center>".format(ano), unsafe_allow_html=True)
