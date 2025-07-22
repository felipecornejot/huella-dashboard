import streamlit as st
import pandas as pd

# Configuración de la página
st.set_page_config(page_title="Huella de Carbono - Dashboard", layout="wide")  
# (layout="wide" aprovecha todo el ancho de pantalla para el dashboard)

# Título del dashboard
st.title("📊 Dashboard de Huella de Carbono Organizacional - SustREND")
st.markdown("Este dashboard interactivo muestra las **emisiones de GEI** y métricas asociadas, basado en los datos de la huella de carbono organizacional calculada para el Comité de Desarrollo Productivo Regional Bío-Bío (2024).")

# Widget para cargar archivo Excel (permite múltiples archivos para comparar años)
st.sidebar.header("📁 Cargar datos")
uploaded_files = st.sidebar.file_uploader(
    "Sube la planilla Excel de huella de carbono (*.xlsx)", 
    type=["xlsx"], accept_multiple_files=True, 
    help="Cargue uno o varios archivos Excel con el cálculo de huella de carbono. Si no se cargan archivos, se usarán datos de ejemplo."
)
# Nota: accept_multiple_files=True permite seleccionar varios archivos a la vez:contentReference[oaicite:14]{index=14}

# Función auxiliar para leer el Excel y extraer datos relevantes
def cargar_datos_desde_excel(file):
    # Lee hojas necesarias
    xls = pd.ExcelFile(file)
    # Hoja "Resumen Cálculo" contiene totales por categoría y alcance
    resumen_df = pd.read_excel(xls, sheet_name="Resumen Cálculo", skiprows=2, usecols="A:C") 
    # Las columnas A, B, C tienen: Alcance/Categoría, Emisión tCO2e, % del total.
    resumen_df.columns = ["Categoria", "Emision_tCO2e", "Porcentaje"]
    resumen_df.dropna(inplace=True)  # quitar filas vacías
    # Extraer el total de emisiones (tCO2e) de la última fila:
    total_emisiones = 0.0
    if "total" in resumen_df.iloc[-1]["Categoria"].lower():
        total_emisiones = resumen_df.iloc[-1]["Emision_tCO2e"]
        resumen_df = resumen_df.iloc[:-1]  # quitar la fila de total para listar solo categorías
    else:
        total_emisiones = resumen_df["Emision_tCO2e"].sum()
    # Leer hoja "Alcances" para obtener datos de año, empleados, superficie
    alcances_df = pd.read_excel(xls, sheet_name="Alcances", usecols="A:B", nrows=30)
    alcances_df = alcances_df.fillna("")  # reemplazar NaN con cadena vacía
    # Buscar año inventario, trabajadores y superficie
    ano = None
    empleados = None
    superficie = None
    for i, row in alcances_df.iterrows():
        celda = str(row[0]).strip().lower()
        if "año inventario" in celda:
            try:
                ano = int(float(row[1]))  # convertir a entero
            except:
                ano = str(row[1]).strip()
        if "trabajadores" in celda:
            try:
                empleados = float(row[1])
            except:
                empleados = None
        if "superficie" in celda:
            try:
                superficie = float(row[1])
            except:
                superficie = None
    # Leer hoja "Resumen Emisiones" para detalles por fuente (usando fila 1 como cabecera)
    emisiones_df = pd.read_excel(xls, sheet_name="Resumen Emisiones", header=1)
    # Renombrar columnas clave para facilitar manejo
    emisiones_df = emisiones_df.rename(columns={
        "Emisión CO2e\n[t]": "Emisiones (tCO2e)",
        "% del total": "% del total"
    })
    # Agregar columna de año para identificación si se combinan múltiples archivos
    emisiones_df["Año"] = ano
    return {
        "ano": ano,
        "total": total_emisiones,
        "empleados": empleados,
        "superficie": superficie,
        "resumen": resumen_df,
        "emisiones_detalle": emisiones_df
    }

# Cargar datos ya sea desde archivos subidos o desde el archivo de ejemplo incluido
data_list = []
if uploaded_files:
    for file in uploaded_files:
        try:
            data = cargar_datos_desde_excel(file)
            data_list.append(data)
        except Exception as e:
            st.error(f"⚠️ Error al leer {file.name}: {e}")
else:
    # Si no se subió archivo, usar datos de ejemplo (Excel incorporado en el repositorio)
    try:
        data = cargar_datos_desde_excel("Calculadora_Memoria_Huella.xlsx")
        data_list.append(data)
        st.sidebar.info("Usando datos de ejemplo 2024 (CDPR Bío-Bío) - cargue un archivo en la barra lateral para actualizar.")
    except Exception as e:
        st.sidebar.error(f"No se pudo cargar el archivo de ejemplo: {e}")

if not data_list:
    st.stop()  # Si no hay datos, detener la ejecución

# Si se cargó más de un año, preparar comparación
if len(data_list) > 1:
    st.header("📈 Comparación entre Años")
    # Crear dataframe de comparación de totales e intensidad
    comp_df = pd.DataFrame([{"Año": d["ano"], "Emisiones tCO2e": d["total"], 
                              "tCO2e/empleado": (d["total"]/d["empleados"]) if d["empleados"] else None,
                              "tCO2e/m2": (d["total"]/d["superficie"]) if d["superficie"] else None
                             } for d in data_list])
    comp_df = comp_df.sort_values("Año")
    # Mostrar tabla comparativa de indicadores por año
    st.dataframe(comp_df.set_index("Año"), use_container_width=True)
    # Gráfico de barras de emisiones totales por año
    st.bar_chart(data=comp_df.set_index("Año")["Emisiones tCO2e"], use_container_width=True)
    st.markdown("En el gráfico anterior se observa la tendencia de las emisiones totales por año. A continuación se detallan los resultados del año más reciente.")

# Considerar solo el último conjunto de datos (más reciente) para el resto del análisis
data = data_list[-1]
ano = data["ano"]
total_emisiones = data["total"]
empleados = data["empleados"]
superficie = data["superficie"]
resumen_df = data["resumen"]
emisiones_df = data["emisiones_detalle"]

# Título de sección para el año específico
st.header(f"Resultados del Inventario {ano}")
st.markdown(f"**Emisiones GEI totales:** {total_emisiones:.2f} tCO2e en el año {ano}.") 

# Mostrar indicadores de intensidad si disponibles
st.subheader("Indicadores de Intensidad")
cols = st.columns(3)
if empleados:
    cols[0].metric("Emisiones por trabajador", f"{(total_emisiones/empleados):.2f} tCO2e/persona")
else:
    cols[0].metric("Emisiones por trabajador", "N/A")
if superficie:
    cols[1].metric("Emisiones por superficie", f"{(total_emisiones/superficie):.2f} tCO2e/m²")
else:
    cols[1].metric("Emisiones por superficie", "N/A")
cols[2].metric("Emisiones totales", f"{total_emisiones:.2f} tCO2e") 
# Los metric() muestran valores destacados en formato tarjeta para rápida visualización

# Gráficos de emisiones por alcance / categoría
st.subheader("Distribución de Emisiones por Categoría")
# Separar emisiones por alcances 1, 2, 3 sumando categorías correspondientes
alcance1 = resumen_df[resumen_df["Categoria"].str.startswith("Alcance 1")]["Emision_tCO2e"].sum()
alcance2 = resumen_df[resumen_df["Categoria"].str.startswith("Alcance 2")]["Emision_tCO2e"].sum()
alcance3 = resumen_df[resumen_df["Categoria"].str.startswith("Alcance 3")]["Emision_tCO2e"].sum()
alcances_data = pd.DataFrame({
    "Alcance": ["Alcance 1", "Alcance 2", "Alcance 3"],
    "Emisiones (tCO2e)": [alcance1, alcance2, alcance3]
})
# Grafico de barras de emisiones por alcance
st.plotly_chart(
    # Usamos Plotly Express para personalizar colores y estilo fácilmente
    px.bar(alcances_data, x="Alcance", y="Emisiones (tCO2e)", 
           color="Alcance", color_discrete_sequence=["#4CAF50","#2196F3","#FFC107"],
           title="Emisiones por Alcance (tCO2e)"),
    use_container_width=True
)
st.write("Alcance 3 (otras emisiones indirectas) representa el porcentaje mayor del total, seguido por Alcance 1 y 2, acorde con el informe donde Alcance 3 aportó ~58,7%:contentReference[oaicite:15]{index=15}.")

# Grafico de emisiones por categoría principal (filtrando categorías incluidas con emisiones > 0)
categorias_plot = resumen_df[resumen_df["Emision_tCO2e"] > 0].copy()
# Excluir filas de títulos "Alcance X" para solo graficar subcategorías:
categorias_plot = categorias_plot[~categorias_plot["Categoria"].str.contains("Alcance")]
# Ordenar por emisiones descendente
categorias_plot = categorias_plot.sort_values("Emision_tCO2e", ascending=False)
fig_categorias = px.bar(categorias_plot, x="Categoria", y="Emision_tCO2e", 
                        title="Emisiones por Categoría/Subcategoría",
                        labels={"Emision_tCO2e":"Emisiones (tCO2e)"},
                        color="Emision_tCO2e", color_continuous_scale="Blues")
fig_categorias.update_xaxes(categoryorder='total descending')
st.plotly_chart(fig_categorias, use_container_width=True)
st.write("En el gráfico anterior se aprecia el detalle de emisiones por subcategoría. Por ejemplo, '3.07 - Traslado de Colaboradores' (transporte de empleados) y '2.1 - Compra de energía' destacan como categorías significativas en este inventario.")

# Gráficos por fuente de emisión y tipo de GEI
st.subheader("Emisiones por Fuente Específica")
# Calcular emisiones totales por cada fuente (columna 'Fuente de Emisión')
top_fuentes = emisiones_df.groupby("Fuente de Emisión")["Emisiones (tCO2e)"].sum().sort_values(ascending=False)
top5 = top_fuentes.head(5)
# Bar chart de las 5 principales fuentes de emisión
fig_fuentes = px.bar(x=top5.values, y=top5.index, orientation='h', 
                     labels={'x':'Emisiones (tCO2e)', 'y':'Fuente de Emisión'},
                     title="Top 5 Fuentes de Emisión de GEI")
fig_fuentes.update_yaxes(categoryorder="total ascending")  # ordenar de mayor a menor
st.plotly_chart(fig_fuentes, use_container_width=True)
st.write(f"Las cinco fuentes más importantes de {ano} explican gran parte de las emisiones totales. En este caso, se observa que **{top5.index[0]}** y **{top5.index[1]}** son las fuentes principales, contribuyendo con aproximadamente {top5.iloc[0]:.1f} y {top5.iloc[1]:.1f} tCO2e respectivamente:contentReference[oaicite:16]{index=16}.") 

# (Opcional) Gráfico de proporción por tipo de GEI en Alcance 1 (CO2, CH4, N2O)
if alcance1 > 0:
    st.subheader("Composición por Tipo de GEI (Alcance 1)")
    # Calcular CO2, CH4, N2O para Alcance 1 usando datos detallados (solo emisiones directas)
    directos = emisiones_df[emisiones_df["Alcance"].str.contains("Alcance 1")]
    co2_kg = directos["Emision CO2 \n[kg]"].fillna(0).sum()    if "Emision CO2 \n[kg]" in directos.columns else 0
    ch4_kg_co2e = directos["Emisión CH4 \n[kg]"].fillna(0).sum()  if "Emisión CH4 \n[kg]" in directos.columns else 0
    n2o_kg_co2e = directos["Emisión N2O \n[kg]"].fillna(0).sum()  if "Emisión N2O \n[kg]" in directos.columns else 0
    # Nota: en este Excel, las columnas CH4 y N2O están en [kg CO2e] de cada gas
    total_co2e_direct = co2_kg/1000 + ch4_kg_co2e/1000 + n2o_kg_co2e/1000  # en toneladas CO2e
    tipos = pd.DataFrame({
        "Tipo de GEI": ["CO2 (fósil)", "CH4 (eq)", "N2O (eq)"],
        "Contribución (tCO2e)": [co2_kg/1000, ch4_kg_co2e/1000, n2o_kg_co2e/1000]
    })
    fig_gases = px.pie(tipos, values="Contribución (tCO2e)", names="Tipo de GEI", 
                       title="Proporción por tipo de gas en emisiones directas")
    st.plotly_chart(fig_gases, use_container_width=True)
    st.write("Las emisiones directas (Alcance 1) están compuestas principalmente por CO₂ fósil. Las contribuciones de CH₄ y N₂O son marginales en comparación (p. ej., N₂O aporta alrededor del 2% de las emisiones directas en CO₂e).") 

# Tabla de emisiones detallada con filtros
st.subheader("📋 Detalle de Emisiones por Fuente (Tabla Interactiva)")
# Filtros de alcance
alcances_opt = sorted(emisiones_df["Alcance"].unique())
alcance_sel = st.multiselect("Filtrar por alcance:", options=alcances_opt, default=alcances_opt)
# Filtros de categoría (opcional, a partir de las categorías codificadas tipo '1.2 - Combustión móvil')
categorias_opt = sorted(emisiones_df["Categoría"].unique())
cat_sel = st.multiselect("Filtrar por categoría específica:", options=categorias_opt, default=categorias_opt)
# Aplicar filtros al dataframe
df_filtrado = emisiones_df[ (emisiones_df["Alcance"].isin(alcance_sel)) & (emisiones_df["Categoría"].isin(cat_sel)) ]
# Seleccionar columnas relevantes para mostrar
columnas_mostrar = ["Alcance", "Categoría", "Fuente de Emisión", "Dato de Actividad\n(DA)", "Unidad DA", "Emisiones (tCO2e)", "% del total"]
tabla = df_filtrado[columnas_mostrar]
st.dataframe(tabla, use_container_width=True)
# Botón de descarga de la tabla filtrada como CSV
csv_data = tabla.to_csv(index=False).encode('utf-8')
st.download_button("💾 Exportar resultados filtrados a CSV", data=csv_data, file_name=f"emisiones_filtradas_{ano}.csv", mime="text/csv")
# (Al hacer clic en el botón, se descargará un archivo CSV con los datos filtrados):contentReference[oaicite:17]{index=17}

# Pie de página con logo SustREND
st.markdown("---")
footer = st.container()
with footer:
    col1, col2, col3 = st.columns([1,1,1])
    with col2:
        # Centrar el logo en la columna central
        st.image("sustrend_logo.png", use_column_width=True)
    st.write("<center><em>Dashboard elaborado por SustREND - Datos Programa HuellaChile, {}</em></center>".format(ano), unsafe_allow_html=True)
