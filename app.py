import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# Cargar el archivo Excel
file_path = 'MATRIZ_DE_CONTRATACION_VIGENCIA_2024_2025.xlsx'
data = pd.ExcelFile(file_path)

# Obtener las pestañas del archivo Excel
sheet_names = data.sheet_names

# Configurar la interfaz de usuario de Streamlit
st.title("Dashboard de Control de Contratación")

# Filtro global por pestaña
selected_sheet = st.sidebar.selectbox("Seleccione la pestaña de datos", sheet_names)

data_sheet = data.parse(selected_sheet, skiprows=2)

data_sheet.columns = [f"Col_{i}" for i in range(data_sheet.shape[1])]

# Filtros desplegables
ref_adtivo = st.sidebar.multiselect("Ref. Adtivo", options=data_sheet["Col_1"].dropna().unique())
grupo_trabajo = st.sidebar.multiselect("Grupo de trabajo", options=data_sheet["Col_2"].dropna().unique())

# Filtro de periodo y semáforo por "Fecha de Finalización"
data_sheet["Col_9"] = pd.to_datetime(data_sheet["Col_9"], errors='coerce')

# Crear semáforo
hoy = datetime.now()
data_sheet["Semaforo"] = data_sheet["Col_9"].apply(
    lambda x: "Rojo" if pd.notnull(x) and x <= hoy + timedelta(days=15) else
              "Amarillo" if pd.notnull(x) and x <= hoy + timedelta(days=30) else
              "Verde" if pd.notnull(x) else "Sin Fecha"
)

# Filtro de búsqueda
nombre_contratista = st.sidebar.text_input("Buscar por Nombre Contratista")
ref_adtivo_search = st.sidebar.text_input("Buscar por Ref. Adtivo")
objeto = st.sidebar.text_input("Buscar por Objeto")

# Control deslizante para Perfil
perfil_range = st.sidebar.slider("Filtrar Perfil", min_value=0, max_value=10000000, step=500000)

# Aplicar filtros
filtered_data = data_sheet.copy()
if ref_adtivo:
    filtered_data = filtered_data[filtered_data["Col_1"].isin(ref_adtivo)]
if grupo_trabajo:
    filtered_data = filtered_data[filtered_data["Col_2"].isin(ref_adtivo)]
if nombre_contratista:
    filtered_data = filtered_data[filtered_data["Col_4"].str.contains(nombre_contratista, na=False, case=False)]
if ref_adtivo_search:
    filtered_data = filtered_data[filtered_data["Col_1"].str.contains(ref_adtivo_search, na=False, case=False)]
if objeto:
    filtered_data = filtered_data[filtered_data["Col_7"].str.contains(objeto, na=False, case=False)]
filtered_data = filtered_data[filtered_data["Col_5"].fillna(0).astype(int) <= perfil_range]

# Mostrar datos filtrados
st.dataframe(filtered_data)

# Visualizaciones
st.subheader("Distribución del Semáforo")
st.bar_chart(filtered_data["Semaforo"].value_counts())
