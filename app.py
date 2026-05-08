"""
Sistema de análisis de inscripciones por área y semana.
Interfaz web creada con Streamlit.
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
from io import BytesIO

# Configuración de la página
st.set_page_config(
    page_title="Análisis de Inscripciones",
    page_icon="🎓",
    layout="wide"
)

# Título principal
st.title("🎓 Sistema de Análisis de Inscripciones")
st.markdown("---")

@st.cache_data
def cargar_y_validar_datos(file_fechas, file_alumnos):
    """Carga y valida los datos cruzando ambos archivos Excel."""
    df_fechas = pd.read_excel(file_fechas)
    df_alumnos = pd.read_excel(file_alumnos)
    
    # Limpiar y estandarizar valores de GASTOS_ADM, PAGO1 y PAGO2
    if 'GASTOS_ADM' in df_alumnos.columns:
        df_alumnos['GASTOS_ADM_clean'] = df_alumnos['GASTOS_ADM'].astype(str).str.lower().str.strip()
    else:
        df_alumnos['GASTOS_ADM_clean'] = 'no'
    df_alumnos['PAGO1_clean'] = df_alumnos['PAGO1'].astype(str).str.lower().str.strip()
    df_alumnos['PAGO2_clean'] = df_alumnos['PAGO2'].astype(str).str.lower().str.strip()
    
    # Filtrar Alumnos con GASTOS_ADM, PAGO1 o PAGO2 realizados ('si')
    alumnos_pagados = df_alumnos[
        (df_alumnos['GASTOS_ADM_clean'] == 'si') | (df_alumnos['PAGO1_clean'] == 'si') | (df_alumnos['PAGO2_clean'] == 'si')
    ]
    
    # Llave compuesta por MATRICULA y AREA para asegurar congruencia
    valid_keys = set(zip(alumnos_pagados['MATRICULA'], alumnos_pagados['AREA']))
    
    # Generar llave en fechas
    df_fechas['key'] = list(zip(df_fechas['MATRICULA'], df_fechas['AREA']))
    
    # Filtrar fechas para dejar solo válidos
    df_fechas_valido = df_fechas[df_fechas['key'].isin(valid_keys)].copy()
    df_fechas_eliminado = df_fechas[~df_fechas['key'].isin(valid_keys)].copy()
    
    # Contar registros quitados por AREA
    removed_counts = df_fechas_eliminado['AREA'].value_counts().reset_index()
    removed_counts.columns = ['AREA', 'Registros Eliminados']
    
    # Identificar registros de AlumnosNI que no están en Fechas (por MATRICULA)
    matriculas_fechas = set(df_fechas['MATRICULA'])
    alumnos_no_en_fechas = df_alumnos[~df_alumnos['MATRICULA'].isin(matriculas_fechas)].copy()
    
    # Limpiar columnas auxiliares
    df_fechas_valido = df_fechas_valido.drop(columns=['key'])
    
    return df_fechas_valido, removed_counts, alumnos_no_en_fechas

@st.cache_data
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Inscritos')
    return output.getvalue()

# Carga de archivos
st.sidebar.header("📁 Carga de Archivos")
archivo_fechas = st.sidebar.file_uploader("Archivo 'Fechas Inscripcion'", type=['xlsx', 'xls'])
archivo_alumnos = st.sidebar.file_uploader("Archivo 'Alumnos NI'", type=['xlsx', 'xls'])

if archivo_fechas is None or archivo_alumnos is None:
    st.info("👈 Por favor, sube ambos archivos de Excel en el menú lateral para comenzar el análisis.")
    st.stop()

# Cargar y validar datos
with st.spinner("Validando registros con archivo de Alumnos NI..."):
    df, removed_counts, alumnos_no_en_fechas = cargar_y_validar_datos(archivo_fechas, archivo_alumnos)

# Sección de Validación
with st.expander("📊 Resultados del Proceso de Validación de Registros", expanded=False):
    st.info("Se validó que los registros de 'Fechas Inscripcion y nacimiento' tengan GASTOS_ADM, PAGO1 o PAGO2 realizados en 'Alumnos NI status 1er ciclo' para la misma ÁREA.")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Registros Eliminados por ÁREA")
        if not removed_counts.empty:
            st.dataframe(removed_counts, hide_index=True)
        else:
            st.success("No se eliminó ningún registro.")
            
    with col2:
        st.subheader("No encontrados en Fechas Inscripcion")
        st.write(f"Hubo {len(alumnos_no_en_fechas)} registros en 'Alumnos NI' que no se encontraron en 'Fechas Inscripcion'.")
        if not alumnos_no_en_fechas.empty:
            # Mostrar solo columnas relevantes para no saturar
            cols_to_show = [col for col in ['MATRICULA', 'AREA', 'GASTOS_ADM', 'PAGO1', 'PAGO2', 'Status'] if col in alumnos_no_en_fechas.columns]
            if not cols_to_show:
                cols_to_show = alumnos_no_en_fechas.columns.tolist()
            st.dataframe(alumnos_no_en_fechas[cols_to_show], hide_index=True)

    st.subheader("Descargar Base de Datos de Inscritos Válidos")
    excel_data = convert_df_to_excel(df)
    st.download_button(
        label="📥 Descargar Inscritos.xlsx",
        data=excel_data,
        file_name='Inscritos_Validos.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    
st.markdown("---")

# Preparar datos
df['FECHA_INSC'] = pd.to_datetime(df['FECHA_INSC'], errors='coerce')
df['AÑO'] = df['FECHA_INSC'].dt.year
df['SEMANA'] = df['FECHA_INSC'].dt.isocalendar().week

# Obtener áreas ordenadas
areas = ["TODAS"] + sorted(df['AREA'].unique())

# Sidebar para selección
st.sidebar.header("📋 Configuración")
area_seleccionada = st.sidebar.selectbox(
    "Selecciona un área:",
    options=areas,
    index=0
)

# Filtrar datos por área
if area_seleccionada == "TODAS":
    df_area = df.copy()
else:
    df_area = df[df['AREA'] == area_seleccionada].copy()

# Información del área
st.header(f"📚 {area_seleccionada}")

if area_seleccionada != "TODAS":
    excel_area = convert_df_to_excel(df_area[['MATRICULA', 'FECHA_INSC']])
    st.download_button(
        label="📥 Descargar Excel de Matrículas y Fechas",
        data=excel_area,
        file_name=f'Matriculas_Fechas_{area_seleccionada}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        key="download_area_excel"
    )

col1, col2, col3 = st.columns(3)
with col1:
    st.metric("Total Registros", f"{len(df_area):,}")
with col2:
    st.metric("Años con datos", f"{df_area['AÑO'].nunique()}")
with col3:
    primer_año = df_area['AÑO'].min()
    ultimo_año = df_area['AÑO'].max()
    st.metric("Período", f"{primer_año} - {ultimo_año}")

st.markdown("---")

# Calcular inscripciones por semana
conteo = df_area.groupby(['AÑO', 'SEMANA']).size().reset_index(name='INSCRITOS')

# Agregar fecha de inicio de semana (lunes)
def get_week_start(row):
    año = int(row['AÑO'])
    semana = int(row['SEMANA'])
    # Encontrar el lunes de esa semana
    date = pd.Timestamp(f"{año}-01-01")
    # Ajustar al primer lunes del año
    first_monday = date - pd.Timedelta(days=date.dayofweek)
    if date.dayofweek != 0:  # Si no es lunes
        first_monday = date + pd.Timedelta(days=(7 - date.dayofweek))
    # Sumar semanas menos 1
    week_start = first_monday + pd.Timedelta(weeks=semana-1)
    return week_start.strftime('%Y-%m-%d')

conteo['FECHA_INICIO'] = conteo.apply(get_week_start, axis=1)

# Selector de año
años_disponibles = sorted(conteo['AÑO'].unique(), reverse=True)
año_seleccionado = st.selectbox(
    "Selecciona el año a analizar:",
    options=años_disponibles,
    index=0
)

# Selector de año de referencia
años_ref = [a for a in años_disponibles if a != año_seleccionado]
año_referencia = st.selectbox(
    "Selecciona el año de referencia:",
    options=años_ref,
    index=0
)

# Determinar hasta qué semana mostrar (52 o 53)
max_semana_posible = max(52, conteo['SEMANA'].max() if not conteo.empty else 52)

# Crear DataFrame con todas las semanas para el año seleccionado
df_todas_semanas = pd.DataFrame({'SEMANA': range(1, max_semana_posible + 1)})
df_todas_semanas['AÑO'] = año_seleccionado

# Extraer datos reales del año seleccionado
datos_año = conteo[conteo['AÑO'] == año_seleccionado][['SEMANA', 'INSCRITOS']]

# Unir (Left Join) para tener todas las semanas
conteo_año = pd.merge(df_todas_semanas, datos_año, on='SEMANA', how='left')
conteo_año['INSCRITOS'] = conteo_año['INSCRITOS'].fillna(0).astype(int)

# Calcular FECHA_INICIO para todas las semanas
conteo_año['FECHA_INICIO'] = conteo_año.apply(get_week_start, axis=1)

# Tabla de datos
st.subheader(f"📋 Detalle por Semana - {año_seleccionado}")

# Acumulado año seleccionado
conteo_año['ACUMULADO'] = conteo_año['INSCRITOS'].cumsum()

# Agregar columna del año de referencia
conteo_ref = conteo[conteo['AÑO'] == año_referencia].set_index('SEMANA')['INSCRITOS'].to_dict()
conteo_año['AÑO_REFERENCIA'] = conteo_año['SEMANA'].map(conteo_ref).fillna(0).astype(int)

# Calcular acumulado del año de referencia
serie_ref = pd.Series(conteo_ref, dtype=float).reindex(range(1, max_semana_posible + 1), fill_value=0)
acum_ref_dict = serie_ref.cumsum().to_dict()
conteo_año['ACUM_REFERENCIA'] = conteo_año['SEMANA'].map(acum_ref_dict).astype(int)

# Calcular variación porcentual del acumulado vs acumulado año de referencia
def calc_variacion_acum(row):
    col_acum_ref = f'Acum {año_referencia}'
    if row[col_acum_ref] == 0:
        return None
    return ((row['Acumulado'] - row[col_acum_ref]) / row[col_acum_ref]) * 100

# Renombrar columnas para mostrar y calcular variación acumulada
tabla_display = conteo_año[['AÑO', 'SEMANA', 'FECHA_INICIO', 'INSCRITOS', 'AÑO_REFERENCIA', 'ACUMULADO', 'ACUM_REFERENCIA']].copy()
tabla_display = tabla_display.rename(columns={
    'AÑO': 'Año',
    'SEMANA': 'Semana',
    'FECHA_INICIO': 'Fecha Inicio',
    'INSCRITOS': 'Inscripciones',
    'AÑO_REFERENCIA': f'Inscripciones {año_referencia}',
    'ACUMULADO': 'Acumulado',
    'ACUM_REFERENCIA': f'Acum {año_referencia}'
})
tabla_display['Var %'] = tabla_display.apply(calc_variacion_acum, axis=1)

# Aplicar estilo condicional para variación negativa
def highlight_variacion(val):
    if pd.isna(val):
        return 'color: gray'
    if val < 0:
        return 'color: red; font-weight: bold'
    return 'color: green'

st.dataframe(
    tabla_display.style.format({
        'Año': '{:.0f}',
        'Var %': lambda x: f'{x:.1f}%' if pd.notna(x) else '—'
    }).map(highlight_variacion, subset=['Var %']),
    use_container_width=True,
    hide_index=True
)

# Resumen por año
st.markdown("---")
st.subheader("📈 Resumen por Año")

resumen = conteo.groupby('AÑO')['INSCRITOS'].sum().reset_index()
resumen = resumen.sort_values('AÑO', ascending=False)

# Gráfico de línea
fig2 = px.line(
    resumen,
    x='AÑO',
    y='INSCRITOS',
    markers=True,
    title=f"Inscripciones por Año - {area_seleccionada}"
)

fig2.update_layout(
    xaxis_title="Año",
    yaxis_title="Total Inscripciones",
    height=350
)

st.plotly_chart(fig2, use_container_width=True)

# Tabla resumen
st.dataframe(
    resumen.rename(columns={'AÑO': 'Año', 'INSCRITOS': 'Total'}),
    use_container_width=True,
    hide_index=True
)

# Resumen General (Tabla Pivote por Área y Año)
st.markdown("---")
st.subheader("📊 Resumen General por Área y Año")

# Crear tabla pivote
pivot_areas = pd.pivot_table(
    df, 
    index='AREA', 
    columns='AÑO', 
    aggfunc='size', 
    fill_value=0
)

# Agregar totales
pivot_areas['Total Área'] = pivot_areas.sum(axis=1)
pivot_areas.loc['TOTAL GENERAL'] = pivot_areas.sum()

# Resetear índice para mostrar 'AREA' como columna normal
pivot_areas = pivot_areas.reset_index()

st.dataframe(pivot_areas, use_container_width=True, hide_index=True)

excel_pivot = convert_df_to_excel(pivot_areas)
st.download_button(
    label="📥 Descargar Resumen por Área.xlsx",
    data=excel_pivot,
    file_name='Resumen_Areas_Anio.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)

# Análisis Year-to-Date (Acumulado al Mes)
st.markdown("---")
st.subheader("📅 Análisis Comparativo hasta un Mes Específico (YTD)")
st.write("Selecciona un mes de corte. La tabla mostrará el total de inscritos desde el 1 de enero hasta el final del mes elegido para cada año. Esto permite comparar el año actual en curso (ej. 2026) con el mismo periodo de años anteriores.")

nombres_meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

# Mes actual por defecto (si hay error, usamos enero)
try:
    mes_actual = datetime.now().month
except:
    mes_actual = 1

mes_seleccionado = st.selectbox(
    "Selecciona el mes de corte:",
    options=range(1, 13),
    format_func=lambda x: nombres_meses[x-1],
    index=mes_actual - 1
)

# Filtrar df hasta el mes seleccionado
df_ytd = df[df['FECHA_INSC'].dt.month <= mes_seleccionado].copy()

# Crear tabla pivote YTD
pivot_ytd = pd.pivot_table(
    df_ytd, 
    index='AREA', 
    columns='AÑO', 
    aggfunc='size', 
    fill_value=0
)

# Agregar totales
pivot_ytd['Total Área (YTD)'] = pivot_ytd.sum(axis=1)
pivot_ytd.loc['TOTAL GENERAL'] = pivot_ytd.sum()

# Resetear índice para mostrar 'AREA' como columna normal
pivot_ytd = pivot_ytd.reset_index()

st.dataframe(pivot_ytd, use_container_width=True, hide_index=True)

excel_ytd = convert_df_to_excel(pivot_ytd)
st.download_button(
    label=f"📥 Descargar Resumen YTD (Enero a {nombres_meses[mes_seleccionado-1]}).xlsx",
    data=excel_ytd,
    file_name=f'Resumen_YTD_Enero_a_{nombres_meses[mes_seleccionado-1]}.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)

# Footer
st.markdown("---")
st.caption(f"📂 Total de registros válidos en el sistema: {len(df):,}")