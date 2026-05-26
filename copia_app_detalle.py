import streamlit as st
import pandas as pd
import numpy as np
import holidays
import io
import os
from datetime import date

st.set_page_config(page_title="Comparativa EVA vs Modelación", layout="wide", page_icon="📊")

@st.cache_data
def get_template_excel(tipo):
    if tipo == 'EVA':
        vars_list = ['Venta', 'Costo Alimento', 'Gasto Manipulación', 'Gasto Fijo', 'Gasto Variable', 'Margen de Contribución']
    else:
        vars_list = ['Venta', 'Costo', 'Manipulación', 'Fijo', 'Variable', 'Margen']
        
    df = pd.DataFrame({
        'CC': ['1001'] * 6,
        'Nombre Cliente': ['Cliente Ejemplo'] * 6,
        'Tipo Modelo': [tipo] * 6,
        'Variable': vars_list,
        '01-2024': [1000000, 300000, 150000, 100000, 50000, 400000],
        '02-2024': [1100000, 330000, 160000, 100000, 55000, 455000]
    })
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=f'Plantilla {tipo}')
    return output.getvalue()

# --- ESTILOS VISUALES (CSS) ---
st.markdown("""
<style>
    .main { background-color: #f8f9fa; }
    h1 { color: #1E3A8A; font-family: 'Inter', sans-serif; font-weight: 800; }
    .stButton>button { background-color: #2563EB; color: white; border-radius: 8px; font-weight: 600; border: none; }
    .stButton>button:hover { background-color: #1D4ED8; }
    .metric-card { background: white; padding: 1.5rem; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); text-align: center; }
    .metric-value { font-size: 2rem; font-weight: 700; color: #2563EB; }
    .metric-label { font-size: 1rem; color: #6B7280; text-transform: uppercase; letter-spacing: 0.05em; }
    .verdict-good { background-color: #d1fae5; border-left: 5px solid #10b981; padding: 1rem; border-radius: 4px; margin-bottom: 1rem; color: #065f46; }
    .verdict-bad { background-color: #fee2e2; border-left: 5px solid #ef4444; padding: 1rem; border-radius: 4px; margin-bottom: 1rem; color: #991b1b; }
    .verdict-neutral { background-color: #f3f4f6; border-left: 5px solid #6b7280; padding: 1rem; border-radius: 4px; margin-bottom: 1rem; color: #374151; }
    .diff-table { width: 100%; border-collapse: collapse; font-size: 0.85rem; font-family: 'Inter', sans-serif; }
    .diff-table th { background: #1E3A8A; color: white; padding: 8px 10px; text-align: center; font-weight: 600; white-space: nowrap; }
    .diff-table td { padding: 6px 10px; text-align: right; border-bottom: 1px solid #e5e7eb; white-space: nowrap; }
    .diff-table td:first-child { text-align: left; font-weight: 600; color: #374151; }
    .diff-table tr:hover { background-color: #f0f4ff; }
    .diff-positive { color: #059669; font-weight: 600; }
    .diff-negative { color: #DC2626; font-weight: 600; }
    .diff-zero { color: #6B7280; }
    .summary-table { width: 100%; border-collapse: collapse; font-size: 0.9rem; margin-top: 0.5rem; }
    .summary-table th { background: linear-gradient(135deg, #1E3A8A, #2563EB); color: #FFFFFF !important; padding: 12px 14px; text-align: center; font-weight: 600; }
    .summary-table td { padding: 10px 14px; border-bottom: 1px solid #e5e7eb; text-align: right; color: #1F2937 !important; background-color: #FFFFFF !important; }
    .summary-table td:first-child { text-align: left; font-weight: 600; color: #1E3A8A !important; }
    .summary-table tr { background-color: #FFFFFF !important; }
    .summary-table tr:nth-child(even) { background-color: #F9FAFB !important; }
    .summary-table tr:hover { background-color: #EFF6FF !important; }
</style>
""", unsafe_allow_html=True)

st.title("📊 Sistema de Comparativa: EVA vs Modelación")
st.write("Análisis Anualizado con Proyección por Días Hábiles e Inercia de Gastos")

def get_dias_habiles(year, month):
    start_date = date(year, month, 1)
    if month == 12:
        end_date = date(year + 1, 1, 1)
    else:
        end_date = date(year, month + 1, 1)
        
    cl_holidays = holidays.Chile(years=year)
    feriados = list(cl_holidays.keys())
    
    # np.busday_count expects YYYY-MM-DD string or datetime.date array, feriados must be same
    return int(np.busday_count(start_date, end_date, holidays=feriados))

def safe_div(n, d):
    return n / d if d and d != 0 else 0

def is_venta(var_name):
    return 'venta' in str(var_name).lower()

def is_margen(var_name):
    return 'margen' in str(var_name).lower()

def process_data(df):
    required_cols = ['CC', 'Nombre Cliente', 'Tipo Modelo', 'Variable']
    for col in required_cols:
        if col not in df.columns:
            return None, None, f"Falta la columna requerida: {col}"

    month_cols = [col for col in df.columns if col not in required_cols]
    
    # Limpieza de datos clave para evitar problemas de cruce
    df['CC'] = df['CC'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df['Nombre Cliente'] = df['Nombre Cliente'].astype(str).str.strip().str.upper()
    df['Tipo Modelo'] = df['Tipo Modelo'].astype(str).str.strip().str.upper()
    df['Tipo Modelo'] = df['Tipo Modelo'].replace('MODELACIÓN', 'MODELACION')
    df['Variable'] = df['Variable'].astype(str).str.strip()
    
    # Mapear nombres largos (EVA) a nombres cortos (Modelación) para unificarlos
    mapping_vars = {
        'Costo Alimento': 'Costo',
        'Gasto Manipulación': 'Manipulación',
        'Gasto Fijo': 'Fijo',
        'Gasto Variable': 'Variable',
        'Margen de Contribución': 'Margen'
    }
    df['Variable'] = df['Variable'].replace(mapping_vars)
    
    df_melt = pd.melt(df, id_vars=required_cols, value_vars=month_cols, var_name='Mes', value_name='Valor')
    df_melt['Valor'] = pd.to_numeric(df_melt['Valor'], errors='coerce').fillna(0)
    
    try:
        df_melt['Fecha'] = pd.to_datetime(df_melt['Mes'], format='%m-%Y', errors='coerce')
        mask = df_melt['Fecha'].isna()
        if mask.any():
            df_melt.loc[mask, 'Fecha'] = pd.to_datetime(df_melt.loc[mask, 'Mes'], errors='coerce')
    except:
        df_melt['Fecha'] = pd.to_datetime(df_melt['Mes'], errors='coerce')

    if df_melt['Fecha'].isna().all():
         return None, None, "No se pudieron interpretar las columnas de meses como fechas. Usa formato MM-YYYY."

    df_eva = df_melt[df_melt['Tipo Modelo'] == 'EVA'].copy()
    df_mod = df_melt[df_melt['Tipo Modelo'].isin(['MODELACION', 'INTERNO'])].copy()

    resultados_mensuales = []
    resumen_ejecutivo = []

    ccs = df_eva['CC'].drop_duplicates()
    
    for cc in ccs:
        eva_cli = df_eva[df_eva['CC'] == cc]
        mod_cli = df_mod[df_mod['CC'] == cc]
        
        # Obtener el nombre del cliente desde la modelación de preferencia, si no, del EVA
        if not mod_cli.empty:
            nombre = mod_cli['Nombre Cliente'].iloc[0]
        else:
            nombre = eva_cli['Nombre Cliente'].iloc[0]
        
        eva_ventas = eva_cli[eva_cli['Variable'].str.contains('Venta', case=False, na=False)]
        if eva_ventas.empty: 
            eva_ventas = eva_cli 
        
        meses_validos_eva = eva_ventas[eva_ventas['Valor'] != 0]['Fecha'].sort_values().drop_duplicates()
        if meses_validos_eva.empty:
            continue
            
        fecha_inicio = meses_validos_eva.min()
        fechas_ciclo_eva = pd.date_range(start=fecha_inicio, periods=12, freq='MS')
        meses_str = [f.strftime('%m-%Y') for f in fechas_ciclo_eva]
        
        eva_pivot = eva_cli.pivot_table(index='Variable', columns='Fecha', values='Valor', aggfunc='sum').fillna(0)
        mod_pivot_raw = mod_cli.pivot_table(index='Variable', columns='Fecha', values='Valor', aggfunc='sum').fillna(0)
        
        variables = set(eva_pivot.index).union(set(mod_pivot_raw.index))
        var_venta = next((v for v in variables if is_venta(v)), None)
        var_margen = next((v for v in variables if is_margen(v)), None)
        
        eva_12m = pd.DataFrame(index=list(variables), columns=fechas_ciclo_eva).fillna(0)
        mod_12m = pd.DataFrame(index=list(variables), columns=fechas_ciclo_eva).fillna(0)
        
        # Último mes real disponible en la modelación (ej. Diciembre)
        # Se usará para inercia de gastos
        fechas_disponibles_mod = [col for col in mod_pivot_raw.columns if any(mod_pivot_raw[col] != 0)]
        ultimo_mes_mod = max(fechas_disponibles_mod) if fechas_disponibles_mod else None

        for d in fechas_ciclo_eva:
            # 1. Cargar EVA
            for var in variables:
                if var in eva_pivot.index and d in eva_pivot.columns:
                    eva_12m.at[var, d] = eva_pivot.at[var, d]
            
            # 2. Cargar Modelación Ajustada
            if d in mod_pivot_raw.columns and any(mod_pivot_raw[d] != 0):
                # El mes existe en la modelación real, se copia exacto
                for var in variables:
                    if var in mod_pivot_raw.index:
                        mod_12m.at[var, d] = mod_pivot_raw.at[var, d]
            else:
                # PROYECCIÓN PARA MESES FALTANTES
                # Regla A: Venta (Días hábiles)
                if var_venta:
                    mes_buscado = d.month
                    # Buscar el mismo mes en el año base que tenga datos reales (distinto de 0)
                    columnas_mismo_mes = [c for c in mod_pivot_raw.columns if c.month == mes_buscado and var_venta in mod_pivot_raw.index and mod_pivot_raw.at[var_venta, c] != 0]
                    if columnas_mismo_mes:
                        mes_base = max(columnas_mismo_mes)
                        venta_base = mod_pivot_raw.at[var_venta, mes_base]
                        
                        dias_base = get_dias_habiles(mes_base.year, mes_base.month)
                        dias_proy = get_dias_habiles(d.year, d.month)
                        
                        venta_diaria = safe_div(venta_base, dias_base)
                        mod_12m.at[var_venta, d] = venta_diaria * dias_proy
                    else:
                        mod_12m.at[var_venta, d] = 0
                
                # Regla B: Costos y Gastos
                # - Costo Alimento / Gasto Variable → % sobre venta Dic aplicado a venta proyectada
                # - Gasto Fijo / Manipulación → valor absoluto de Dic (inercia)
                costos_gastos = 0
                venta_proyectada = mod_12m.at[var_venta, d] if var_venta else 0

                # Venta del último mes real de modelación (base para calcular los %)
                venta_ultimo_mes = 0
                if ultimo_mes_mod and var_venta and var_venta in mod_pivot_raw.index:
                    venta_ultimo_mes = mod_pivot_raw.at[var_venta, ultimo_mes_mod]

                for var in variables:
                    if var != var_venta and var != var_margen:
                        val_proyectado = 0
                        if ultimo_mes_mod and var in mod_pivot_raw.index:
                            val_ultimo = mod_pivot_raw.at[var, ultimo_mes_mod]
                            var_lower = str(var).lower()
                            # Ítems proporcionales a la venta: Costo Alimento, Gasto Variable
                            if 'costo' in var_lower or 'variable' in var_lower:
                                pct_sobre_venta = safe_div(val_ultimo, venta_ultimo_mes)
                                val_proyectado = pct_sobre_venta * venta_proyectada
                            else:
                                # Gasto Fijo, Manipulación → mantener valor absoluto de Dic
                                val_proyectado = val_ultimo
                        mod_12m.at[var, d] = val_proyectado
                        costos_gastos += val_proyectado

                # Regla Margen: Venta Proyectada - Costos Proyectados
                if var_margen:
                    mod_12m.at[var_margen, d] = venta_proyectada - costos_gastos

        # --- CÁLCULO DE PORCENTAJES SOBRE VENTA Y DIFERENCIALES MENSUALES ---
        for tipo, df_datos in [('EVA', eva_12m), ('Modelación Ajustada', mod_12m)]:
            # Agregar Días Hábiles
            fila_dias = {'CC': cc, 'Nombre Cliente': nombre, 'Tipo Modelo': tipo, 'Variable': 'Días Hábiles'}
            for i, d in enumerate(fechas_ciclo_eva):
                fila_dias[meses_str[i]] = get_dias_habiles(d.year, d.month)
            resultados_mensuales.append(fila_dias)
            
            for var in variables:
                fila = {'CC': cc, 'Nombre Cliente': nombre, 'Tipo Modelo': tipo, 'Variable': var}
                for i, d in enumerate(fechas_ciclo_eva):
                    fila[meses_str[i]] = df_datos.at[var, d]
                resultados_mensuales.append(fila)
                
                # Porcentaje sobre venta
                if var_venta and var != var_venta:
                    fila_pct = {'CC': cc, 'Nombre Cliente': nombre, 'Tipo Modelo': tipo, 'Variable': f"{var} %"}
                    for i, d in enumerate(fechas_ciclo_eva):
                        venta_mes = df_datos.at[var_venta, d]
                        valor_mes = df_datos.at[var, d]
                        fila_pct[meses_str[i]] = safe_div(valor_mes, venta_mes)
                    resultados_mensuales.append(fila_pct)
                    
        # Diferencial (EVA - Modelación)
        dif_12m = eva_12m - mod_12m
        for var in variables:
            fila_dif = {'CC': cc, 'Nombre Cliente': nombre, 'Tipo Modelo': 'Diferencial (EVA - Mod)', 'Variable': var}
            for i, d in enumerate(fechas_ciclo_eva):
                fila_dif[meses_str[i]] = dif_12m.at[var, d]
            resultados_mensuales.append(fila_dif)

        # --- RESUMEN ANUALIZADO Y VEREDICTO ---
        resumen_cli = {'CC': cc, 'Nombre Cliente': nombre}
        
        suma_ventas_eva = eva_12m.loc[var_venta].sum() if var_venta else 0
        suma_ventas_mod = mod_12m.loc[var_venta].sum() if var_venta else 0
        
        suma_mc_eva = eva_12m.loc[var_margen].sum() if var_margen else 0
        suma_mc_mod = mod_12m.loc[var_margen].sum() if var_margen else 0
        
        mc_pct_eva = safe_div(suma_mc_eva, suma_ventas_eva)
        mc_pct_mod = safe_div(suma_mc_mod, suma_ventas_mod)
        
        dif_abs_mc = suma_mc_eva - suma_mc_mod
        dif_pct_mc = mc_pct_eva - mc_pct_mod
        
        # Calcular Días Hábiles Totales
        total_dias_habiles = sum([get_dias_habiles(d.year, d.month) for d in fechas_ciclo_eva])
        
        # Evaluar Veredicto
        veredicto = "Neutral"
        if dif_abs_mc > 0 and dif_pct_mc > 0:
            veredicto = "Mejor (El EVA supera a la Modelación en monto Final y %)"
        elif dif_abs_mc < 0 and dif_pct_mc < 0:
            veredicto = "Peor (El EVA es peor respecto a la Modelación)"
        elif dif_abs_mc > 0 and dif_pct_mc < 0:
            veredicto = "Mixto (El EVA genera más dinero pero con menor margen porcentual)"
        elif dif_abs_mc < 0 and dif_pct_mc > 0:
            veredicto = "Mixto (El EVA genera menos dinero pero con mayor margen porcentual)"
            
        resumen_cli['Veredicto MC'] = veredicto
        resumen_cli['Total Días Hábiles (12m)'] = total_dias_habiles
        resumen_cli['Variables'] = list(variables)
        
        # Totales de TODAS las variables (no solo Venta y MC)
        for var in variables:
            total_eva = eva_12m.loc[var].sum()
            total_mod = mod_12m.loc[var].sum()
            resumen_cli[f'Total {var} EVA'] = total_eva
            resumen_cli[f'Total {var} Modelación'] = total_mod
            resumen_cli[f'Dif {var}'] = total_eva - total_mod
            # Porcentaje sobre venta para cada variable
            resumen_cli[f'% {var} EVA'] = safe_div(total_eva, suma_ventas_eva)
            resumen_cli[f'% {var} Modelación'] = safe_div(total_mod, suma_ventas_mod)
        
        # Agregar promedios de todas las variables
        for var in variables:
            resumen_cli[f'Promedio Mensual {var} (EVA)'] = eva_12m.loc[var].mean()
            resumen_cli[f'Promedio Mensual {var} (Mod)'] = mod_12m.loc[var].mean()
            
        resumen_ejecutivo.append(resumen_cli)

    if not resultados_mensuales:
        return None, None, "No se encontraron datos procesables."

    df_mensual = pd.DataFrame(resultados_mensuales)
    df_resumen = pd.DataFrame(resumen_ejecutivo)
    
    return df_mensual, df_resumen, None


# --- CONFIGURACIÓN DE REPOSITORIO ---
REPO_DIR = "repo_modelaciones"
if not os.path.exists(REPO_DIR):
    os.makedirs(REPO_DIR)

# --- BARRA LATERAL (REPOSITORIO) ---
with st.sidebar:
    st.header("🗄️ Repositorio de Modelaciones")
    st.write("Sube aquí los archivos base de Modelación de todos los clientes.")
    
    st.download_button(
        label="📥 Descargar Plantilla Modelación",
        data=get_template_excel("MODELACION"),
        file_name="Plantilla_Modelacion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    nuevo_archivo = st.file_uploader("Subir nueva Modelación", type=["xlsx", "xls"], key="upload_mod")
    if nuevo_archivo is not None:
        file_path = os.path.join(REPO_DIR, nuevo_archivo.name)
        with open(file_path, "wb") as f:
            f.write(nuevo_archivo.getbuffer())
        st.success(f"Archivo guardado con éxito.")
    
    st.markdown("---")
    st.subheader("📁 Archivos Disponibles")
    archivos_repo = [f for f in os.listdir(REPO_DIR) if f.endswith('.xlsx') or f.endswith('.xls')]
    
    if archivos_repo:
        for f in archivos_repo:
            col_name, col_del = st.columns([4, 1])
            with col_name:
                st.write(f"📄 {f}")
            with col_del:
                if st.button("❌", key=f"del_{f}"):
                    os.remove(os.path.join(REPO_DIR, f))
                    st.rerun()
    else:
        st.info("No hay archivos guardados.")

# --- INTERFAZ PRINCIPAL ---
tab1, tab2 = st.tabs(["📊 Comparativa EVA vs Modelación", "🔍 Visualizador y Proyección de Modelación"])

with tab1:
    st.header("Comparativa EVA vs Modelación")
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("1. Seleccionar Modelación")
        if archivos_repo:
            mod_seleccionada = st.selectbox("Elige la modelación base:", options=["Seleccionar..."] + archivos_repo, key="mod_sel_tab1")
        else:
            st.warning("Sube un archivo en la barra lateral.")
            mod_seleccionada = "Seleccionar..."

    with col2:
        st.subheader("2. Subir Archivo EVA")
        
        st.download_button(
            label="📥 Descargar Plantilla EVA",
            data=get_template_excel("EVA"),
            file_name="Plantilla_EVA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_dl_eva_tab1"
        )
        
        uploaded_eva = st.file_uploader("Sube el archivo Excel de EVA", type=["xlsx", "xls"], key="upload_eva_tab1")

    st.markdown("---")

    if mod_seleccionada != "Seleccionar..." and uploaded_eva is not None:
        try:
            with st.spinner("Procesando archivos e infiriendo estacionalidad..."):
                path_mod = os.path.join(REPO_DIR, mod_seleccionada)
                df_mod_raw = pd.read_excel(path_mod)
                df_eva_raw = pd.read_excel(uploaded_eva)
                df = pd.concat([df_mod_raw, df_eva_raw], ignore_index=True)
                df_mensual, df_resumen, error = process_data(df)
            if error:
                st.error(error)
            else:
                st.markdown("### Resumen Anualizado por Cliente")
                
                # Orden deseado de variables (definido una sola vez)
                desired_order = ['Venta', 'Costo', 'Manipulación', 'Fijo', 'Variable', 'Margen', 'Días Hábiles']

                for _, row in df_resumen.iterrows():
                    st.markdown(f"#### Cliente: {row['Nombre Cliente']} (CC: {row['CC']})")
                    ver = str(row['Veredicto MC'])
                    if "Mejor" in ver:
                        st.markdown(f'<div class="verdict-good">✅ <b>Veredicto:</b> {ver}</div>', unsafe_allow_html=True)
                    elif "Peor" in ver:
                        st.markdown(f'<div class="verdict-bad">❌ <b>Veredicto:</b> {ver}</div>', unsafe_allow_html=True)
                    else:
                        st.markdown(f'<div class="verdict-neutral">⚠️ <b>Veredicto:</b> {ver}</div>', unsafe_allow_html=True)
                    
                    st.caption(f"🗓️ **Período Analizado:** 12 Meses | **Días Hábiles Totales:** {row['Total Días Hábiles (12m)']} días")
                    
                    # --- TABLA RESUMEN DE TOTALES POR VARIABLE ---
                    st.markdown("##### 💰 Resumen de Totales Anualizados")
                    variables_cli = row.get('Variables', [])
                    variables_cli = [v for v in desired_order if v in variables_cli]
                    if variables_cli:
                        summary_html = '<table class="summary-table"><thead><tr>'
                        summary_html += '<th>Variable</th><th>Total EVA</th><th>% s/Venta EVA</th><th>Total Modelación</th><th>% s/Venta Mod</th><th>Diferencia ($)</th><th>Resultado</th>'
                        summary_html += '</tr></thead><tbody>'
                        
                        for var in variables_cli:
                            total_eva = row.get(f'Total {var} EVA', 0)
                            total_mod = row.get(f'Total {var} Modelación', 0)
                            dif = row.get(f'Dif {var}', 0)
                            pct_eva = row.get(f'% {var} EVA', 0)
                            pct_mod = row.get(f'% {var} Modelación', 0)
                            
                            # Determinar color de la diferencia
                            var_lower = str(var).lower()
                            es_costo_gasto = any(k in var_lower for k in ['costo', 'gasto', 'manipulación', 'manipulacion', 'fijo', 'variable'])
                            
                            if dif > 0:
                                css_class = 'diff-negative' if es_costo_gasto else 'diff-positive'
                                icono = '🔺' if es_costo_gasto else '✅'
                            elif dif < 0:
                                css_class = 'diff-positive' if es_costo_gasto else 'diff-negative'
                                icono = '✅' if es_costo_gasto else '🔻'
                            else:
                                css_class = 'diff-zero'
                                icono = '➖'
                            
                            summary_html += f'<tr>'
                            summary_html += f'<td>{var}</td>'
                            summary_html += f'<td>${total_eva:,.0f}</td>'
                            summary_html += f'<td>{pct_eva:.1%}</td>'
                            summary_html += f'<td>${total_mod:,.0f}</td>'
                            summary_html += f'<td>{pct_mod:.1%}</td>'
                            summary_html += f'<td class="{css_class}">${dif:,.0f}</td>'
                            summary_html += f'<td style="text-align:center">{icono}</td>'
                            summary_html += f'</tr>'
                        
                        summary_html += '</tbody></table>'
                        st.markdown(summary_html, unsafe_allow_html=True)
                    
                    st.write("")
                    st.write("---")
                    st.markdown("##### 📊 Matrices Detalladas")
                    df_cli = df_mensual[(df_mensual['CC'] == row['CC']) & (df_mensual['Nombre Cliente'] == row['Nombre Cliente'])]
                    
                    col_mod, col_eva = st.columns(2)
                    with col_mod:
                        st.markdown("**🔧 Modelación Ajustada**")
                        df_mod_view = df_cli[(df_cli['Tipo Modelo'] == 'Modelación Ajustada') & (~df_cli['Variable'].str.endswith(' %'))].copy()
                        df_mod_disp = df_mod_view.drop(columns=['CC', 'Nombre Cliente', 'Tipo Modelo'])
                        df_mod_disp['Variable'] = pd.Categorical(df_mod_disp['Variable'], categories=desired_order, ordered=True)
                        df_mod_disp = df_mod_disp.sort_values('Variable').reset_index(drop=True)
                        mes_cols_mod = [c for c in df_mod_disp.columns if c != 'Variable']
                        df_mod_disp['TOTAL'] = df_mod_disp[mes_cols_mod].apply(pd.to_numeric, errors='coerce').sum(axis=1)
                        st.dataframe(df_mod_disp.set_index('Variable'), use_container_width=True)
                    with col_eva:
                        st.markdown("**📋 EVA**")
                        df_eva_view = df_cli[(df_cli['Tipo Modelo'] == 'EVA') & (~df_cli['Variable'].str.endswith(' %'))].copy()
                        df_eva_disp = df_eva_view.drop(columns=['CC', 'Nombre Cliente', 'Tipo Modelo'])
                        df_eva_disp['Variable'] = pd.Categorical(df_eva_disp['Variable'], categories=desired_order, ordered=True)
                        df_eva_disp = df_eva_disp.sort_values('Variable').reset_index(drop=True)
                        mes_cols_eva = [c for c in df_eva_disp.columns if c != 'Variable']
                        df_eva_disp['TOTAL'] = df_eva_disp[mes_cols_eva].apply(pd.to_numeric, errors='coerce').sum(axis=1)
                        st.dataframe(df_eva_disp.set_index('Variable'), use_container_width=True)
                    
                    # --- Tabla de Diferencias: ancho completo, con estilos ---
                    st.markdown("**📐 Diferencias (EVA - Modelación)**")
                    df_dif_view = df_cli[(df_cli['Tipo Modelo'] == 'Diferencial (EVA - Mod)') & (~df_cli['Variable'].str.endswith(' %'))].copy()
                    df_dif_disp = df_dif_view.drop(columns=['CC', 'Nombre Cliente', 'Tipo Modelo'])
                    df_dif_disp['Variable'] = pd.Categorical(df_dif_disp['Variable'], categories=desired_order, ordered=True)
                    df_dif_disp = df_dif_disp.sort_values('Variable').reset_index(drop=True)
                    mes_cols_dif = [c for c in df_dif_disp.columns if c != 'Variable']
                    df_dif_disp['TOTAL'] = df_dif_disp[mes_cols_dif].apply(pd.to_numeric, errors='coerce').sum(axis=1)
                    
                    # Generar tabla HTML con colores
                    diff_html = '<table class="diff-table"><thead><tr><th>Variable</th>'
                    for col in mes_cols_dif:
                        diff_html += f'<th>{col}</th>'
                    diff_html += '<th style="background:#0F2557;">TOTAL</th></tr></thead><tbody>'
                    
                    for _, dif_row in df_dif_disp.iterrows():
                        diff_html += '<tr>'
                        diff_html += f'<td>{dif_row["Variable"]}</td>'
                        var_lower = str(dif_row['Variable']).lower()
                        es_costo = any(k in var_lower for k in ['costo', 'gasto', 'manipulación', 'manipulacion', 'fijo', 'variable'])
                        
                        for col in mes_cols_dif + ['TOTAL']:
                            val = dif_row[col]
                            try:
                                val_num = float(val)
                            except:
                                val_num = 0
                            
                            if val_num > 0:
                                css = 'diff-positive' if not es_costo else 'diff-negative'
                            elif val_num < 0:
                                css = 'diff-negative' if not es_costo else 'diff-positive'
                            else:
                                css = 'diff-zero'
                            
                            style_extra = ' font-weight:800;' if col == 'TOTAL' else ''
                            diff_html += f'<td class="{css}" style="{style_extra}">${val_num:,.0f}</td>'
                        diff_html += '</tr>'
                    
                    diff_html += '</tbody></table>'
                    st.markdown(diff_html, unsafe_allow_html=True)
                    
                    st.write("")
                    st.write("---")
                    st.markdown("##### 📈 Análisis Porcentual sobre la Venta")
                    df_pct_view = df_cli[df_cli['Variable'].str.endswith(' %')].copy()
                    desired_order_pct = [f"{v} %" for v in desired_order if v not in ['Venta', 'Días Hábiles']]
                    df_pct_view['Variable'] = pd.Categorical(df_pct_view['Variable'], categories=desired_order_pct, ordered=True)
                    df_pct_view = df_pct_view.sort_values(['Variable', 'Tipo Modelo']).dropna(subset=['Variable']).reset_index(drop=True)
                    st.dataframe(df_pct_view.drop(columns=['CC', 'Nombre Cliente']).set_index(['Variable', 'Tipo Modelo']), use_container_width=True)
                    
                st.markdown("---")
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_resumen.to_excel(writer, index=False, sheet_name='Resumen Ejecutivo')
                    df_mensual.to_excel(writer, index=False, sheet_name='Datos Mensuales')
                
                st.write(" ")
                st.download_button(
                    label="📥 Descargar Excel con Resumen y Mes a Mes",
                    data=output.getvalue(),
                    file_name="Comparativa_Avanzada.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_dl_full_excel_tab1"
                )
                
        except Exception as e:
            st.error(f"Error al procesar la comparativa: {e}")

with tab2:
    st.header("Visualizador y Proyección de Modelación (12 Meses)")
    st.write("Módulo para analizar una modelación base y simular proyecciones personalizadas definiendo un mes límite.")
    
    st.subheader("1. Seleccionar Modelación")
    if archivos_repo:
        mod_seleccionada_v2 = st.selectbox("Elige la modelación base:", options=["Seleccionar..."] + archivos_repo, key="mod_sel_tab2")
    else:
        st.warning("Sube un archivo en la barra lateral.")
        mod_seleccionada_v2 = "Seleccionar..."

    if mod_seleccionada_v2 != "Seleccionar...":
        try:
            path_mod = os.path.join(REPO_DIR, mod_seleccionada_v2)
            df_mod_raw_all = pd.read_excel(path_mod)
            
            # Limpieza rápida
            df_mod_raw_all['CC'] = df_mod_raw_all['CC'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            df_mod_raw_all['Nombre Cliente'] = df_mod_raw_all['Nombre Cliente'].astype(str).str.strip().str.upper()
            df_mod_raw_all['Variable'] = df_mod_raw_all['Variable'].astype(str).str.strip()
            
            # --- FILTRO POR CC Y CASINO (BUSCADOR INTEGRADO) ---
            st.markdown("### 🎯 Seleccionar Centro de Costo (CC) o Casino")
            
            df_unique_clients = df_mod_raw_all[['CC', 'Nombre Cliente']].drop_duplicates().sort_values('CC')
            opciones_busqueda = ["Seleccionar..."] + [f"{row['CC']} - {row['Nombre Cliente']}" for _, row in df_unique_clients.iterrows()]
            
            seleccion_busqueda = st.selectbox(
                "Escribe o selecciona el CC o Nombre del Casino (Cliente):",
                options=opciones_busqueda,
                index=0,
                key="cc_sel_tab2"
            )
            
            if seleccion_busqueda == "Seleccionar...":
                st.info("Por favor, selecciona un Centro de Costo (CC) o Casino arriba para proyectar e iniciar el análisis.")
                st.stop()
                
            cc_seleccionado = seleccion_busqueda.split(" - ")[0]
            
            # Filtrar modelación para el CC seleccionado
            df_mod_raw = df_mod_raw_all[df_mod_raw_all['CC'] == cc_seleccionado].copy()
            
            mapping_vars = {
                'Costo Alimento': 'Costo',
                'Gasto Manipulación': 'Manipulación',
                'Gasto Fijo': 'Fijo',
                'Gasto Variable': 'Variable',
                'Margen de Contribución': 'Margen'
            }
            df_mod_raw['Variable'] = df_mod_raw['Variable'].replace(mapping_vars)
            
            # Obtener datos generales de cliente
            nombre_cliente = df_mod_raw['Nombre Cliente'].iloc[0] if not df_mod_raw.empty else "CLIENTE DESCONOCIDO"
            cc_cliente = cc_seleccionado
            
            # Columnas de meses
            excluded_cols = ['CC', 'Nombre Cliente', 'Tipo Modelo', 'Variable']
            month_cols = [col for col in df_mod_raw.columns if col not in excluded_cols]
            
            # Melt para obtener las fechas ordenadas
            df_melt = pd.melt(df_mod_raw, id_vars=['CC', 'Nombre Cliente', 'Variable'], value_vars=month_cols, var_name='Mes', value_name='Valor')
            df_melt['Valor'] = pd.to_numeric(df_melt['Valor'], errors='coerce').fillna(0)
            
            try:
                df_melt['Fecha'] = pd.to_datetime(df_melt['Mes'], format='%m-%Y', errors='coerce')
                mask = df_melt['Fecha'].isna()
                if mask.any():
                    df_melt.loc[mask, 'Fecha'] = pd.to_datetime(df_melt.loc[mask, 'Mes'], errors='coerce')
            except:
                df_melt['Fecha'] = pd.to_datetime(df_melt['Mes'], errors='coerce')
                
            fechas_ordenadas = sorted(df_melt['Fecha'].dropna().unique())
            year_base = fechas_ordenadas[0].year if fechas_ordenadas else 2026
            
            # --- CONTROLES DE PROYECCIÓN ---
            st.markdown("### ⚙️ Configuración de la Proyección")
            
            # Mapeo de meses en español
            MESES_ES = {
                1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio',
                7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
            }
            
            col_c1, col_c2 = st.columns(2)
            
            with col_c1:
                # Selector de mes límite
                mes_limite_nombre = st.selectbox(
                    "Proyectar hasta el mes de:",
                    options=list(MESES_ES.values()),
                    index=4 # Mayo por defecto
                )
                mes_limite_num = next(k for k, v in MESES_ES.items() if v == mes_limite_nombre)
                
            # Generar ciclo de 12 meses continuo:
            # Reales: Junio a Diciembre del año base (Y)
            # Proyectados: Enero a Mayo del año futuro (Y + 1)
            fechas_reales = [pd.to_datetime(f"01-{m:02d}-{year_base}", format="%d-%m-%Y") for m in range(mes_limite_num + 1, 13)]
            fechas_proyectadas = [pd.to_datetime(f"01-{m:02d}-{year_base + 1}", format="%d-%m-%Y") for m in range(1, mes_limite_num + 1)]
            
            fechas_ciclo = fechas_reales + fechas_proyectadas
            meses_str = [f.strftime('%m-%Y') for f in fechas_ciclo]
            
            with col_c2:
                st.write("**Resumen de Períodos:**")
                st.write(f"🟢 **Meses Reales ({year_base}):** {', '.join([MESES_ES[f.month] for f in fechas_reales]) if fechas_reales else 'Ninguno'}")
                st.write(f"🔵 **Meses Proyectados ({year_base + 1}):** {', '.join([MESES_ES[f.month] for f in fechas_proyectadas]) if fechas_proyectadas else 'Ninguno'}")
                
            # --- CÁLCULO AUTOMÁTICO DE AUMENTO DE VENTAS (USANDO EL EVA COMPARTIDO DE LA PESTAÑA 1 SI EXISTE) ---
            aumento_auto = 0.0
            usar_eva = False
            
            # Verificamos si hay archivo EVA subido en la Pestaña 1
            if 'upload_eva_tab1' in st.session_state and st.session_state['upload_eva_tab1'] is not None:
                try:
                    df_eva_raw_v2 = pd.read_excel(st.session_state['upload_eva_tab1'])
                    df_eva_raw_v2['CC'] = df_eva_raw_v2['CC'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                    df_eva_raw_v2['Variable'] = df_eva_raw_v2['Variable'].astype(str).str.strip()
                    df_eva_raw_v2['Variable'] = df_eva_raw_v2['Variable'].replace(mapping_vars)
                    
                    df_eva_raw_cc = df_eva_raw_v2[df_eva_raw_v2['CC'] == cc_seleccionado].copy()
                    
                    if not df_eva_raw_cc.empty:
                        df_eva_melt = pd.melt(df_eva_raw_cc, id_vars=['CC', 'Nombre Cliente', 'Variable'], value_vars=month_cols, var_name='Mes', value_name='Valor')
                        df_eva_melt['Valor'] = pd.to_numeric(df_eva_melt['Valor'], errors='coerce').fillna(0)
                        try:
                            df_eva_melt['Fecha'] = pd.to_datetime(df_eva_melt['Mes'], format='%m-%Y', errors='coerce')
                        except:
                            df_eva_melt['Fecha'] = pd.to_datetime(df_eva_melt['Mes'], errors='coerce')
                        
                        # Calcular la diferencia porcentual acumulada de ventas en los meses reales
                        ventas_mod_real_sum = 0.0
                        ventas_eva_real_sum = 0.0
                        
                        for f in fechas_reales:
                            v_mod = df_melt[(df_melt['Variable'] == 'Venta') & (df_melt['Fecha'] == f)]['Valor'].sum()
                            v_eva = df_eva_melt[(df_eva_melt['Variable'] == 'Venta') & (df_eva_melt['Fecha'] == f)]['Valor'].sum()
                            ventas_mod_real_sum += v_mod
                            ventas_eva_real_sum += v_eva
                            
                        if ventas_mod_real_sum > 0:
                            aumento_auto = (ventas_eva_real_sum - ventas_mod_real_sum) / ventas_mod_real_sum
                            usar_eva = True
                except Exception as ex:
                    pass
            
            st.markdown("---")
            
            # --- PROCESAMIENTO MATEMÁTICO DE LA PROYECCIÓN ---
            mod_pivot = df_melt.pivot_table(index='Variable', columns='Fecha', values='Valor', aggfunc='sum').fillna(0)
            
            variables_existentes = set(mod_pivot.index)
            var_venta = next((v for v in variables_existentes if is_venta(v)), 'Venta')
            
            # --- CAMBIO: OBTENER VALORES DE DICIEMBRE PARA REPLICAR FIJO Y MANIPULACIÓN EN PROYECCIÓN ---
            col_diciembre = next((c for c in mod_pivot.columns if c.month == 12), None)
            if col_diciembre is not None:
                val_fijo_diciembre = mod_pivot.at['Fijo', col_diciembre] if 'Fijo' in mod_pivot.index else 0.0
                val_manipulacion_diciembre = mod_pivot.at['Manipulación', col_diciembre] if 'Manipulación' in mod_pivot.index else 0.0
                
                venta_diciembre = mod_pivot.at[var_venta, col_diciembre] if var_venta in mod_pivot.index and mod_pivot.at[var_venta, col_diciembre] > 0 else 1.0
                pct_costo_diciembre = (mod_pivot.at['Costo', col_diciembre] / venta_diciembre) if 'Costo' in mod_pivot.index else 0.0
                pct_variable_diciembre = (mod_pivot.at['Variable', col_diciembre] / venta_diciembre) if 'Variable' in mod_pivot.index else 0.0
            else:
                val_fijo_diciembre = mod_pivot.loc['Fijo'].mean() if 'Fijo' in mod_pivot.index else 0.0
                val_manipulacion_diciembre = mod_pivot.loc['Manipulación'].mean() if 'Manipulación' in mod_pivot.index else 0.0
                
                venta_promedio = mod_pivot.loc[var_venta].mean() if var_venta in mod_pivot.index and mod_pivot.loc[var_venta].mean() > 0 else 1.0
                pct_costo_diciembre = (mod_pivot.loc['Costo'].mean() / venta_promedio) if 'Costo' in mod_pivot.index else 0.0
                pct_variable_diciembre = (mod_pivot.loc['Variable'].mean() / venta_promedio) if 'Variable' in mod_pivot.index else 0.0
            
            # Crear DataFrame final de proyección
            df_proyeccion = pd.DataFrame(index=['Venta', 'Costo', 'Manipulación', 'Fijo', 'Variable', 'Margen', 'Días Hábiles'], columns=fechas_ciclo).fillna(0.0)
            tipo_mes_list = []
            
            for d in fechas_ciclo:
                dias_proy = get_dias_habiles(d.year, d.month)
                df_proyeccion.at['Días Hábiles', d] = dias_proy
                
                if d in fechas_reales:
                    tipo_mes_list.append("Real")
                    # Valores reales directos de la modelación
                    df_proyeccion.at['Venta', d] = mod_pivot.at[var_venta, d] if var_venta in mod_pivot.index else 0.0
                    df_proyeccion.at['Costo', d] = mod_pivot.at['Costo', d] if 'Costo' in mod_pivot.index else 0.0
                    df_proyeccion.at['Manipulación', d] = mod_pivot.at['Manipulación', d] if 'Manipulación' in mod_pivot.index else 0.0
                    df_proyeccion.at['Fijo', d] = mod_pivot.at['Fijo', d] if 'Fijo' in mod_pivot.index else 0.0
                    df_proyeccion.at['Variable', d] = mod_pivot.at['Variable', d] if 'Variable' in mod_pivot.index else 0.0
                    
                    # Margen real = Venta - Costos reales
                    costos_totales = df_proyeccion.at['Costo', d] + df_proyeccion.at['Manipulación', d] + df_proyeccion.at['Fijo', d] + df_proyeccion.at['Variable', d]
                    df_proyeccion.at['Margen', d] = df_proyeccion.at['Venta', d] - costos_totales
                else:
                    tipo_mes_list.append("Proyectado")
                    # Buscar el mes correspondiente original en el año base
                    fecha_orig = pd.to_datetime(f"01-{d.month:02d}-{year_base}", format="%d-%m-%Y")
                    venta_orig = mod_pivot.at[var_venta, fecha_orig] if var_venta in mod_pivot.index and fecha_orig in mod_pivot.columns else 0.0
                    dias_orig = get_dias_habiles(year_base, d.month)
                    
                    venta_diaria = safe_div(venta_orig, dias_orig)
                    venta_proyectada = venta_diaria * dias_proy * (1.0 + aumento_auto)
                    df_proyeccion.at['Venta', d] = venta_proyectada
                    
                    # --- CAMBIO: Para los meses proyectados, traer Fijo y Manipulación desde DICIEMBRE (monto absoluto) ---
                    df_proyeccion.at['Fijo', d] = val_fijo_diciembre
                    df_proyeccion.at['Manipulación', d] = val_manipulacion_diciembre
                    
                    # Costo y Gasto Variable: proporcionales usando ratio de DICIEMBRE
                    costo_proy = pct_costo_diciembre * venta_proyectada
                    var_proy = pct_variable_diciembre * venta_proyectada
                    df_proyeccion.at['Costo', d] = costo_proy
                    df_proyeccion.at['Variable', d] = var_proy
                    
                    # Margen proyectado
                    costos_totales = costo_proy + val_manipulacion_diciembre + val_fijo_diciembre + var_proy
                    df_proyeccion.at['Margen', d] = venta_proyectada - costos_totales

            # Formatear el DataFrame para visualización con 2 decimales
            df_proy_disp = df_proyeccion.copy()
            df_proy_disp.columns = meses_str
            
            # Limitar a 2 decimales
            for col in df_proy_disp.columns:
                df_proy_disp[col] = df_proy_disp[col].round(2)
                
            df_proy_disp.loc['Tipo de Mes'] = tipo_mes_list
            
            ordered_rows = ['Tipo de Mes', 'Venta', 'Costo', 'Manipulación', 'Fijo', 'Variable', 'Margen', 'Días Hábiles']
            df_proy_disp = df_proy_disp.reindex(ordered_rows)
            
            sumas_totales = []
            for idx in ordered_rows:
                if idx == 'Tipo de Mes':
                    sumas_totales.append("---")
                else:
                    sumas_totales.append(round(df_proyeccion.loc[idx].sum(), 2))
            df_proy_disp['TOTAL MODELO'] = sumas_totales
            
            # --- KPI INDICATORS FOR SALES INCREASE ---
            st.markdown("### 📊 Indicadores de Crecimiento Considerados")
            # Venta original en los meses a proyectar
            venta_orig_acumulada = 0.0
            for d in fechas_proyectadas:
                fecha_orig = pd.to_datetime(f"01-{d.month:02d}-{year_base}", format="%d-%m-%Y")
                venta_orig_acumulada += mod_pivot.at[var_venta, fecha_orig] if var_venta in mod_pivot.index and fecha_orig in mod_pivot.columns else 0.0
            
            venta_proy_acumulada = df_proyeccion.loc['Venta', fechas_proyectadas].sum()
            incremento_monto = venta_proy_acumulada - venta_orig_acumulada
            
            col_kpi1, col_kpi2, col_kpi3 = st.columns(3)
            with col_kpi1:
                st.metric(
                    label="Aumento Porcentual de Venta",
                    value=f"{aumento_auto:.2%}"
                )
            with col_kpi2:
                st.metric(
                    label="Incremento en Monto de Venta",
                    value=f"${incremento_monto:,.2f}"
                )
            with col_kpi3:
                st.metric(
                    label="Venta Total Proyectada (Meses Proy)",
                    value=f"${venta_proy_acumulada:,.2f}"
                )
                
            # --- RENDERIZACIÓN ---
            st.markdown(f"#### 📋 Modelación Proyectada: {nombre_cliente} (CC: {cc_cliente})")
            st.caption(f"Fijos y Manipulación indexados a Diciembre: **Fijo: ${val_fijo_diciembre:,.2f}**, **Manipulación: ${val_manipulacion_diciembre:,.2f}** | Costo y Variable indexados a Diciembre: **Costo: {pct_costo_diciembre:.2%}**, **Variable: {pct_variable_diciembre:.2%}**")
            
            st.dataframe(df_proy_disp, use_container_width=True)
            
            # --- GRÁFICO DE ESTACIONALIDAD CRONOLÓGICO CON COLORES ---
            st.write("")
            st.markdown("##### 📈 Análisis Gráfico de la Estacionalidad (Venta y Margen por Estado)")
            
            # Separar Real de Proyectado para tener colores diferentes
            chart_data = pd.DataFrame(index=fechas_ciclo)
            chart_data['Venta (Real)'] = [df_proyeccion.at['Venta', d] if d in fechas_reales else 0.0 for d in fechas_ciclo]
            chart_data['Venta (Proyectada)'] = [df_proyeccion.at['Venta', d] if d in fechas_proyectadas else 0.0 for d in fechas_ciclo]
            chart_data['Margen (Real)'] = [df_proyeccion.at['Margen', d] if d in fechas_reales else 0.0 for d in fechas_ciclo]
            chart_data['Margen (Proyectada)'] = [df_proyeccion.at['Margen', d] if d in fechas_proyectadas else 0.0 for d in fechas_ciclo]
            
            # Para evitar que Streamlit ordene alfabéticamente el índice en el eje X, 
            # mantenemos el índice como objetos de tipo 'date'. Streamlit lo ordena cronológicamente.
            chart_data.index = [f.date() for f in chart_data.index]
            
            # --- CAMBIO: Tonos de azul para lo real y rojo para la modelación proyectada ---
            st.bar_chart(chart_data, color=["#1E3A8A", "#EF4444", "#3B82F6", "#F87171"])
            
            # --- EXPORTAR EXCEL DE ESTA PROYECCIÓN ---
            output_v2 = io.BytesIO()
            with pd.ExcelWriter(output_v2, engine='openpyxl') as writer:
                df_proy_disp.to_excel(writer, index=True, sheet_name='Proyección Personalizada')
            
            st.write(" ")
            st.download_button(
                label="📥 Descargar Proyección a Excel",
                data=output_v2.getvalue(),
                file_name=f"Proyeccion_Modelacion_{cc_cliente}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_dl_excel_tab2"
            )
            
        except Exception as e:
            st.error(f"Error al procesar la proyección: {e}")
            import traceback
            st.code(traceback.format_exc())
