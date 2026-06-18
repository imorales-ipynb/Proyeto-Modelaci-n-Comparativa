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
    .temporalidad-box { background-color: #fffbeb; border-left: 5px solid #f59e0b; padding: 1rem; border-radius: 4px; margin-bottom: 1rem; color: #92400e; }
    .no-temporalidad-box { background-color: #f0fdf4; border-left: 5px solid #22c55e; padding: 1rem; border-radius: 4px; margin-bottom: 1rem; color: #166534; }
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
    return int(np.busday_count(start_date, end_date, holidays=feriados))

def safe_div(n, d):
    return n / d if d and d != 0 else 0

def is_venta(var_name):
    return 'venta' in str(var_name).lower()

def is_margen(var_name):
    return 'margen' in str(var_name).lower()


# ─────────────────────────────────────────────────────────────────────────────
# NÚCLEO: detección de temporalidad y proyección con la nueva lógica
# ─────────────────────────────────────────────────────────────────────────────

def detectar_temporalidad(mod_pivot, var_venta, umbral=0.20):
    """
    Compara cada mes de la modelación contra octubre (mes base típico)
    usando la venta proporcional diaria.
    Devuelve:
      - tiene_temporalidad (bool)
      - meses_con_temporalidad (dict {mes_num: pct_diferencia})
    """
    col_octubre = next((c for c in mod_pivot.columns if c.month == 10), None)
    if col_octubre is None or var_venta not in mod_pivot.index:
        return False, {}

    venta_octubre = mod_pivot.at[var_venta, col_octubre]
    if venta_octubre == 0:
        return False, {}

    dias_octubre = get_dias_habiles(col_octubre.year, 10)
    v_diaria_oct = safe_div(venta_octubre, dias_octubre)

    if v_diaria_oct == 0:
        return False, {}

    meses_temp = {}
    for col in mod_pivot.columns:
        venta_mes = mod_pivot.at[var_venta, col]
        if venta_mes == 0:
            continue
            
        dias_mes = get_dias_habiles(col.year, col.month)
        v_diaria_mes = safe_div(venta_mes, dias_mes)
        
        diferencia = (v_diaria_mes - v_diaria_oct) / abs(v_diaria_oct)
        
        if abs(diferencia) > umbral:
            meses_temp[col.month] = diferencia

    return len(meses_temp) > 0, meses_temp


def calcular_variacion_oct_dic(mod_pivot, var_venta):
    """
    Calcula el promedio de variación absoluta y porcentual entre octubre y diciembre
    para usarlo como diferencial de proyección cuando no hay temporalidad.
    La comparación se hace usando la venta proporcional diaria.
    Retorna (variacion_pct_promedio, variacion_abs_promedio).
    Si hay más de un paso oct→nov→dic, promedia los incrementos mensuales.
    """
    meses_ventana = []
    for mes in [10, 11, 12]:
        col = next((c for c in mod_pivot.columns if c.month == mes), None)
        if col is not None and var_venta in mod_pivot.index:
            v = mod_pivot.at[var_venta, col]
            dias_mes = get_dias_habiles(col.year, mes)
            v_diaria = safe_div(v, dias_mes)
            if v_diaria != 0:
                meses_ventana.append((col, v_diaria))

    if len(meses_ventana) < 2:
        return 0.0, 0.0

    variaciones_pct = []
    variaciones_abs = []
    for i in range(1, len(meses_ventana)):
        v_ant = meses_ventana[i - 1][1]
        v_act = meses_ventana[i][1]
        if v_ant != 0:
            variaciones_pct.append((v_act - v_ant) / abs(v_ant))
            variaciones_abs.append(v_act - v_ant)

    if not variaciones_pct:
        return 0.0, 0.0

    return float(np.mean(variaciones_pct)), float(np.mean(variaciones_abs))


def proyectar_venta(mod_pivot, var_venta, mes_num, year_base, year_proy,
                    tiene_temporalidad, meses_con_temporalidad,
                    variacion_pct_oct_dic, variacion_abs_oct_dic,
                    aumento_extra_pct=0.0):
    """
    Devuelve la venta proyectada para (mes_num, year_proy) ajustada por días hábiles.

    Lógica:
      - Comparar mes_num de la modelación contra octubre en base a venta diaria.
      - Si |diferencia_diaria| > 20%  → usar el valor proporcional del mes de la modelación.
      - Si |diferencia_diaria| <= 20% → usar venta mensual de octubre + variación oct-dic.
    """
    col_octubre = next((c for c in mod_pivot.columns if c.month == 10), None)
    col_mes_base = next((c for c in mod_pivot.columns if c.month == mes_num), None)

    venta_base_raw = 0.0
    v_base_diaria = 0.0
    dias_base_real = get_dias_habiles(year_base, mes_num)
    if col_mes_base is not None and var_venta in mod_pivot.index:
        venta_base_raw = float(mod_pivot.at[var_venta, col_mes_base])
        dias_base_real = get_dias_habiles(col_mes_base.year, mes_num)
        v_base_diaria = safe_div(venta_base_raw, dias_base_real)

    venta_octubre = 0.0
    v_octubre_diaria = 0.0
    dias_oct_real = get_dias_habiles(year_base, 10)
    if col_octubre is not None and var_venta in mod_pivot.index:
        venta_octubre = float(mod_pivot.at[var_venta, col_octubre])
        dias_oct_real = get_dias_habiles(col_octubre.year, 10)
        v_octubre_diaria = safe_div(venta_octubre, dias_oct_real)

    # Decidir qué valor base usar
    diferencia = 0.0
    if v_octubre_diaria != 0 and v_base_diaria != 0:
        diferencia = (v_base_diaria - v_octubre_diaria) / abs(v_octubre_diaria)

    if abs(diferencia) > 0.20:
        # Temporalidad detectada en este mes → usar valor proporcional de la modelación escalado a los días de proyección
        dias_proy = get_dias_habiles(year_proy, mes_num)
        venta_proyectada = v_base_diaria * dias_proy
    else:
        # Sin temporalidad significativa → venta diaria de octubre + variación, multiplicada por días hábiles del mes proyectado
        if v_octubre_diaria != 0:
            v_proy_pct = v_octubre_diaria * (1 + variacion_pct_oct_dic)
            v_proy_abs = v_octubre_diaria + variacion_abs_oct_dic
            venta_diaria_final = (v_proy_pct + v_proy_abs) / 2
        else:
            venta_diaria_final = 0.0
            
        dias_proy = get_dias_habiles(year_proy, mes_num)
        venta_proyectada = venta_diaria_final * dias_proy

    # Aplicar aumento extra (EVA vs Mod si aplica)
    venta_proyectada *= (1.0 + aumento_extra_pct)

    return venta_proyectada


def proyectar_costos(mod_pivot, var, venta_proyectada, mes_num,
                     tiene_temporalidad,
                     val_fijo_diciembre, val_manipulacion_diciembre,
                     pct_costo_diciembre, pct_variable_diciembre):
    """
    Proyecta cada variable de costo/gasto según la lógica de temporalidad.

    Sin temporalidad:
      - Fijo / Manipulación → valor absoluto de Diciembre
      - Costo / Variable    → % de Diciembre sobre venta proyectada

    Con temporalidad:
      - Fijo / Manipulación → valor absoluto del mes correspondiente en la modelación
      - Costo / Variable    → % del mes correspondiente en la modelación sobre venta proyectada
    """
    var_lower = str(var).lower()
    es_fijo_o_manip = 'fijo' in var_lower or 'manipulación' in var_lower or 'manipulacion' in var_lower
    es_proporcional = 'costo' in var_lower or 'variable' in var_lower

    col_mes = next((c for c in mod_pivot.columns if c.month == mes_num), None)

    if not tiene_temporalidad:
        # Lógica original: todo referenciado a Diciembre
        if 'fijo' in var_lower:
            return val_fijo_diciembre
        elif 'manipulación' in var_lower or 'manipulacion' in var_lower:
            return val_manipulacion_diciembre
        elif 'costo' in var_lower:
            return pct_costo_diciembre * venta_proyectada
        elif 'variable' in var_lower:
            return pct_variable_diciembre * venta_proyectada
        else:
            return 0.0
    else:
        # Con temporalidad: usar el % o valor absoluto del mes correspondiente
        if col_mes is not None and var in mod_pivot.index:
            val_mes = float(mod_pivot.at[var, col_mes])
            if es_fijo_o_manip:
                return val_mes
            elif es_proporcional:
                # Calcular % sobre venta real del mes
                var_venta_key = next((v for v in mod_pivot.index if is_venta(v)), None)
                venta_mes_real = float(mod_pivot.at[var_venta_key, col_mes]) if var_venta_key else 0.0
                pct_mes = safe_div(val_mes, venta_mes_real)
                return pct_mes * venta_proyectada
            else:
                return val_mes
        else:
            # Fallback a Diciembre si no hay dato del mes
            if 'fijo' in var_lower:
                return val_fijo_diciembre
            elif 'manipulación' in var_lower or 'manipulacion' in var_lower:
                return val_manipulacion_diciembre
            elif 'costo' in var_lower:
                return pct_costo_diciembre * venta_proyectada
            elif 'variable' in var_lower:
                return pct_variable_diciembre * venta_proyectada
            else:
                return 0.0


def render_temporalidad_info(tiene_temporalidad, meses_con_temporalidad):
    """Renderiza el bloque informativo de temporalidad."""
    MESES_ES = {
        1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio',
        7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
    }
    if tiene_temporalidad:
        detalle = ", ".join(
            [f"<b>{MESES_ES.get(m, m)}</b> ({v:+.1%})" for m, v in sorted(meses_con_temporalidad.items())]
        )
        st.markdown(
            f'<div class="temporalidad-box">⚠️ <b>Temporalidad detectada</b> — Los siguientes meses presentan una diferencia superior al 20% respecto a Octubre (mes base): {detalle}</div>',
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            '<div class="no-temporalidad-box">✅ <b>Sin temporalidad significativa</b> — Ningún mes supera el 20% de diferencia respecto a Octubre. Se utilizará la lógica de proyección estándar (Octubre + variación Oct-Dic).</div>',
            unsafe_allow_html=True
        )


# ─────────────────────────────────────────────────────────────────────────────
# PROCESAMIENTO PRINCIPAL (Tab 1: EVA vs Modelación)
# ─────────────────────────────────────────────────────────────────────────────

def process_data(df):
    required_cols = ['CC', 'Nombre Cliente', 'Tipo Modelo', 'Variable']
    for col in required_cols:
        if col not in df.columns:
            return None, None, f"Falta la columna requerida: {col}"

    month_cols = [col for col in df.columns if col not in required_cols]
    
    df['CC'] = df['CC'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df['Nombre Cliente'] = df['Nombre Cliente'].astype(str).str.strip().str.upper()
    df['Tipo Modelo'] = df['Tipo Modelo'].astype(str).str.strip().str.upper()
    df['Tipo Modelo'] = df['Tipo Modelo'].replace('MODELACIÓN', 'MODELACION')
    df['Variable'] = df['Variable'].astype(str).str.strip()
    
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
        
        fechas_disponibles_mod = [col for col in mod_pivot_raw.columns if any(mod_pivot_raw[col] != 0)]
        ultimo_mes_mod = max(fechas_disponibles_mod) if fechas_disponibles_mod else None
        year_base_mod = mod_pivot_raw.columns.min().year if not mod_pivot_raw.empty and len(mod_pivot_raw.columns) > 0 else fecha_inicio.year

        # ── Detección de temporalidad para este CC ──
        tiene_temporalidad, meses_con_temporalidad = detectar_temporalidad(mod_pivot_raw, var_venta)
        variacion_pct_oct_dic, variacion_abs_oct_dic = calcular_variacion_oct_dic(mod_pivot_raw, var_venta)

        # ── Valores de referencia de Diciembre (para lógica sin temporalidad) ──
        col_diciembre = next((c for c in mod_pivot_raw.columns if c.month == 12), None)
        if col_diciembre is not None:
            val_fijo_dic = float(mod_pivot_raw.at['Fijo', col_diciembre]) if 'Fijo' in mod_pivot_raw.index else 0.0
            val_manip_dic = float(mod_pivot_raw.at['Manipulación', col_diciembre]) if 'Manipulación' in mod_pivot_raw.index else 0.0
            venta_dic = float(mod_pivot_raw.at[var_venta, col_diciembre]) if var_venta and var_venta in mod_pivot_raw.index else 1.0
            if venta_dic == 0:
                venta_dic = 1.0
            pct_costo_dic = safe_div(float(mod_pivot_raw.at['Costo', col_diciembre]) if 'Costo' in mod_pivot_raw.index else 0.0, venta_dic)
            pct_variable_dic = safe_div(float(mod_pivot_raw.at['Variable', col_diciembre]) if 'Variable' in mod_pivot_raw.index else 0.0, venta_dic)
        else:
            val_fijo_dic = float(mod_pivot_raw.loc['Fijo'].mean()) if 'Fijo' in mod_pivot_raw.index else 0.0
            val_manip_dic = float(mod_pivot_raw.loc['Manipulación'].mean()) if 'Manipulación' in mod_pivot_raw.index else 0.0
            venta_prom = float(mod_pivot_raw.loc[var_venta].mean()) if var_venta and var_venta in mod_pivot_raw.index else 1.0
            if venta_prom == 0:
                venta_prom = 1.0
            pct_costo_dic = safe_div(float(mod_pivot_raw.loc['Costo'].mean()) if 'Costo' in mod_pivot_raw.index else 0.0, venta_prom)
            pct_variable_dic = safe_div(float(mod_pivot_raw.loc['Variable'].mean()) if 'Variable' in mod_pivot_raw.index else 0.0, venta_prom)

        for d in fechas_ciclo_eva:
            # 1. Cargar EVA
            for var in variables:
                if var in eva_pivot.index and d in eva_pivot.columns:
                    eva_12m.at[var, d] = eva_pivot.at[var, d]
            
            # 2. Cargar Modelación Ajustada
            if d in mod_pivot_raw.columns and any(mod_pivot_raw[d] != 0):
                # Mes real en la modelación → copiar directamente
                for var in variables:
                    if var in mod_pivot_raw.index:
                        mod_12m.at[var, d] = mod_pivot_raw.at[var, d]
            else:
                # ── PROYECCIÓN CON NUEVA LÓGICA ──
                if var_venta:
                    venta_proyectada = proyectar_venta(
                        mod_pivot=mod_pivot_raw,
                        var_venta=var_venta,
                        mes_num=d.month,
                        year_base=year_base_mod,
                        year_proy=d.year,
                        tiene_temporalidad=tiene_temporalidad,
                        meses_con_temporalidad=meses_con_temporalidad,
                        variacion_pct_oct_dic=variacion_pct_oct_dic,
                        variacion_abs_oct_dic=variacion_abs_oct_dic,
                        aumento_extra_pct=0.0
                    )
                    mod_12m.at[var_venta, d] = venta_proyectada
                else:
                    venta_proyectada = 0.0

                costos_gastos = 0.0
                for var in variables:
                    if var == var_venta or var == var_margen:
                        continue
                    val_proy = proyectar_costos(
                        mod_pivot=mod_pivot_raw,
                        var=var,
                        venta_proyectada=venta_proyectada,
                        mes_num=d.month,
                        tiene_temporalidad=tiene_temporalidad,
                        val_fijo_diciembre=val_fijo_dic,
                        val_manipulacion_diciembre=val_manip_dic,
                        pct_costo_diciembre=pct_costo_dic,
                        pct_variable_diciembre=pct_variable_dic
                    )
                    mod_12m.at[var, d] = val_proy
                    costos_gastos += val_proy

                if var_margen:
                    mod_12m.at[var_margen, d] = venta_proyectada - costos_gastos

        # ── CÁLCULO DE PORCENTAJES Y DIFERENCIALES ──
        for tipo, df_datos in [('EVA', eva_12m), ('Modelación Ajustada', mod_12m)]:
            fila_dias = {'CC': cc, 'Nombre Cliente': nombre, 'Tipo Modelo': tipo, 'Variable': 'Días Hábiles'}
            for i, d in enumerate(fechas_ciclo_eva):
                fila_dias[meses_str[i]] = get_dias_habiles(d.year, d.month)
            resultados_mensuales.append(fila_dias)
            
            for var in variables:
                fila = {'CC': cc, 'Nombre Cliente': nombre, 'Tipo Modelo': tipo, 'Variable': var}
                for i, d in enumerate(fechas_ciclo_eva):
                    fila[meses_str[i]] = df_datos.at[var, d]
                resultados_mensuales.append(fila)
                
                if var_venta and var != var_venta:
                    fila_pct = {'CC': cc, 'Nombre Cliente': nombre, 'Tipo Modelo': tipo, 'Variable': f"{var} %"}
                    for i, d in enumerate(fechas_ciclo_eva):
                        venta_mes = df_datos.at[var_venta, d]
                        valor_mes = df_datos.at[var, d]
                        fila_pct[meses_str[i]] = safe_div(valor_mes, venta_mes)
                    resultados_mensuales.append(fila_pct)
                    
        dif_12m = eva_12m - mod_12m
        for var in variables:
            fila_dif = {'CC': cc, 'Nombre Cliente': nombre, 'Tipo Modelo': 'Diferencial (EVA - Mod)', 'Variable': var}
            for i, d in enumerate(fechas_ciclo_eva):
                fila_dif[meses_str[i]] = dif_12m.at[var, d]
            resultados_mensuales.append(fila_dif)

        # ── RESUMEN ANUALIZADO ──
        resumen_cli = {'CC': cc, 'Nombre Cliente': nombre}
        
        suma_ventas_eva = eva_12m.loc[var_venta].sum() if var_venta else 0
        suma_ventas_mod = mod_12m.loc[var_venta].sum() if var_venta else 0
        
        suma_mc_eva = eva_12m.loc[var_margen].sum() if var_margen else 0
        suma_mc_mod = mod_12m.loc[var_margen].sum() if var_margen else 0
        
        mc_pct_eva = safe_div(suma_mc_eva, suma_ventas_eva)
        mc_pct_mod = safe_div(suma_mc_mod, suma_ventas_mod)
        
        dif_abs_mc = suma_mc_eva - suma_mc_mod
        dif_pct_mc = mc_pct_eva - mc_pct_mod
        
        total_dias_habiles = sum([get_dias_habiles(d.year, d.month) for d in fechas_ciclo_eva])
        
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
        resumen_cli['Tiene Temporalidad'] = tiene_temporalidad
        resumen_cli['Meses con Temporalidad'] = meses_con_temporalidad
        
        for var in variables:
            total_eva = eva_12m.loc[var].sum()
            total_mod = mod_12m.loc[var].sum()
            resumen_cli[f'Total {var} EVA'] = total_eva
            resumen_cli[f'Total {var} Modelación'] = total_mod
            resumen_cli[f'Dif {var}'] = total_eva - total_mod
            resumen_cli[f'% {var} EVA'] = safe_div(total_eva, suma_ventas_eva)
            resumen_cli[f'% {var} Modelación'] = safe_div(total_mod, suma_ventas_mod)
        
        for var in variables:
            resumen_cli[f'Promedio Mensual {var} (EVA)'] = eva_12m.loc[var].mean()
            resumen_cli[f'Promedio Mensual {var} (Mod)'] = mod_12m.loc[var].mean()
            
        resumen_ejecutivo.append(resumen_cli)

    if not resultados_mensuales:
        return None, None, "No se encontraron datos procesables."

    df_mensual = pd.DataFrame(resultados_mensuales)
    df_resumen = pd.DataFrame(resumen_ejecutivo)
    
    return df_mensual, df_resumen, None


# ─────────────────────────────────────────────────────────────────────────────
# REPOSITORIO
# ─────────────────────────────────────────────────────────────────────────────

REPO_DIR = "repo_modelaciones"
if not os.path.exists(REPO_DIR):
    os.makedirs(REPO_DIR)

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


# ─────────────────────────────────────────────────────────────────────────────
# TAB 1: COMPARATIVA EVA vs MODELACIÓN
# ─────────────────────────────────────────────────────────────────────────────

tab1, tab2, tab3 = st.tabs(["📊 Comparativa EVA vs Modelación", "🔍 Visualizador y Proyección de Modelación", "🗓️ Proyección 2027 (Todos los CC)"])

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

                    # ── Bloque de temporalidad ──
                    render_temporalidad_info(
                        row.get('Tiene Temporalidad', False),
                        row.get('Meses con Temporalidad', {})
                    )
                    
                    st.caption(f"🗓️ **Período Analizado:** 12 Meses | **Días Hábiles Totales:** {row['Total Días Hábiles (12m)']} días")
                    
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
                    
                    st.markdown("**📐 Diferencias (EVA - Modelación)**")
                    df_dif_view = df_cli[(df_cli['Tipo Modelo'] == 'Diferencial (EVA - Mod)') & (~df_cli['Variable'].str.endswith(' %'))].copy()
                    df_dif_disp = df_dif_view.drop(columns=['CC', 'Nombre Cliente', 'Tipo Modelo'])
                    df_dif_disp['Variable'] = pd.Categorical(df_dif_disp['Variable'], categories=desired_order, ordered=True)
                    df_dif_disp = df_dif_disp.sort_values('Variable').reset_index(drop=True)
                    mes_cols_dif = [c for c in df_dif_disp.columns if c != 'Variable']
                    df_dif_disp['TOTAL'] = df_dif_disp[mes_cols_dif].apply(pd.to_numeric, errors='coerce').sum(axis=1)
                    
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


# ─────────────────────────────────────────────────────────────────────────────
# TAB 2: VISUALIZADOR Y PROYECCIÓN DE MODELACIÓN
# ─────────────────────────────────────────────────────────────────────────────

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
            
            df_mod_raw_all['CC'] = df_mod_raw_all['CC'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            df_mod_raw_all['Nombre Cliente'] = df_mod_raw_all['Nombre Cliente'].astype(str).str.strip().str.upper()
            df_mod_raw_all['Variable'] = df_mod_raw_all['Variable'].astype(str).str.strip()
            
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
            
            df_mod_raw = df_mod_raw_all[df_mod_raw_all['CC'] == cc_seleccionado].copy()
            
            mapping_vars = {
                'Costo Alimento': 'Costo',
                'Gasto Manipulación': 'Manipulación',
                'Gasto Fijo': 'Fijo',
                'Gasto Variable': 'Variable',
                'Margen de Contribución': 'Margen'
            }
            df_mod_raw['Variable'] = df_mod_raw['Variable'].replace(mapping_vars)
            
            nombre_cliente = df_mod_raw['Nombre Cliente'].iloc[0] if not df_mod_raw.empty else "CLIENTE DESCONOCIDO"
            cc_cliente = cc_seleccionado
            
            excluded_cols = ['CC', 'Nombre Cliente', 'Tipo Modelo', 'Variable']
            month_cols = [col for col in df_mod_raw.columns if col not in excluded_cols]
            
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
            
            MESES_ES = {
                1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio',
                7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
            }
            
            # ── Pivot y detección de temporalidad ──
            mod_pivot = df_melt.pivot_table(index='Variable', columns='Fecha', values='Valor', aggfunc='sum').fillna(0)
            variables_existentes = set(mod_pivot.index)
            var_venta = next((v for v in variables_existentes if is_venta(v)), 'Venta')

            tiene_temporalidad, meses_con_temporalidad = detectar_temporalidad(mod_pivot, var_venta)
            variacion_pct_oct_dic, variacion_abs_oct_dic = calcular_variacion_oct_dic(mod_pivot, var_venta)

            # ── Valores de referencia de Diciembre ──
            col_diciembre = next((c for c in mod_pivot.columns if c.month == 12), None)
            if col_diciembre is not None:
                val_fijo_diciembre = mod_pivot.at['Fijo', col_diciembre] if 'Fijo' in mod_pivot.index else 0.0
                val_manipulacion_diciembre = mod_pivot.at['Manipulación', col_diciembre] if 'Manipulación' in mod_pivot.index else 0.0
                venta_diciembre = mod_pivot.at[var_venta, col_diciembre] if var_venta in mod_pivot.index and mod_pivot.at[var_venta, col_diciembre] > 0 else 1.0
                pct_costo_diciembre = safe_div(mod_pivot.at['Costo', col_diciembre] if 'Costo' in mod_pivot.index else 0.0, venta_diciembre)
                pct_variable_diciembre = safe_div(mod_pivot.at['Variable', col_diciembre] if 'Variable' in mod_pivot.index else 0.0, venta_diciembre)
            else:
                val_fijo_diciembre = mod_pivot.loc['Fijo'].mean() if 'Fijo' in mod_pivot.index else 0.0
                val_manipulacion_diciembre = mod_pivot.loc['Manipulación'].mean() if 'Manipulación' in mod_pivot.index else 0.0
                venta_promedio = mod_pivot.loc[var_venta].mean() if var_venta in mod_pivot.index and mod_pivot.loc[var_venta].mean() > 0 else 1.0
                pct_costo_diciembre = safe_div(mod_pivot.loc['Costo'].mean() if 'Costo' in mod_pivot.index else 0.0, venta_promedio)
                pct_variable_diciembre = safe_div(mod_pivot.loc['Variable'].mean() if 'Variable' in mod_pivot.index else 0.0, venta_promedio)

            # ── Mostrar bloque de temporalidad ──
            st.markdown("### 📅 Análisis de Temporalidad")
            render_temporalidad_info(tiene_temporalidad, meses_con_temporalidad)

            # ── Detalle técnico de la proyección ──
            if not tiene_temporalidad:
                variacion_pct_display = variacion_pct_oct_dic
                variacion_abs_display = variacion_abs_oct_dic
                st.caption(
                    f"📐 **Lógica de proyección:** Sin temporalidad → Venta diaria base de Octubre con variación promedio Oct-Dic "
                    f"(**{variacion_pct_display:+.2%}** | **${variacion_abs_display:+,.0f}/día**), multiplicada por días hábiles del mes. "
                    f"Costos referenciados a **Diciembre** — Fijo: **${val_fijo_diciembre:,.0f}**, "
                    f"Manipulación: **${val_manipulacion_diciembre:,.0f}**, "
                    f"Costo: **{pct_costo_diciembre:.2%}**, Variable: **{pct_variable_diciembre:.2%}**."
                )
            else:
                st.caption(
                    "📐 **Lógica de proyección:** Con temporalidad → meses con >20% de diferencia vs Octubre usan "
                    "el valor real de la modelación; el resto usa Octubre + variación Oct-Dic. "
                    "Costos y gastos se calculan con el **% real del mes proyectado** de la modelación."
                )

            st.markdown("---")
            st.markdown("### ⚙️ Configuración de la Proyección")
            
            col_c1, col_c2 = st.columns(2)
            
            with col_c1:
                mes_limite_nombre = st.selectbox(
                    "Proyectar hasta el mes de:",
                    options=list(MESES_ES.values()),
                    index=4
                )
                mes_limite_num = next(k for k, v in MESES_ES.items() if v == mes_limite_nombre)
                
            fechas_reales = [pd.to_datetime(f"01-{m:02d}-{year_base}", format="%d-%m-%Y") for m in range(mes_limite_num + 1, 13)]
            fechas_proyectadas = [pd.to_datetime(f"01-{m:02d}-{year_base + 1}", format="%d-%m-%Y") for m in range(1, mes_limite_num + 1)]
            fechas_ciclo = fechas_reales + fechas_proyectadas
            meses_str = [f.strftime('%m-%Y') for f in fechas_ciclo]
            
            with col_c2:
                st.write("**Resumen de Períodos:**")
                st.write(f"🟢 **Meses Reales ({year_base}):** {', '.join([MESES_ES[f.month] for f in fechas_reales]) if fechas_reales else 'Ninguno'}")
                st.write(f"🔵 **Meses Proyectados ({year_base + 1}):** {', '.join([MESES_ES[f.month] for f in fechas_proyectadas]) if fechas_proyectadas else 'Ninguno'}")
                
            # ── Aumento extra por EVA (si está disponible) ──
            aumento_auto = 0.0
            usar_eva = False
            
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
                        
                        ventas_mod_real_sum = 0.0
                        ventas_eva_real_sum = 0.0
                        
                        var_venta_eva = next((v for v in df_eva_melt['Variable'].unique() if is_venta(v)), 'Venta')
                        
                        for f in fechas_reales:
                            v_mod = df_melt[(df_melt['Variable'] == var_venta) & (df_melt['Fecha'] == f)]['Valor'].sum()
                            v_eva = df_eva_melt[(df_eva_melt['Variable'] == var_venta_eva) & (df_eva_melt['Fecha'].dt.month == f.month)]['Valor'].sum()
                            ventas_mod_real_sum += v_mod
                            ventas_eva_real_sum += v_eva
                            
                        if ventas_mod_real_sum > 0:
                            aumento_auto = (ventas_eva_real_sum - ventas_mod_real_sum) / ventas_mod_real_sum
                            usar_eva = True
                except Exception:
                    pass
            
            st.markdown("---")
            
            # ── PROYECCIÓN CON NUEVA LÓGICA (Tab 2) ──
            df_proyeccion = pd.DataFrame(
                index=['Venta', 'Costo', 'Manipulación', 'Fijo', 'Variable', 'Margen', 'Días Hábiles'],
                columns=fechas_ciclo
            ).fillna(0.0)
            tipo_mes_list = []
            
            for d in fechas_ciclo:
                dias_proy = get_dias_habiles(d.year, d.month)
                df_proyeccion.at['Días Hábiles', d] = dias_proy
                
                if d in fechas_reales:
                    tipo_mes_list.append("Real")
                    df_proyeccion.at['Venta', d] = float(mod_pivot.at[var_venta, d]) if var_venta in mod_pivot.index and d in mod_pivot.columns else 0.0
                    df_proyeccion.at['Costo', d] = float(mod_pivot.at['Costo', d]) if 'Costo' in mod_pivot.index and d in mod_pivot.columns else 0.0
                    df_proyeccion.at['Manipulación', d] = float(mod_pivot.at['Manipulación', d]) if 'Manipulación' in mod_pivot.index and d in mod_pivot.columns else 0.0
                    df_proyeccion.at['Fijo', d] = float(mod_pivot.at['Fijo', d]) if 'Fijo' in mod_pivot.index and d in mod_pivot.columns else 0.0
                    df_proyeccion.at['Variable', d] = float(mod_pivot.at['Variable', d]) if 'Variable' in mod_pivot.index and d in mod_pivot.columns else 0.0
                    costos_reales = (df_proyeccion.at['Costo', d] + df_proyeccion.at['Manipulación', d] +
                                     df_proyeccion.at['Fijo', d] + df_proyeccion.at['Variable', d])
                    df_proyeccion.at['Margen', d] = df_proyeccion.at['Venta', d] - costos_reales
                else:
                    tipo_mes_list.append("Proyectado")
                    
                    # Venta con nueva lógica
                    venta_proyectada = proyectar_venta(
                        mod_pivot=mod_pivot,
                        var_venta=var_venta,
                        mes_num=d.month,
                        year_base=year_base,
                        year_proy=d.year,
                        tiene_temporalidad=tiene_temporalidad,
                        meses_con_temporalidad=meses_con_temporalidad,
                        variacion_pct_oct_dic=variacion_pct_oct_dic,
                        variacion_abs_oct_dic=variacion_abs_oct_dic,
                        aumento_extra_pct=aumento_auto
                    )
                    df_proyeccion.at['Venta', d] = venta_proyectada

                    # Costos con nueva lógica
                    costos_totales = 0.0
                    for var in ['Costo', 'Manipulación', 'Fijo', 'Variable']:
                        val_proy = proyectar_costos(
                            mod_pivot=mod_pivot,
                            var=var,
                            venta_proyectada=venta_proyectada,
                            mes_num=d.month,
                            tiene_temporalidad=tiene_temporalidad,
                            val_fijo_diciembre=val_fijo_diciembre,
                            val_manipulacion_diciembre=val_manipulacion_diciembre,
                            pct_costo_diciembre=pct_costo_diciembre,
                            pct_variable_diciembre=pct_variable_diciembre
                        )
                        df_proyeccion.at[var, d] = val_proy
                        costos_totales += val_proy

                    df_proyeccion.at['Margen', d] = venta_proyectada - costos_totales

            # ── Formatear para display ──
            df_proy_disp = df_proyeccion.copy()
            df_proy_disp.columns = meses_str
            
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
            
            # ── KPIs ──
            st.markdown("### 📊 Indicadores de Crecimiento Considerados")
            venta_orig_acumulada = 0.0
            for d in fechas_proyectadas:
                fecha_orig = pd.to_datetime(f"01-{d.month:02d}-{year_base}", format="%d-%m-%Y")
                venta_orig_acumulada += float(mod_pivot.at[var_venta, fecha_orig]) if var_venta in mod_pivot.index and fecha_orig in mod_pivot.columns else 0.0
            
            venta_proy_acumulada = float(df_proyeccion.loc['Venta', fechas_proyectadas].sum())
            incremento_monto = venta_proy_acumulada - venta_orig_acumulada
            
            col_kpi1, col_kpi2, col_kpi3 = st.columns(3)
            with col_kpi1:
                st.metric(label="Aumento Base Proyección (Oct-Dic)", value=f"{variacion_pct_oct_dic:.2%}")
            with col_kpi2:
                st.metric(label="Incremento en Monto de Venta", value=f"${incremento_monto:,.2f}")
            with col_kpi3:
                st.metric(label="Venta Total Proyectada (Meses Proy)", value=f"${venta_proy_acumulada:,.2f}")
                
            st.markdown(f"#### 📋 Modelación Proyectada: {nombre_cliente} (CC: {cc_cliente})")

            # Caption con lógica activa
            if tiene_temporalidad:
                st.caption("🔶 Proyección con temporalidad: venta usa valor real de la modelación para meses con >20% vs Octubre. Costos y gastos calculados con porcentaje real del mes correspondiente.")
            else:
                st.caption(
                    f"🔷 Proyección sin temporalidad: venta calculada como (Venta Diaria Octubre + variación) × días hábiles. Variación Oct-Dic "
                    f"({variacion_pct_oct_dic:+.2%} | ${variacion_abs_oct_dic:+,.0f}/día). "
                    f"Fijo: ${val_fijo_diciembre:,.2f} | Manipulación: ${val_manipulacion_diciembre:,.2f} | "
                    f"Costo: {pct_costo_diciembre:.2%} | Variable: {pct_variable_diciembre:.2%} (todos referenciados a Diciembre)."
                )
            
            st.dataframe(df_proy_disp, use_container_width=True)
            
            # ── Gráfico ──
            st.write("")
            st.markdown("##### 📈 Análisis Gráfico de la Estacionalidad (Venta y Margen por Estado)")
            
            chart_data = pd.DataFrame(index=fechas_ciclo)
            chart_data['Venta (Real)'] = [float(df_proyeccion.at['Venta', d]) if d in fechas_reales else 0.0 for d in fechas_ciclo]
            chart_data['Venta (Proyectada)'] = [float(df_proyeccion.at['Venta', d]) if d in fechas_proyectadas else 0.0 for d in fechas_ciclo]
            chart_data['Margen (Real)'] = [float(df_proyeccion.at['Margen', d]) if d in fechas_reales else 0.0 for d in fechas_ciclo]
            chart_data['Margen (Proyectada)'] = [float(df_proyeccion.at['Margen', d]) if d in fechas_proyectadas else 0.0 for d in fechas_ciclo]
            chart_data.index = [f.date() for f in chart_data.index]
            
            st.bar_chart(chart_data, color=["#1E3A8A", "#EF4444", "#3B82F6", "#F87171"])
            
            # ── Exportar ──
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


# ─────────────────────────────────────────────────────────────────────────────
# TAB 3: PROYECCIÓN 2027 — TODOS LOS CC
# ─────────────────────────────────────────────────────────────────────────────

MESES_ES_NOMBRES = {
    1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio',
    7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
}

with tab3:
    st.header("Proyección 2027 — Todos los Centros de Costo")
    st.write(
        "Genera la proyección completa de enero a diciembre 2027 para **todos los CC** "
        "usando la misma lógica del Visualizador. El resultado se puede exportar a Excel como Modelación Proyectada 2027."
    )

    st.subheader("1. Seleccionar Modelación Base")
    if archivos_repo:
        mod_sel_tab3 = st.selectbox(
            "Elige la modelación base:",
            options=["Seleccionar..."] + archivos_repo,
            key="mod_sel_tab3"
        )
    else:
        st.warning("Sube un archivo en la barra lateral.")
        mod_sel_tab3 = "Seleccionar..."

    if mod_sel_tab3 != "Seleccionar...":
        try:
            with st.spinner("Calculando proyecciones 2027 para todos los centros de costo..."):
                path_mod_t3 = os.path.join(REPO_DIR, mod_sel_tab3)
                df_base_t3 = pd.read_excel(path_mod_t3)

                df_base_t3['CC'] = df_base_t3['CC'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                df_base_t3['Nombre Cliente'] = df_base_t3['Nombre Cliente'].astype(str).str.strip().str.upper()
                df_base_t3['Variable'] = df_base_t3['Variable'].astype(str).str.strip()

                mapping_vars_t3 = {
                    'Costo Alimento': 'Costo',
                    'Gasto Manipulación': 'Manipulación',
                    'Gasto Fijo': 'Fijo',
                    'Gasto Variable': 'Variable',
                    'Margen de Contribución': 'Margen'
                }
                df_base_t3['Variable'] = df_base_t3['Variable'].replace(mapping_vars_t3)

                excluded_t3 = ['CC', 'Nombre Cliente', 'Tipo Modelo', 'Variable']
                month_cols_t3 = [c for c in df_base_t3.columns if c not in excluded_t3]

                ccs_all_t3 = df_base_t3['CC'].drop_duplicates().tolist()
                year_proy_2027 = 2027
                meses_2027 = [pd.Timestamp(year=2027, month=m, day=1) for m in range(1, 13)]
                meses_2027_str = [f"{m:02d}-2027" for m in range(1, 13)]

                filas_excel_t3 = []
                filas_resumen_t3 = []

                for cc in ccs_all_t3:
                    df_cc = df_base_t3[df_base_t3['CC'] == cc].copy()
                    nombre_cc = df_cc['Nombre Cliente'].iloc[0] if not df_cc.empty else "DESCONOCIDO"

                    df_melt_cc = pd.melt(
                        df_cc,
                        id_vars=['CC', 'Nombre Cliente', 'Variable'],
                        value_vars=month_cols_t3,
                        var_name='Mes',
                        value_name='Valor'
                    )
                    df_melt_cc['Valor'] = pd.to_numeric(df_melt_cc['Valor'], errors='coerce').fillna(0)

                    try:
                        df_melt_cc['Fecha'] = pd.to_datetime(df_melt_cc['Mes'], format='%m-%Y', errors='coerce')
                        mask_cc = df_melt_cc['Fecha'].isna()
                        if mask_cc.any():
                            df_melt_cc.loc[mask_cc, 'Fecha'] = pd.to_datetime(df_melt_cc.loc[mask_cc, 'Mes'], errors='coerce')
                    except Exception:
                        df_melt_cc['Fecha'] = pd.to_datetime(df_melt_cc['Mes'], errors='coerce')

                    fechas_cc = sorted(df_melt_cc['Fecha'].dropna().unique())
                    if not fechas_cc:
                        continue
                    year_base_cc = fechas_cc[0].year

                    mod_pivot_cc = df_melt_cc.pivot_table(
                        index='Variable', columns='Fecha', values='Valor', aggfunc='sum'
                    ).fillna(0)

                    var_venta_cc = next((v for v in mod_pivot_cc.index if is_venta(v)), 'Venta')

                    tiene_temp_cc, meses_temp_cc = detectar_temporalidad(mod_pivot_cc, var_venta_cc)
                    var_pct_cc, var_abs_cc = calcular_variacion_oct_dic(mod_pivot_cc, var_venta_cc)

                    col_dic_cc = next((c for c in mod_pivot_cc.columns if c.month == 12), None)
                    if col_dic_cc is not None:
                        val_fijo_dic_cc = float(mod_pivot_cc.at['Fijo', col_dic_cc]) if 'Fijo' in mod_pivot_cc.index else 0.0
                        val_manip_dic_cc = float(mod_pivot_cc.at['Manipulación', col_dic_cc]) if 'Manipulación' in mod_pivot_cc.index else 0.0
                        _vd = float(mod_pivot_cc.at[var_venta_cc, col_dic_cc]) if var_venta_cc in mod_pivot_cc.index and mod_pivot_cc.at[var_venta_cc, col_dic_cc] > 0 else 1.0
                        pct_costo_dic_cc = safe_div(float(mod_pivot_cc.at['Costo', col_dic_cc]) if 'Costo' in mod_pivot_cc.index else 0.0, _vd)
                        pct_var_dic_cc = safe_div(float(mod_pivot_cc.at['Variable', col_dic_cc]) if 'Variable' in mod_pivot_cc.index else 0.0, _vd)
                    else:
                        _vp = float(mod_pivot_cc.loc[var_venta_cc].mean()) if var_venta_cc in mod_pivot_cc.index and mod_pivot_cc.loc[var_venta_cc].mean() > 0 else 1.0
                        val_fijo_dic_cc = float(mod_pivot_cc.loc['Fijo'].mean()) if 'Fijo' in mod_pivot_cc.index else 0.0
                        val_manip_dic_cc = float(mod_pivot_cc.loc['Manipulación'].mean()) if 'Manipulación' in mod_pivot_cc.index else 0.0
                        pct_costo_dic_cc = safe_div(float(mod_pivot_cc.loc['Costo'].mean()) if 'Costo' in mod_pivot_cc.index else 0.0, _vp)
                        pct_var_dic_cc = safe_div(float(mod_pivot_cc.loc['Variable'].mean()) if 'Variable' in mod_pivot_cc.index else 0.0, _vp)

                    valores_cc = {v: {} for v in ['Venta', 'Costo', 'Manipulación', 'Fijo', 'Variable', 'Margen', 'Días Hábiles']}

                    for d in meses_2027:
                        valores_cc['Días Hábiles'][d] = get_dias_habiles(d.year, d.month)

                        venta_p = proyectar_venta(
                            mod_pivot=mod_pivot_cc,
                            var_venta=var_venta_cc,
                            mes_num=d.month,
                            year_base=year_base_cc,
                            year_proy=year_proy_2027,
                            tiene_temporalidad=tiene_temp_cc,
                            meses_con_temporalidad=meses_temp_cc,
                            variacion_pct_oct_dic=var_pct_cc,
                            variacion_abs_oct_dic=var_abs_cc,
                            aumento_extra_pct=0.0
                        )
                        valores_cc['Venta'][d] = venta_p

                        costos_tot = 0.0
                        for var_c in ['Costo', 'Manipulación', 'Fijo', 'Variable']:
                            val_p = proyectar_costos(
                                mod_pivot=mod_pivot_cc,
                                var=var_c,
                                venta_proyectada=venta_p,
                                mes_num=d.month,
                                tiene_temporalidad=tiene_temp_cc,
                                val_fijo_diciembre=val_fijo_dic_cc,
                                val_manipulacion_diciembre=val_manip_dic_cc,
                                pct_costo_diciembre=pct_costo_dic_cc,
                                pct_variable_diciembre=pct_var_dic_cc
                            )
                            valores_cc[var_c][d] = val_p
                            costos_tot += val_p

                        valores_cc['Margen'][d] = venta_p - costos_tot

                    # Filas para Excel (formato compatible con plantilla original)
                    for var in ['Venta', 'Costo', 'Manipulación', 'Fijo', 'Variable', 'Margen']:
                        fila_exc = {'CC': cc, 'Nombre Cliente': nombre_cc, 'Tipo Modelo': 'PROYECCIÓN 2027', 'Variable': var}
                        for d, ms in zip(meses_2027, meses_2027_str):
                            fila_exc[ms] = round(valores_cc[var][d], 2)
                        fila_exc['TOTAL'] = round(sum(valores_cc[var][d] for d in meses_2027), 2)
                        filas_excel_t3.append(fila_exc)

                    # Fila de días hábiles para Excel
                    fila_dias_exc = {'CC': cc, 'Nombre Cliente': nombre_cc, 'Tipo Modelo': 'PROYECCIÓN 2027', 'Variable': 'Días Hábiles'}
                    for d, ms in zip(meses_2027, meses_2027_str):
                        fila_dias_exc[ms] = int(valores_cc['Días Hábiles'][d])
                    fila_dias_exc['TOTAL'] = int(sum(valores_cc['Días Hábiles'][d] for d in meses_2027))
                    filas_excel_t3.append(fila_dias_exc)

                    # Fila resumen para display
                    total_v = sum(valores_cc['Venta'][d] for d in meses_2027)
                    total_c = sum(valores_cc['Costo'][d] for d in meses_2027)
                    total_m_val = sum(valores_cc['Manipulación'][d] for d in meses_2027)
                    total_f = sum(valores_cc['Fijo'][d] for d in meses_2027)
                    total_var = sum(valores_cc['Variable'][d] for d in meses_2027)
                    total_mg = sum(valores_cc['Margen'][d] for d in meses_2027)
                    filas_resumen_t3.append({
                        'CC': cc,
                        'Nombre Cliente': nombre_cc,
                        'Venta 2027': round(total_v, 0),
                        'Costo 2027': round(total_c, 0),
                        'Manipulación 2027': round(total_m_val, 0),
                        'Fijo 2027': round(total_f, 0),
                        'Variable 2027': round(total_var, 0),
                        'Margen 2027': round(total_mg, 0),
                        'MC %': f"{safe_div(total_mg, total_v):.1%}",
                        'Temporalidad': 'Sí' if tiene_temp_cc else 'No'
                    })

            if filas_excel_t3:
                df_proy_2027_excel = pd.DataFrame(filas_excel_t3)
                df_resumen_2027 = pd.DataFrame(filas_resumen_t3)

                st.success(f"Proyección 2027 lista para {len(ccs_all_t3)} centros de costo.")

                # ── Tabla resumen ejecutivo ──
                st.markdown("### 📊 Resumen Anual por Casino — 2027")
                st.dataframe(
                    df_resumen_2027.set_index('CC'),
                    use_container_width=True
                )

                # ── Detalle por CC (expandibles) ──
                st.markdown("### 📋 Detalle Mes a Mes por Casino")
                vars_detalle = ['Venta', 'Costo', 'Manipulación', 'Fijo', 'Variable', 'Margen', 'Días Hábiles']

                for cc in ccs_all_t3:
                    df_cc_det = df_proy_2027_excel[
                        (df_proy_2027_excel['CC'] == cc) &
                        (df_proy_2027_excel['Variable'].isin(vars_detalle))
                    ].copy()
                    if df_cc_det.empty:
                        continue
                    nombre_det = df_cc_det['Nombre Cliente'].iloc[0]
                    with st.expander(f"📍 {cc} — {nombre_det}"):
                        cols_show = ['Variable'] + meses_2027_str + ['TOTAL']
                        df_det_view = df_cc_det[cols_show].set_index('Variable')
                        df_det_view['Variable'] = pd.Categorical(
                            df_det_view.index, categories=vars_detalle, ordered=True
                        )
                        df_det_view = df_det_view.reindex(vars_detalle)
                        st.dataframe(df_det_view, use_container_width=True)

                # ── Exportar ──
                st.markdown("---")
                output_t3 = io.BytesIO()
                with pd.ExcelWriter(output_t3, engine='openpyxl') as writer:
                    # Hoja principal compatible con formato de modelación
                    df_proy_2027_excel.to_excel(writer, index=False, sheet_name='Modelación Proyectada 2027')
                    # Hoja de resumen ejecutivo
                    df_resumen_2027.to_excel(writer, index=False, sheet_name='Resumen Ejecutivo 2027')

                st.download_button(
                    label="📥 Descargar Modelación Proyectada 2027 (Excel)",
                    data=output_t3.getvalue(),
                    file_name="Modelacion_Proyectada_2027.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_dl_excel_tab3"
                )
            else:
                st.warning("No se encontraron datos procesables en la modelación seleccionada.")

        except Exception as e:
            st.error(f"Error al calcular proyección 2027: {e}")
            import traceback
            st.code(traceback.format_exc())