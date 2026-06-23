import streamlit as st
import pandas as pd
import numpy as np
import holidays
import io
import os
from datetime import date

st.set_page_config(page_title="Comparativa EVA vs Modelación", layout="wide", page_icon="📊")

MESES_ES_NOMBRES = {
    1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio',
    7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
}

@st.cache_data
def get_template_ucp_excel():
    """
    Plantilla UCP: CC transportados con COSTO y VENTA reales.
    Si el mismo CC repite en el mismo periodo, se suman COSTO y VENTA antes de calcular el %.
    El % Costo/Venta se calcula automáticamente (COSTO / VENTA).
    """
    df = pd.DataFrame({
        'CC':      ['201','201','201', '202','202','202', '205','205','205'],
        'COSTO':   [200000, 210000, 205000,
                    175000, 165000, 170000,
                    125000, 125000, 125000],
        'VENTA':   [500000, 520000, 510000,
                    450000, 430000, 440000,
                    320000, 310000, 315000],
        'MES/AÑO': ['01-2026','02-2026','03-2026',
                    '01-2026','02-2026','03-2026',
                    '01-2026','02-2026','03-2026'],
    })
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Plantilla UCP')
    return output.getvalue()


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

def get_col_referencia(mod_pivot, var_venta):
    """Último mes con datos reales de Venta Y Costo (prefiere Diciembre si tiene datos)."""
    if var_venta not in mod_pivot.index:
        return None
    cols_con_venta = [c for c in mod_pivot.columns if mod_pivot.at[var_venta, c] > 0]
    if not cols_con_venta:
        return None
    cols_con_costo = [c for c in cols_con_venta if 'Costo' in mod_pivot.index and mod_pivot.at['Costo', c] > 0]
    candidatos = cols_con_costo if cols_con_costo else cols_con_venta
    col_dic = next((c for c in candidatos if c.month == 12), None)
    return col_dic if col_dic is not None else max(candidatos)

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
                    aumento_extra_pct=0.0, meses_excluidos_temp=None):
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

    # Si el usuario excluyó este mes de la temporalidad, forzar rama estándar
    if meses_excluidos_temp and mes_num in meses_excluidos_temp:
        diferencia = 0.0

    if abs(diferencia) > 0.20:
        # Temporalidad: usar valor proporcional del mes de la modelación + 4% (inflación/reajuste año móvil)
        dias_proy = get_dias_habiles(year_proy, mes_num)
        venta_proyectada = v_base_diaria * dias_proy * 1.04
    else:
        # Sin temporalidad: venta diaria de octubre; variación Oct-Dic solo si es positiva
        if v_octubre_diaria != 0:
            if variacion_pct_oct_dic > 0:
                v_proy_pct = v_octubre_diaria * (1 + variacion_pct_oct_dic)
                v_proy_abs = v_octubre_diaria + variacion_abs_oct_dic
                venta_diaria_final = (v_proy_pct + v_proy_abs) / 2
            else:
                # Variación negativa o nula → se usa octubre sin modificar
                venta_diaria_final = v_octubre_diaria
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
        var_venta_key = next((v for v in mod_pivot.index if is_venta(v)), None)
        venta_mes_real = float(mod_pivot.at[var_venta_key, col_mes]) if (col_mes is not None and var_venta_key and var_venta_key in mod_pivot.index) else 0.0

        if col_mes is not None and var in mod_pivot.index and venta_mes_real > 0:
            val_mes = float(mod_pivot.at[var, col_mes])
            if es_fijo_o_manip:
                return val_mes if val_mes != 0 else (val_fijo_diciembre if 'fijo' in var_lower else val_manipulacion_diciembre)
            elif es_proporcional:
                pct_mes = safe_div(val_mes, venta_mes_real)
                return pct_mes * venta_proyectada
            else:
                return val_mes
        else:
            # Fallback al mes de referencia cuando el mes base no tiene datos reales
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
# PRE-SCAN: detección de temporalidades sin proyección completa
# ─────────────────────────────────────────────────────────────────────────────

def prescan_temporalidades(df):
    """Detecta temporalidades por CC sin realizar proyección. Devuelve {cc: {nombre, tiene, meses}}."""
    required = ['CC', 'Nombre Cliente', 'Tipo Modelo', 'Variable']
    # Excluir columna División si existe (no es una columna de mes)
    _div_col = next((c for c in df.columns if 'divis' in str(c).lower()), None)
    if _div_col and _div_col not in required:
        required = required + [_div_col]
    if any(c not in df.columns for c in ['CC', 'Nombre Cliente', 'Tipo Modelo', 'Variable']):
        return {}

    df = df.copy()
    month_cols = [c for c in df.columns if c not in required]
    df['CC'] = df['CC'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df['Nombre Cliente'] = df['Nombre Cliente'].astype(str).str.strip().str.upper()
    df['Tipo Modelo'] = df['Tipo Modelo'].astype(str).str.strip().str.upper().replace('MODELACIÓN', 'MODELACION')
    df['Variable'] = df['Variable'].astype(str).str.strip().replace({
        'Costo Alimento': 'Costo', 'Gasto Manipulación': 'Manipulación',
        'Gasto Fijo': 'Fijo', 'Gasto Variable': 'Variable', 'Margen de Contribución': 'Margen'
    })

    df_melt = pd.melt(df, id_vars=required, value_vars=month_cols, var_name='Mes', value_name='Valor')
    df_melt['Valor'] = pd.to_numeric(df_melt['Valor'], errors='coerce').fillna(0)
    try:
        df_melt['Fecha'] = pd.to_datetime(df_melt['Mes'], format='%m-%Y', errors='coerce')
        mask = df_melt['Fecha'].isna()
        if mask.any():
            df_melt.loc[mask, 'Fecha'] = pd.to_datetime(df_melt.loc[mask, 'Mes'], errors='coerce')
    except Exception:
        df_melt['Fecha'] = pd.to_datetime(df_melt['Mes'], errors='coerce')

    df_mod = df_melt[df_melt['Tipo Modelo'].isin(['MODELACION', 'INTERNO'])].copy()
    resultado = {}
    for cc in df_mod['CC'].drop_duplicates():
        mod_cc = df_mod[df_mod['CC'] == cc]
        nombre = mod_cc['Nombre Cliente'].iloc[0] if not mod_cc.empty else cc
        mod_pivot = mod_cc.pivot_table(index='Variable', columns='Fecha', values='Valor', aggfunc='sum').fillna(0)
        var_venta = next((v for v in mod_pivot.index if is_venta(v)), None)
        if var_venta:
            tiene, meses = detectar_temporalidad(mod_pivot, var_venta)
        else:
            tiene, meses = False, {}
        resultado[cc] = {'nombre': nombre, 'tiene': tiene, 'meses': meses}
    return resultado


# ─────────────────────────────────────────────────────────────────────────────
# PROCESAMIENTO PRINCIPAL (Tab 1: EVA vs Modelación)
# ─────────────────────────────────────────────────────────────────────────────

def process_data(df, overrides_temporalidad=None):
    required_cols = ['CC', 'Nombre Cliente', 'Tipo Modelo', 'Variable']
    for col in required_cols:
        if col not in df.columns:
            return None, None, f"Falta la columna requerida: {col}"
    # Excluir columna División si existe (no es una columna de mes)
    _div_col_pd = next((c for c in df.columns if 'divis' in str(c).lower()), None)
    if _div_col_pd and _div_col_pd not in required_cols:
        required_cols = required_cols + [_div_col_pd]

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

        # ── Aplicar overrides de temporalidad del usuario ──
        meses_excluidos_cc = overrides_temporalidad.get(cc, set()) if overrides_temporalidad else set()
        if meses_excluidos_cc:
            for m_exc in meses_excluidos_cc:
                meses_con_temporalidad.pop(m_exc, None)
            tiene_temporalidad = len(meses_con_temporalidad) > 0

        # ── Valores de referencia (último mes con datos reales, preferentemente Diciembre) ──
        col_diciembre = get_col_referencia(mod_pivot_raw, var_venta)
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
                usar_temp_mes = tiene_temporalidad and d.month not in meses_excluidos_cc
                if var_venta:
                    venta_proyectada = proyectar_venta(
                        mod_pivot=mod_pivot_raw,
                        var_venta=var_venta,
                        mes_num=d.month,
                        year_base=year_base_mod,
                        year_proy=d.year,
                        tiene_temporalidad=usar_temp_mes,
                        meses_con_temporalidad=meses_con_temporalidad,
                        variacion_pct_oct_dic=variacion_pct_oct_dic,
                        variacion_abs_oct_dic=variacion_abs_oct_dic,
                        aumento_extra_pct=0.0,
                        meses_excluidos_temp=meses_excluidos_cc
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
                        tiene_temporalidad=usar_temp_mes,
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

tab1, tab2, tab3, tab4 = st.tabs(["📊 Comparativa EVA vs Modelación", "🔍 Visualizador y Proyección de Modelación", "🗓️ Proyección 2027 (Todos los CC)", "🚛 Distribución Costos UCP"])

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
            # ── 1. Carga de archivos (cacheado en session_state, solo se reprocesa si cambian) ──
            files_key = f"{mod_seleccionada}_{uploaded_eva.name}_{uploaded_eva.size}"
            if st.session_state.get('tab1_files_key') != files_key:
                path_mod_t1 = os.path.join(REPO_DIR, mod_seleccionada)
                _df_mod = pd.read_excel(path_mod_t1)
                _df_eva = pd.read_excel(uploaded_eva)
                _df_combined = pd.concat([_df_mod, _df_eva], ignore_index=True)
                st.session_state['tab1_files_key'] = files_key
                st.session_state['tab1_df_combined'] = _df_combined
                st.session_state['tab1_temp_info'] = prescan_temporalidades(_df_combined)
                # Limpiar checkboxes y resultados cuando cambian los archivos
                for k in [k for k in st.session_state if k.startswith('temp_t1_')]:
                    del st.session_state[k]
                st.session_state.pop('tab1_results_key', None)
                st.session_state.pop('tab1_results', None)

            df_combined_t1 = st.session_state['tab1_df_combined']
            temp_info_t1 = st.session_state['tab1_temp_info']
            ccs_con_temp = {cc: info for cc, info in temp_info_t1.items() if info['tiene']}

            # ── 2. Leer valores actuales de checkboxes desde session_state (antes de renderizarlos) ──
            overrides_t1 = {}
            for cc, info in ccs_con_temp.items():
                meses_excl = set()
                for mes_num in info['meses']:
                    if not st.session_state.get(f"temp_t1_{cc}_{mes_num}", True):
                        meses_excl.add(mes_num)
                if meses_excl:
                    overrides_t1[cc] = meses_excl

            # ── 3. Procesar solo si cambió la combinación archivos+overrides ──
            overrides_key = str(sorted([(cc, sorted(m)) for cc, m in overrides_t1.items()]))
            results_key = f"{files_key}__{overrides_key}"
            if st.session_state.get('tab1_results_key') != results_key:
                with st.spinner("Procesando archivos e infiriendo estacionalidad..."):
                    _dm, _dr, _err = process_data(df_combined_t1, overrides_t1)
                st.session_state['tab1_results_key'] = results_key
                st.session_state['tab1_results'] = (_dm, _dr, _err)

            df_mensual, df_resumen, error = st.session_state['tab1_results']

            if error:
                st.error(error)
            else:
                # ── 4. Mostrar resultados ──
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

                    # ── Ajuste de temporalidad (solo si tiene temporalidad) ──
                    cc_this = str(row['CC'])
                    if cc_this in ccs_con_temp:
                        info_temp = ccs_con_temp[cc_this]
                        meses_ord_temp = sorted(info_temp['meses'].items())
                        st.caption("Desmarca los meses que no quieres tratar como temporales. La comparativa recalcula al instante.")
                        n_cols_t = min(len(meses_ord_temp), 4)
                        cols_t = st.columns(n_cols_t)
                        for i_t, (mes_num_t, pct_t) in enumerate(meses_ord_temp):
                            with cols_t[i_t % n_cols_t]:
                                st.checkbox(
                                    f"{MESES_ES_NOMBRES.get(mes_num_t, str(mes_num_t))} ({pct_t:+.1%})",
                                    value=True,
                                    key=f"temp_t1_{cc_this}_{mes_num_t}"
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
            excluded_cols += [c for c in df_mod_raw.columns if 'divis' in str(c).lower()]
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

            # ── Valores de referencia (último mes con datos reales) ──
            col_diciembre = get_col_referencia(mod_pivot, var_venta)
            if col_diciembre is not None:
                val_fijo_diciembre = float(mod_pivot.at['Fijo', col_diciembre]) if 'Fijo' in mod_pivot.index else 0.0
                val_manipulacion_diciembre = float(mod_pivot.at['Manipulación', col_diciembre]) if 'Manipulación' in mod_pivot.index else 0.0
                venta_diciembre = float(mod_pivot.at[var_venta, col_diciembre]) if var_venta in mod_pivot.index and mod_pivot.at[var_venta, col_diciembre] > 0 else 1.0
                pct_costo_diciembre = safe_div(float(mod_pivot.at['Costo', col_diciembre]) if 'Costo' in mod_pivot.index else 0.0, venta_diciembre)
                pct_variable_diciembre = safe_div(float(mod_pivot.at['Variable', col_diciembre]) if 'Variable' in mod_pivot.index else 0.0, venta_diciembre)
            else:
                val_fijo_diciembre = mod_pivot.loc['Fijo'].mean() if 'Fijo' in mod_pivot.index else 0.0
                val_manipulacion_diciembre = mod_pivot.loc['Manipulación'].mean() if 'Manipulación' in mod_pivot.index else 0.0
                venta_promedio = mod_pivot.loc[var_venta].mean() if var_venta in mod_pivot.index and mod_pivot.loc[var_venta].mean() > 0 else 1.0
                pct_costo_diciembre = safe_div(mod_pivot.loc['Costo'].mean() if 'Costo' in mod_pivot.index else 0.0, venta_promedio)
                pct_variable_diciembre = safe_div(mod_pivot.loc['Variable'].mean() if 'Variable' in mod_pivot.index else 0.0, venta_promedio)
            mes_ref_nombre = MESES_ES.get(col_diciembre.month, 'referencia') if col_diciembre is not None else 'promedio'

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
                    f"Costos referenciados a **{mes_ref_nombre}** — Fijo: **${val_fijo_diciembre:,.0f}**, "
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
                excluded_t3 += [c for c in df_base_t3.columns if 'divis' in str(c).lower()]
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

                    col_dic_cc = get_col_referencia(mod_pivot_cc, var_venta_cc)
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

                    # Filtrar CC sin datos (todos los totales financieros en cero)
                    total_v = sum(valores_cc['Venta'][d] for d in meses_2027)
                    total_c = sum(valores_cc['Costo'][d] for d in meses_2027)
                    total_m_val = sum(valores_cc['Manipulación'][d] for d in meses_2027)
                    total_f = sum(valores_cc['Fijo'][d] for d in meses_2027)
                    total_var = sum(valores_cc['Variable'][d] for d in meses_2027)
                    total_mg = sum(valores_cc['Margen'][d] for d in meses_2027)

                    if abs(total_v) + abs(total_c) + abs(total_m_val) + abs(total_f) + abs(total_var) + abs(total_mg) == 0:
                        continue

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

                ccs_con_datos = df_resumen_2027['CC'].tolist()
                st.success(f"Proyección 2027 lista — {len(ccs_con_datos)} casinos con datos.")

                # ── Tabla resumen ejecutivo ──
                st.markdown("### 📊 Resumen Anual por Casino — 2027")
                st.dataframe(
                    df_resumen_2027.set_index('CC'),
                    use_container_width=True
                )

                # ── Detalle por CC con buscador y límite de 10 ──
                st.markdown("### 📋 Detalle Mes a Mes por Casino")
                vars_detalle = ['Venta', 'Costo', 'Manipulación', 'Fijo', 'Variable', 'Margen', 'Días Hábiles']

                busqueda_t3 = st.text_input(
                    "Buscar casino por CC o Nombre:",
                    placeholder="Escribe el código CC o parte del nombre...",
                    key="busqueda_det_t3"
                )

                termino = busqueda_t3.strip().lower()
                if termino:
                    ccs_filtrados = [
                        cc for cc in ccs_con_datos
                        if termino in cc.lower() or termino in df_resumen_2027.loc[
                            df_resumen_2027['CC'] == cc, 'Nombre Cliente'
                        ].iloc[0].lower()
                    ]
                else:
                    ccs_filtrados = ccs_con_datos

                LIMITE_DETALLE = 10
                ccs_mostrar = ccs_filtrados[:LIMITE_DETALLE]

                if len(ccs_filtrados) > LIMITE_DETALLE:
                    st.info(
                        f"Mostrando {LIMITE_DETALLE} de {len(ccs_filtrados)} casinos. "
                        "Usa el buscador para filtrar por CC o nombre."
                    )
                elif len(ccs_filtrados) == 0:
                    st.warning("No se encontraron casinos que coincidan con la búsqueda.")

                for cc in ccs_mostrar:
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


# ─────────────────────────────────────────────────────────────────────────────
# TAB 4: DISTRIBUCIÓN COSTOS UCP — CASINOS TRANSPORTADOS
# ─────────────────────────────────────────────────────────────────────────────

with tab4:
    st.header("Distribución Costos UCP — Casinos Transportados")
    st.write(
        "Carga el archivo de costos UCP entregado por Planificación (ya distribuido por casino). "
        "Los meses con datos reales se usan directamente; los meses faltantes se proyectan aplicando "
        "el **% Costo/Venta** de referencia (preferentemente Diciembre) sobre la venta proyectada de la modelación."
    )

    col_t4a, col_t4b = st.columns(2)
    with col_t4a:
        st.subheader("1. Modelación Base")
        if archivos_repo:
            mod_sel_t4 = st.selectbox(
                "Elige la modelación:",
                options=["Seleccionar..."] + archivos_repo,
                key="mod_sel_tab4"
            )
        else:
            st.warning("Sube un archivo en la barra lateral.")
            mod_sel_t4 = "Seleccionar..."

    with col_t4b:
        st.subheader("2. Archivo UCP (Planificación)")
        st.download_button(
            label="📥 Descargar Plantilla UCP",
            data=get_template_ucp_excel(),
            file_name="Plantilla_UCP.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_dl_plantilla_ucp"
        )
        ucp_file_t4 = st.file_uploader(
            "Sube el archivo Excel UCP",
            type=["xlsx", "xls"],
            key="upload_ucp_t4"
        )

    st.markdown("---")

    if mod_sel_t4 != "Seleccionar..." and ucp_file_t4 is not None:
        try:
            # ── 1. Parsear archivo UCP (CC, COSTO, VENTA, MES/AÑO) ──
            df_ucp_t4 = pd.read_excel(ucp_file_t4)
            df_ucp_t4.columns = [
                str(c).strip().upper()
                    .replace(' ', '_').replace('/', '_')
                    .replace('Á','A').replace('É','E').replace('Í','I')
                    .replace('Ó','O').replace('Ú','U').replace('Ñ','N')
                for c in df_ucp_t4.columns
            ]
            rename_ucp = {}
            seen_u = set()
            for c in df_ucp_t4.columns:
                if c == 'CC' and 'CC' not in seen_u:
                    rename_ucp[c] = 'CC'; seen_u.add('CC')
                elif 'COSTO' in c and 'COSTO' not in seen_u:
                    rename_ucp[c] = 'COSTO'; seen_u.add('COSTO')
                elif 'VENTA' in c and 'VENTA' not in seen_u:
                    rename_ucp[c] = 'VENTA'; seen_u.add('VENTA')
                elif ('MES' in c or 'ANO' in c or 'PERIODO' in c or 'FECHA' in c) and 'MESANIO' not in seen_u:
                    rename_ucp[c] = 'MESANIO'; seen_u.add('MESANIO')
            df_ucp_t4 = df_ucp_t4.rename(columns=rename_ucp)

            for col in ['CC', 'COSTO', 'VENTA', 'MESANIO']:
                if col not in df_ucp_t4.columns:
                    st.error(f"El archivo UCP no tiene la columna '{col}'. Descarga la plantilla de referencia.")
                    st.stop()

            def _norm_cc(v):
                s = str(v).strip()
                try:
                    return str(int(float(s)))
                except Exception:
                    return s
            df_ucp_t4['CC']    = df_ucp_t4['CC'].apply(_norm_cc)
            df_ucp_t4['COSTO'] = pd.to_numeric(df_ucp_t4['COSTO'], errors='coerce').fillna(0)
            df_ucp_t4['VENTA'] = pd.to_numeric(df_ucp_t4['VENTA'], errors='coerce').fillna(0)

            fecha_p = None
            for fmt in ['%m-%Y', '%m/%Y', '%Y-%m-%d', '%Y-%m', '%d-%m-%Y', '%d/%m/%Y']:
                _p = pd.to_datetime(df_ucp_t4['MESANIO'], format=fmt, errors='coerce')
                if _p.notna().sum() >= len(df_ucp_t4) * 0.5:
                    fecha_p = _p; break
            if fecha_p is None:
                fecha_p = pd.to_datetime(df_ucp_t4['MESANIO'], errors='coerce')
            df_ucp_t4['FECHA'] = fecha_p
            df_ucp_t4 = df_ucp_t4.dropna(subset=['FECHA'])
            df_ucp_t4['FECHA'] = df_ucp_t4['FECHA'].dt.to_period('M').dt.to_timestamp()

            if df_ucp_t4.empty:
                st.error("No se pudo interpretar la columna MES/AÑO. Usa el formato MM-YYYY (ej: 01-2026).")
                st.stop()

            # ── 2. Agregar por (CC, mes): sumar COSTO y VENTA; derivar % ──
            year_ucp  = int(df_ucp_t4['FECHA'].dt.year.max())
            year_next = year_ucp + 1
            df_anio = df_ucp_t4[df_ucp_t4['FECHA'].dt.year == year_ucp].copy()
            meses_reales_ucp = sorted(df_anio['FECHA'].dt.month.unique().tolist())
            meses_proy_ucp   = sorted(set(range(1, 13)) - set(meses_reales_ucp))

            agg = (
                df_anio.groupby(['CC', 'FECHA'])
                .agg(COSTO_SUM=('COSTO', 'sum'), VENTA_SUM=('VENTA', 'sum'))
                .reset_index()
            )
            agg['MES_NUM'] = agg['FECHA'].dt.month
            agg['PCT']     = agg.apply(lambda r: safe_div(r['COSTO_SUM'], r['VENTA_SUM']), axis=1)

            ccs_trans = sorted(agg['CC'].unique().tolist())
            if not ccs_trans:
                st.warning("No se encontraron CCs en el archivo UCP.")
                st.stop()

            # Lookups rápidos por (cc, mes)
            costo_ucp  = {(r['CC'], r['MES_NUM']): r['COSTO_SUM'] for _, r in agg.iterrows()}
            venta_ucp  = {(r['CC'], r['MES_NUM']): r['VENTA_SUM'] for _, r in agg.iterrows()}
            pct_ucp    = {(r['CC'], r['MES_NUM']): r['PCT']       for _, r in agg.iterrows()}

            # ── 3. Cargar modelación: detectar División TRANSP., leer venta y costo 138 ──
            with st.spinner("Cargando modelación..."):
                path_mod_t4 = os.path.join(REPO_DIR, mod_sel_t4)
                df_mod_t4_all = pd.read_excel(path_mod_t4)
                df_mod_t4_all['CC'] = df_mod_t4_all['CC'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                df_mod_t4_all['Nombre Cliente'] = df_mod_t4_all['Nombre Cliente'].astype(str).str.strip().str.upper()
                df_mod_t4_all['Tipo Modelo'] = (
                    df_mod_t4_all['Tipo Modelo'].astype(str).str.strip().str.upper()
                    .replace('MODELACIÓN', 'MODELACION')
                )
                df_mod_t4_all = df_mod_t4_all[
                    df_mod_t4_all['Tipo Modelo'].isin(['MODELACION', 'INTERNO'])
                ].copy()
                df_mod_t4_all['Variable'] = df_mod_t4_all['Variable'].astype(str).str.strip().replace({
                    'Costo Alimento': 'Costo', 'Gasto Manipulación': 'Manipulación',
                    'Gasto Fijo': 'Fijo', 'Gasto Variable': 'Variable',
                    'Margen de Contribución': 'Margen'
                })

                # Detectar columna División y excluirla de las columnas de mes
                div_col = next((c for c in df_mod_t4_all.columns if 'divis' in str(c).lower()), None)
                excl_t4 = ['CC', 'Nombre Cliente', 'Tipo Modelo', 'Variable']
                if div_col:
                    excl_t4 = excl_t4 + [div_col]
                    df_mod_t4_all[div_col] = df_mod_t4_all[div_col].astype(str).str.strip().str.upper()
                mcols_t4 = [c for c in df_mod_t4_all.columns if c not in excl_t4]

                # Filtrar TRANSP.
                if div_col:
                    df_transp = df_mod_t4_all[df_mod_t4_all[div_col].str.contains('TRANSP', na=False)].copy()
                    ccs_trans = sorted(df_transp['CC'].unique().tolist())
                else:
                    df_transp = df_mod_t4_all.copy()
                    ccs_trans = sorted(df_mod_t4_all['CC'].unique().tolist())
                    st.info("No se encontró la columna 'División' en la modelación. Se usan todos los CCs.")

                ccs_ucp_no_mod = [cc for cc in sorted(agg['CC'].unique()) if cc not in ccs_trans]
                if ccs_ucp_no_mod:
                    st.warning(f"CCs del Excel UCP no presentes como TRANSP. en la modelación: {', '.join(ccs_ucp_no_mod)}")

                # Nombres por CC
                nombres_cc_t4 = (
                    df_transp.drop_duplicates('CC').set_index('CC')['Nombre Cliente'].to_dict()
                )

                # Venta por (CC, mes) — leer directamente de la modelación, sin proyecciones externas
                df_venta_t = df_transp[df_transp['Variable'].apply(is_venta)].copy()
                df_mv = pd.melt(df_venta_t, id_vars=['CC'],
                                value_vars=mcols_t4, var_name='Mes', value_name='Venta')
                df_mv['Venta'] = pd.to_numeric(df_mv['Venta'], errors='coerce').fillna(0)
                df_mv['Fecha'] = pd.to_datetime(df_mv['Mes'], format='%m-%Y', errors='coerce')
                _na = df_mv['Fecha'].isna()
                if _na.any():
                    df_mv.loc[_na, 'Fecha'] = pd.to_datetime(df_mv.loc[_na, 'Mes'], errors='coerce')
                df_mv = df_mv.dropna(subset=['Fecha'])
                df_mv = df_mv[df_mv['Fecha'].dt.year == year_ucp]
                df_mv['MES_NUM'] = df_mv['Fecha'].dt.month
                # Suma por (CC, mes) para colapsar duplicados
                venta_cc_mes = (
                    df_mv.groupby(['CC', 'MES_NUM'])['Venta'].sum().to_dict()
                )  # {(cc, mes_num): venta}

                # Venta año siguiente — misma lógica de proyección del Tab 3 (proyectar_venta)
                venta_cc_mes_next = {}
                meses_next_dates = [pd.Timestamp(year=year_next, month=m, day=1) for m in range(1, 13)]
                for cc_n in ccs_trans:
                    df_cc_n = df_transp[df_transp['CC'] == cc_n].copy()
                    if df_cc_n.empty:
                        continue
                    df_melt_n = pd.melt(
                        df_cc_n, id_vars=['CC', 'Variable'],
                        value_vars=mcols_t4, var_name='Mes', value_name='Valor'
                    )
                    df_melt_n['Valor'] = pd.to_numeric(df_melt_n['Valor'], errors='coerce').fillna(0)
                    df_melt_n['Fecha'] = pd.to_datetime(df_melt_n['Mes'], format='%m-%Y', errors='coerce')
                    _na_n = df_melt_n['Fecha'].isna()
                    if _na_n.any():
                        df_melt_n.loc[_na_n, 'Fecha'] = pd.to_datetime(
                            df_melt_n.loc[_na_n, 'Mes'], errors='coerce')
                    df_melt_n = df_melt_n.dropna(subset=['Fecha'])
                    if df_melt_n.empty:
                        continue
                    year_base_n = int(df_melt_n['Fecha'].dt.year.min())
                    mod_pivot_n = df_melt_n.pivot_table(
                        index='Variable', columns='Fecha', values='Valor', aggfunc='sum'
                    ).fillna(0)
                    var_venta_n = next((v for v in mod_pivot_n.index if is_venta(v)), None)
                    if var_venta_n is None:
                        continue
                    tiene_temp_n, meses_temp_n = detectar_temporalidad(mod_pivot_n, var_venta_n)
                    var_pct_n, var_abs_n = calcular_variacion_oct_dic(mod_pivot_n, var_venta_n)
                    for d in meses_next_dates:
                        venta_cc_mes_next[(cc_n, d.month)] = proyectar_venta(
                            mod_pivot=mod_pivot_n,
                            var_venta=var_venta_n,
                            mes_num=d.month,
                            year_base=year_base_n,
                            year_proy=year_next,
                            tiene_temporalidad=tiene_temp_n,
                            meses_con_temporalidad=meses_temp_n,
                            variacion_pct_oct_dic=var_pct_n,
                            variacion_abs_oct_dic=var_abs_n,
                            aumento_extra_pct=0.0
                        )

                # Costo CC 138 desde la modelación
                CC_138 = '138'
                df_cc138 = df_mod_t4_all[df_mod_t4_all['CC'] == CC_138].copy()
                costo_138_mes = {}
                costo_138_mes_next = {}
                if not df_cc138.empty:
                    var_c138 = next((v for v in df_cc138['Variable'].unique() if 'costo' in v.lower()), None)
                    if var_c138:
                        df_m138 = pd.melt(
                            df_cc138[df_cc138['Variable'] == var_c138],
                            id_vars=['CC'], value_vars=mcols_t4,
                            var_name='Mes', value_name='Costo'
                        )
                        df_m138['Costo'] = pd.to_numeric(df_m138['Costo'], errors='coerce').fillna(0)
                        df_m138['Fecha'] = pd.to_datetime(df_m138['Mes'], format='%m-%Y', errors='coerce')
                        _na138 = df_m138['Fecha'].isna()
                        if _na138.any():
                            df_m138.loc[_na138, 'Fecha'] = pd.to_datetime(df_m138.loc[_na138, 'Mes'], errors='coerce')
                        df_m138 = df_m138.dropna(subset=['Fecha'])
                        df_m138 = df_m138[df_m138['Fecha'].dt.year == year_ucp]
                        df_m138['MES_NUM'] = df_m138['Fecha'].dt.month
                        costo_138_mes = (
                            df_m138.groupby('MES_NUM')['Costo'].sum()
                            .apply(abs).to_dict()
                        )
                        # Costo 138 año siguiente — proyectar el Costo como variable principal
                        # (CC 138 no tiene Venta; se usa proyectar_venta sobre la variable Costo)
                        df_melt_138 = pd.melt(
                            df_cc138[df_cc138['Variable'] == var_c138],
                            id_vars=['CC', 'Variable'],
                            value_vars=mcols_t4, var_name='Mes', value_name='Valor'
                        )
                        df_melt_138['Valor'] = pd.to_numeric(df_melt_138['Valor'], errors='coerce').fillna(0)
                        df_melt_138['Fecha'] = pd.to_datetime(df_melt_138['Mes'], format='%m-%Y', errors='coerce')
                        _na138b = df_melt_138['Fecha'].isna()
                        if _na138b.any():
                            df_melt_138.loc[_na138b, 'Fecha'] = pd.to_datetime(
                                df_melt_138.loc[_na138b, 'Mes'], errors='coerce')
                        df_melt_138 = df_melt_138.dropna(subset=['Fecha'])
                        if not df_melt_138.empty:
                            year_base_138 = int(df_melt_138['Fecha'].dt.year.min())
                            mod_piv_138 = df_melt_138.pivot_table(
                                index='Variable', columns='Fecha', values='Valor', aggfunc='sum'
                            ).fillna(0)
                            if var_c138 in mod_piv_138.index:
                                tiene_tmp_138, meses_tmp_138 = detectar_temporalidad(
                                    mod_piv_138, var_c138)
                                var_pct_138, var_abs_138 = calcular_variacion_oct_dic(
                                    mod_piv_138, var_c138)
                                for d in meses_next_dates:
                                    costo_138_mes_next[d.month] = abs(proyectar_venta(
                                        mod_pivot=mod_piv_138,
                                        var_venta=var_c138,
                                        mes_num=d.month,
                                        year_base=year_base_138,
                                        year_proy=year_next,
                                        tiene_temporalidad=tiene_tmp_138,
                                        meses_con_temporalidad=meses_tmp_138,
                                        variacion_pct_oct_dic=var_pct_138,
                                        variacion_abs_oct_dic=var_abs_138,
                                        aumento_extra_pct=0.0
                                    ))

                # Hoja Refacturación: CC × mes → monto a restar de la venta antes de calcular costo UCP
                refac_cc_mes = {}
                refac_cc_mes_next = {}
                try:
                    xl_file = pd.ExcelFile(path_mod_t4)
                    sheet_refac = next(
                        (s for s in xl_file.sheet_names if 'refac' in s.lower()), None
                    )
                    if sheet_refac:
                        df_refac = pd.read_excel(path_mod_t4, sheet_name=sheet_refac)
                        df_refac.columns = [str(c).strip() for c in df_refac.columns]
                        cc_col_r = next(
                            (c for c in df_refac.columns if c.upper() == 'CC'),
                            df_refac.columns[0]
                        )
                        excl_r = [cc_col_r]
                        cas_col_r = next(
                            (c for c in df_refac.columns
                             if 'casino' in c.lower() or 'nombre' in c.lower()),
                            None
                        )
                        if cas_col_r:
                            excl_r.append(cas_col_r)
                        rcols = [c for c in df_refac.columns if c not in excl_r]
                        df_rm = pd.melt(
                            df_refac, id_vars=[cc_col_r],
                            value_vars=rcols, var_name='Mes', value_name='Refac'
                        )
                        df_rm['Refac'] = pd.to_numeric(df_rm['Refac'], errors='coerce').fillna(0)
                        df_rm['Fecha'] = pd.to_datetime(df_rm['Mes'], format='%m-%Y', errors='coerce')
                        _na_r = df_rm['Fecha'].isna()
                        if _na_r.any():
                            df_rm.loc[_na_r, 'Fecha'] = pd.to_datetime(
                                df_rm.loc[_na_r, 'Mes'], errors='coerce'
                            )
                        df_rm = df_rm.dropna(subset=['Fecha'])
                        df_rm['CC'] = df_rm[cc_col_r].apply(_norm_cc)
                        df_rm['MES_NUM'] = df_rm['Fecha'].dt.month

                        def _fill_refac_yr(df_yr):
                            result = {}
                            if df_yr.empty:
                                return result
                            base = df_yr.groupby(['CC', 'MES_NUM'])['Refac'].sum()
                            for cc_r in base.index.get_level_values('CC').unique():
                                meses_dato = base[cc_r].to_dict()
                                ultimo_m   = max(meses_dato.keys())
                                val_ult    = meses_dato[ultimo_m]
                                for m in range(1, 13):
                                    result[(cc_r, m)] = meses_dato.get(m, val_ult)
                            return result

                        refac_cc_mes = _fill_refac_yr(
                            df_rm[df_rm['Fecha'].dt.year == year_ucp]
                        )
                        _rn = _fill_refac_yr(df_rm[df_rm['Fecha'].dt.year == year_next])
                        if _rn:
                            refac_cc_mes_next = _rn
                        else:
                            # Arrastre: último mes de 2026 aplanado para todos los meses de 2027
                            _ccs_r = {cc for (cc, _) in refac_cc_mes}
                            refac_cc_mes_next = {
                                (cc_r, m): refac_cc_mes.get((cc_r, 12),
                                    refac_cc_mes.get((cc_r,
                                        max((mm for (c, mm) in refac_cc_mes if c == cc_r), default=1)
                                    ), 0.0))
                                for cc_r in _ccs_r
                                for m in range(1, 13)
                            }
                except Exception:
                    pass  # sin hoja Refacturación → refac = 0 para todos

            # ── 4. % de referencia para meses sin dato en el Excel UCP ──
            mes_ref = 12 if 12 in meses_reales_ucp else (max(meses_reales_ucp) if meses_reales_ucp else None)
            nombre_mes_ref = MESES_ES_NOMBRES.get(mes_ref, str(mes_ref)) if mes_ref else '—'

            pct_ref_cc = {cc_t: pct_ucp.get((cc_t, mes_ref), 0.0) for cc_t in ccs_trans} if mes_ref else {}

            if mes_ref:
                venta_tot_ref = sum(venta_cc_mes.get((cc_t, mes_ref), 0.0) for cc_t in ccs_trans)
                refac_tot_ref = sum(refac_cc_mes.get((cc_t, mes_ref), 0.0) for cc_t in ccs_trans)
                venta_net_ref = venta_tot_ref - refac_tot_ref
                pct_138_ref   = safe_div(costo_138_mes.get(mes_ref, 0.0), venta_net_ref)
            else:
                pct_138_ref = 0.0

            # ── 5. Construir tablas de resultados ──
            resumen_t4 = []
            detalle_t4 = []

            for mes_n in range(1, 13):
                es_proy     = mes_n not in meses_reales_ucp
                venta_tot_m = 0.0
                refac_tot_m = 0.0
                vnet_tot_m  = 0.0
                costo_tot_m = 0.0
                filas_mes   = []

                for cc_t in ccs_trans:
                    v_cc   = venta_cc_mes.get((cc_t, mes_n), 0.0)
                    refac  = refac_cc_mes.get((cc_t, mes_n), 0.0)
                    v_net  = v_cc - refac
                    pct_cc = pct_ucp.get((cc_t, mes_n), pct_ref_cc.get(cc_t, 0.0)) if not es_proy else pct_ref_cc.get(cc_t, 0.0)
                    c_cc   = pct_cc * v_net
                    venta_tot_m += v_cc
                    refac_tot_m += refac
                    vnet_tot_m  += v_net
                    costo_tot_m += c_cc
                    filas_mes.append({
                        'MES/AÑO':        f"{mes_n:02d}-{year_ucp}",
                        'Tipo':           'Proyectado' if es_proy else 'Real',
                        'CC':             cc_t,
                        'Nombre Casino':  nombres_cc_t4.get(cc_t, cc_t),
                        'Venta':          v_cc,
                        'Refacturación':  refac,
                        'Venta Neta':     v_net,
                        'Costo UCP':      c_cc,
                        '% Costo/Venta':  safe_div(c_cc, v_net),
                    })

                costo_138_m  = costo_138_mes.get(mes_n, pct_138_ref * vnet_tot_m)
                diferencial_m = costo_138_m - costo_tot_m

                # Distribuir diferencial proporcional al Costo UCP de cada CC sobre el total del mes
                for fila in filas_mes:
                    cargo = safe_div(fila['Costo UCP'], costo_tot_m) * diferencial_m
                    fila['Cargo Diferencial']    = cargo
                    fila['Costo Final']          = fila['Costo UCP'] + cargo
                    fila['% Costo Final/Venta']  = safe_div(fila['Costo Final'], fila['Venta Neta'])
                detalle_t4.extend(filas_mes)

                resumen_t4.append({
                    'MES/AÑO':                    f"{mes_n:02d}-{year_ucp}",
                    'Tipo':                       'Proyectado' if es_proy else 'Real',
                    'Venta Total Transportados':  venta_tot_m,
                    'Refacturación Total':         refac_tot_m,
                    'Venta Neta Transportados':   vnet_tot_m,
                    'Costo Total UCP':            costo_tot_m,
                    '% Costo/Venta (UCP)':        safe_div(costo_tot_m, vnet_tot_m),
                    'Costo CC 138 (Modelación)':  costo_138_m,
                    '% Costo CC138/Venta':        safe_div(costo_138_m, vnet_tot_m),
                    'Diferencial':                diferencial_m,
                })

            # ── 5b. Proyección año siguiente ──
            resumen_next = []
            detalle_next = []

            for mes_n in range(1, 13):
                venta_tot_m = 0.0
                refac_tot_m = 0.0
                vnet_tot_m  = 0.0
                costo_tot_m = 0.0
                filas_mes   = []

                for cc_t in ccs_trans:
                    v_cc   = venta_cc_mes_next.get((cc_t, mes_n), 0.0)
                    refac  = refac_cc_mes_next.get((cc_t, mes_n), 0.0)
                    v_net  = v_cc - refac
                    pct_cc = pct_ref_cc.get(cc_t, 0.0)
                    c_cc   = pct_cc * v_net
                    venta_tot_m += v_cc
                    refac_tot_m += refac
                    vnet_tot_m  += v_net
                    costo_tot_m += c_cc
                    filas_mes.append({
                        'MES/AÑO':        f"{mes_n:02d}-{year_next}",
                        'Tipo':           'Proyectado',
                        'CC':             cc_t,
                        'Nombre Casino':  nombres_cc_t4.get(cc_t, cc_t),
                        'Venta':          v_cc,
                        'Refacturación':  refac,
                        'Venta Neta':     v_net,
                        'Costo UCP':      c_cc,
                        '% Costo/Venta':  safe_div(c_cc, v_net),
                    })

                costo_138_m  = costo_138_mes_next.get(mes_n, pct_138_ref * vnet_tot_m)
                diferencial_m = costo_138_m - costo_tot_m

                for fila in filas_mes:
                    cargo = safe_div(fila['Costo UCP'], costo_tot_m) * diferencial_m
                    fila['Cargo Diferencial']   = cargo
                    fila['Costo Final']         = fila['Costo UCP'] + cargo
                    fila['% Costo Final/Venta'] = safe_div(fila['Costo Final'], fila['Venta Neta'])
                detalle_next.extend(filas_mes)

                resumen_next.append({
                    'MES/AÑO':                    f"{mes_n:02d}-{year_next}",
                    'Tipo':                       'Proyectado',
                    'Venta Total Transportados':  venta_tot_m,
                    'Refacturación Total':         refac_tot_m,
                    'Venta Neta Transportados':   vnet_tot_m,
                    'Costo Total UCP':            costo_tot_m,
                    '% Costo/Venta (UCP)':        safe_div(costo_tot_m, vnet_tot_m),
                    'Costo CC 138 (Modelación)':  costo_138_m,
                    '% Costo CC138/Venta':        safe_div(costo_138_m, vnet_tot_m),
                    'Diferencial':                diferencial_m,
                })

            # ── 6. Mostrar resultados ──
            meses_reales_138 = sorted(costo_138_mes.keys())
            st.markdown(
                f"### 📅 Año UCP: **{year_ucp}** &nbsp;|&nbsp; "
                f"Casinos Transportados: **{len(ccs_trans)}** &nbsp;|&nbsp; "
                f"Referencia proyección: **{nombre_mes_ref}**"
            )
            st.caption(
                f"🟢 Meses reales (UCP): "
                f"{', '.join(MESES_ES_NOMBRES.get(m,'') for m in meses_reales_ucp) or 'Ninguno'}  |  "
                f"📊 CC 138 con dato en modelación: "
                f"{', '.join(MESES_ES_NOMBRES.get(m,'') for m in meses_reales_138) or 'Ninguno'}"
            )
            # Diagnóstico: CCs por mes
            cc_por_mes_ucp = agg.groupby('MES_NUM')['CC'].apply(
                lambda s: set(s.tolist()) & set(ccs_trans)
            ).to_dict()
            filas_diag = []
            for mes_n in range(1, 13):
                reales = len(cc_por_mes_ucp.get(mes_n, set()))
                proy   = len(ccs_trans) - reales
                filas_diag.append({
                    'Mes': MESES_ES_NOMBRES[mes_n],
                    'CCs TRANSP. en modelación': len(ccs_trans),
                    'Con dato % en Excel UCP': reales,
                    'Sin dato (se proyecta)': proy,
                })
            with st.expander("🔍 CCs considerados por mes"):
                st.dataframe(pd.DataFrame(filas_diag).set_index('Mes'), use_container_width=True)

            # Tabla mensual consolidada
            st.markdown("### 📊 Resumen Mensual")
            df_res = pd.DataFrame(resumen_t4).copy()
            df_disp = df_res.copy()
            for col in ['Venta Total Transportados', 'Refacturación Total', 'Venta Neta Transportados',
                        'Costo Total UCP', 'Costo CC 138 (Modelación)', 'Diferencial']:
                if col in df_disp.columns:
                    df_disp[col] = df_disp[col].apply(lambda x: f"${x:,.0f}")
            for col in ['% Costo/Venta (UCP)', '% Costo CC138/Venta']:
                df_disp[col] = df_disp[col].apply(lambda x: f"{x:.2%}")
            st.dataframe(df_disp.set_index('MES/AÑO'), use_container_width=True)

            # Detalle por casino
            st.markdown("### 📋 Detalle por Casino Transportado")
            for cc_t in ccs_trans:
                nombre_t = nombres_cc_t4.get(cc_t, cc_t)
                filas_cc = [r for r in detalle_t4 if r['CC'] == cc_t]
                if not filas_cc:
                    continue
                with st.expander(f"📍 {cc_t} — {nombre_t}"):
                    df_det = pd.DataFrame(filas_cc)[
                        ['MES/AÑO', 'Tipo', 'Venta', 'Refacturación', 'Venta Neta',
                         'Costo UCP', '% Costo/Venta', 'Cargo Diferencial', 'Costo Final', '% Costo Final/Venta']
                    ].copy().set_index('MES/AÑO')
                    for col in ['Venta', 'Refacturación', 'Venta Neta', 'Costo UCP', 'Cargo Diferencial', 'Costo Final']:
                        df_det[col] = df_det[col].apply(lambda x: f"${x:,.0f}")
                    for col in ['% Costo/Venta', '% Costo Final/Venta']:
                        df_det[col] = df_det[col].apply(lambda x: f"{x:.2%}")
                    st.dataframe(df_det, use_container_width=True)
                    v_a    = sum(r['Venta']              for r in filas_cc)
                    rfac_a = sum(r['Refacturación']      for r in filas_cc)
                    vnet_a = sum(r['Venta Neta']         for r in filas_cc)
                    c_a    = sum(r['Costo UCP']          for r in filas_cc)
                    cd_a   = sum(r['Cargo Diferencial']  for r in filas_cc)
                    cf_a   = sum(r['Costo Final']        for r in filas_cc)
                    pct_r  = pct_ref_cc.get(cc_t, 0.0)
                    st.caption(
                        f"Anual — Venta: **${v_a:,.0f}** | Refacturación: **${rfac_a:,.0f}** | "
                        f"Venta Neta: **${vnet_a:,.0f}** | Costo UCP: **${c_a:,.0f}** | "
                        f"Cargo Diferencial: **${cd_a:,.0f}** | Costo Final: **${cf_a:,.0f}** | "
                        f"% Costo Final/Venta: **{safe_div(cf_a, vnet_a):.2%}** | "
                        f"% Ref. proyección ({nombre_mes_ref}): **{pct_r:.2%}**"
                    )

            # Resumen anual
            st.markdown("### 💰 Resumen Anual por Casino Transportado")
            filas_anual = []
            for cc_t in ccs_trans:
                filas_cc = [r for r in detalle_t4 if r['CC'] == cc_t]
                v_a    = sum(r['Venta']             for r in filas_cc)
                rfac_a = sum(r['Refacturación']     for r in filas_cc)
                vnet_a = sum(r['Venta Neta']        for r in filas_cc)
                c_a    = sum(r['Costo UCP']         for r in filas_cc)
                cd_a   = sum(r['Cargo Diferencial'] for r in filas_cc)
                cf_a   = sum(r['Costo Final']       for r in filas_cc)
                filas_anual.append({
                    'CC':                       cc_t,
                    'Nombre Casino':            nombres_cc_t4.get(cc_t, cc_t),
                    'Venta Anual':              v_a,
                    'Refacturación':            rfac_a,
                    'Venta Neta Anual':         vnet_a,
                    'Costo UCP Anual':          c_a,
                    'Cargo Diferencial Anual':  cd_a,
                    'Costo Final Anual':        cf_a,
                    '% Costo Final s/Venta':    safe_div(cf_a, vnet_a),
                    '% Ref. Proyección':        pct_ref_cc.get(cc_t, 0.0),
                })
            df_anual  = pd.DataFrame(filas_anual)
            tot_v     = df_anual['Venta Anual'].sum()
            tot_rfac  = df_anual['Refacturación'].sum()
            tot_vnet  = df_anual['Venta Neta Anual'].sum()
            tot_c     = df_anual['Costo UCP Anual'].sum()
            tot_cd    = df_anual['Cargo Diferencial Anual'].sum()
            tot_cf    = df_anual['Costo Final Anual'].sum()
            df_anual = pd.concat([df_anual, pd.DataFrame([{
                'CC': 'TOTAL', 'Nombre Casino': '—',
                'Venta Anual': tot_v, 'Refacturación': tot_rfac,
                'Venta Neta Anual': tot_vnet, 'Costo UCP Anual': tot_c,
                'Cargo Diferencial Anual': tot_cd, 'Costo Final Anual': tot_cf,
                '% Costo Final s/Venta': safe_div(tot_cf, tot_vnet), '% Ref. Proyección': '',
            }])], ignore_index=True)
            df_anual_d = df_anual.copy().set_index('CC')
            for col in ['Venta Anual', 'Refacturación', 'Venta Neta Anual',
                        'Costo UCP Anual', 'Cargo Diferencial Anual', 'Costo Final Anual']:
                df_anual_d[col] = df_anual_d[col].apply(lambda x: f"${x:,.0f}" if isinstance(x, (int,float)) else x)
            for col in ['% Costo Final s/Venta', '% Ref. Proyección']:
                df_anual_d[col] = df_anual_d[col].apply(
                    lambda x: f"{x:.2%}" if isinstance(x, float) else x)
            st.dataframe(df_anual_d, use_container_width=True)

            # ── 7. Proyección año siguiente ──
            st.markdown(f"---")
            st.markdown(f"## 📅 Proyección {year_next} (Arrastre)")
            st.caption(f"Todos los meses usan el % de referencia ({nombre_mes_ref} {year_ucp}). "
                       f"Venta y Costo 138 leídos desde la modelación para {year_next} si existen columnas; "
                       f"si no, proyectados con % de referencia.")

            df_res_nx = pd.DataFrame(resumen_next).copy()
            df_disp_nx = df_res_nx.copy()
            for col in ['Venta Total Transportados', 'Refacturación Total', 'Venta Neta Transportados',
                        'Costo Total UCP', 'Costo CC 138 (Modelación)', 'Diferencial']:
                if col in df_disp_nx.columns:
                    df_disp_nx[col] = df_disp_nx[col].apply(lambda x: f"${x:,.0f}")
            for col in ['% Costo/Venta (UCP)', '% Costo CC138/Venta']:
                df_disp_nx[col] = df_disp_nx[col].apply(lambda x: f"{x:.2%}")
            st.markdown(f"### 📊 Resumen Mensual {year_next}")
            st.dataframe(df_disp_nx.set_index('MES/AÑO'), use_container_width=True)

            st.markdown(f"### 📋 Detalle por Casino Transportado {year_next}")
            for cc_t in ccs_trans:
                nombre_t = nombres_cc_t4.get(cc_t, cc_t)
                filas_cc = [r for r in detalle_next if r['CC'] == cc_t]
                if not filas_cc:
                    continue
                with st.expander(f"📍 {cc_t} — {nombre_t}"):
                    df_det_nx = pd.DataFrame(filas_cc)[
                        ['MES/AÑO', 'Tipo', 'Venta', 'Refacturación', 'Venta Neta',
                         'Costo UCP', '% Costo/Venta', 'Cargo Diferencial', 'Costo Final', '% Costo Final/Venta']
                    ].copy().set_index('MES/AÑO')
                    for col in ['Venta', 'Refacturación', 'Venta Neta', 'Costo UCP', 'Cargo Diferencial', 'Costo Final']:
                        df_det_nx[col] = df_det_nx[col].apply(lambda x: f"${x:,.0f}")
                    for col in ['% Costo/Venta', '% Costo Final/Venta']:
                        df_det_nx[col] = df_det_nx[col].apply(lambda x: f"{x:.2%}")
                    st.dataframe(df_det_nx, use_container_width=True)

            st.markdown(f"### 💰 Resumen Anual {year_next}")
            filas_anual_nx = []
            for cc_t in ccs_trans:
                filas_cc = [r for r in detalle_next if r['CC'] == cc_t]
                v_a    = sum(r['Venta']             for r in filas_cc)
                rfac_a = sum(r['Refacturación']     for r in filas_cc)
                vnet_a = sum(r['Venta Neta']        for r in filas_cc)
                c_a    = sum(r['Costo UCP']         for r in filas_cc)
                cd_a   = sum(r['Cargo Diferencial'] for r in filas_cc)
                cf_a   = sum(r['Costo Final']       for r in filas_cc)
                filas_anual_nx.append({
                    'CC':                      cc_t,
                    'Nombre Casino':           nombres_cc_t4.get(cc_t, cc_t),
                    'Venta Anual':             v_a,
                    'Refacturación':           rfac_a,
                    'Venta Neta Anual':        vnet_a,
                    'Costo UCP Anual':         c_a,
                    'Cargo Diferencial Anual': cd_a,
                    'Costo Final Anual':       cf_a,
                    '% Costo Final s/Venta':   safe_div(cf_a, vnet_a),
                })
            df_anual_nx = pd.DataFrame(filas_anual_nx)
            tot_v_nx  = df_anual_nx['Venta Anual'].sum()
            tot_rf_nx = df_anual_nx['Refacturación'].sum()
            tot_vn_nx = df_anual_nx['Venta Neta Anual'].sum()
            tot_c_nx  = df_anual_nx['Costo UCP Anual'].sum()
            tot_cd_nx = df_anual_nx['Cargo Diferencial Anual'].sum()
            tot_cf_nx = df_anual_nx['Costo Final Anual'].sum()
            df_anual_nx = pd.concat([df_anual_nx, pd.DataFrame([{
                'CC': 'TOTAL', 'Nombre Casino': '—',
                'Venta Anual': tot_v_nx, 'Refacturación': tot_rf_nx,
                'Venta Neta Anual': tot_vn_nx, 'Costo UCP Anual': tot_c_nx,
                'Cargo Diferencial Anual': tot_cd_nx, 'Costo Final Anual': tot_cf_nx,
                '% Costo Final s/Venta': safe_div(tot_cf_nx, tot_vn_nx),
            }])], ignore_index=True)
            df_anual_nx_d = df_anual_nx.copy().set_index('CC')
            for col in ['Venta Anual', 'Refacturación', 'Venta Neta Anual',
                        'Costo UCP Anual', 'Cargo Diferencial Anual', 'Costo Final Anual']:
                df_anual_nx_d[col] = df_anual_nx_d[col].apply(
                    lambda x: f"${x:,.0f}" if isinstance(x, (int, float)) else x)
            df_anual_nx_d['% Costo Final s/Venta'] = df_anual_nx_d['% Costo Final s/Venta'].apply(
                lambda x: f"{x:.2%}" if isinstance(x, float) else x)
            st.dataframe(df_anual_nx_d, use_container_width=True)

            # Exportar Excel
            st.markdown("---")
            out_t4 = io.BytesIO()
            df_det_all  = pd.DataFrame(detalle_t4)
            df_det_next = pd.DataFrame(detalle_next)
            _det_cols = ['MES/AÑO', 'Tipo', 'CC', 'Nombre Casino',
                         'Venta', 'Refacturación', 'Venta Neta',
                         'Costo UCP', '% Costo/Venta',
                         'Cargo Diferencial', 'Costo Final', '% Costo Final/Venta']
            with pd.ExcelWriter(out_t4, engine='openpyxl') as writer:
                df_res[[
                    'MES/AÑO', 'Tipo', 'Venta Total Transportados',
                    'Refacturación Total', 'Venta Neta Transportados',
                    'Costo Total UCP', '% Costo/Venta (UCP)',
                    'Costo CC 138 (Modelación)', '% Costo CC138/Venta', 'Diferencial',
                ]].to_excel(writer, index=False, sheet_name=f'Resumen {year_ucp}')
                df_det_all[_det_cols].to_excel(writer, index=False, sheet_name=f'Detalle {year_ucp}')
                df_anual.to_excel(writer, index=False, sheet_name=f'Anual {year_ucp}')
                df_res_nx[[
                    'MES/AÑO', 'Tipo', 'Venta Total Transportados',
                    'Refacturación Total', 'Venta Neta Transportados',
                    'Costo Total UCP', '% Costo/Venta (UCP)',
                    'Costo CC 138 (Modelación)', '% Costo CC138/Venta', 'Diferencial',
                ]].to_excel(writer, index=False, sheet_name=f'Resumen {year_next}')
                df_det_next[_det_cols].to_excel(writer, index=False, sheet_name=f'Detalle {year_next}')
                df_anual_nx.to_excel(writer, index=False, sheet_name=f'Anual {year_next}')
            st.download_button(
                label="📥 Descargar Distribución UCP (Excel)",
                data=out_t4.getvalue(),
                file_name=f"Distribucion_UCP_{year_ucp}_{year_next}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_dl_ucp_t4"
            )

        except Exception as e:
            st.error(f"Error al procesar distribución UCP: {e}")
            import traceback
            st.code(traceback.format_exc())