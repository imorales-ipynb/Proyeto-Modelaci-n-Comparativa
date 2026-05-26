import pandas as pd
import numpy as np

variables = ['Venta', 'Costo Alimento', 'Gasto Manipulación', 'Gasto Fijo', 'Gasto Variable', 'Margen de Contribución']

# Cliente A: Modelación 2026 completo. EVA Mayo 2026 a Abril 2027.
meses_mod_a = [f"{m:02d}-2026" for m in range(1, 13)]
meses_eva_a = [f"{m:02d}-2026" for m in range(5, 13)] + [f"{m:02d}-2027" for m in range(1, 5)]

# Todas las columnas que deben existir en el Excel
todas_columnas_meses = sorted(list(set(meses_mod_a + meses_eva_a)))

datos = []

# --- CLIENTE A ---
for var in variables:
    # FILA MODELACIÓN (Enero a Dic 2026)
    fila_mod = {'CC': 101, 'Nombre Cliente': 'Cliente A', 'Tipo Modelo': 'Modelación', 'Variable': var}
    for m in todas_columnas_meses:
        if m in meses_mod_a:
            # Ventas mas altas, costos menores, margen aprox 20%
            if var == 'Venta': val = round(np.random.uniform(1000, 1500), 2)
            elif var == 'Margen de Contribución': val = round(np.random.uniform(200, 300), 2)
            else: val = round(np.random.uniform(100, 200), 2)
            fila_mod[m] = val
        else:
            fila_mod[m] = ""
    datos.append(fila_mod)
    
    # FILA EVA (Mayo 2026 a Abril 2027)
    fila_eva = {'CC': 101, 'Nombre Cliente': 'Cliente A', 'Tipo Modelo': 'EVA', 'Variable': var}
    for m in todas_columnas_meses:
        if m in meses_eva_a:
            # EVA ligeramente mejor en ventas, mucho mejor margen
            if var == 'Venta': val = round(np.random.uniform(1100, 1600), 2)
            elif var == 'Margen de Contribución': val = round(np.random.uniform(250, 400), 2) # Margen mejorado
            else: val = round(np.random.uniform(90, 180), 2)
            fila_eva[m] = val
        else:
            fila_eva[m] = ""
    datos.append(fila_eva)

df = pd.DataFrame(datos)

# Reordenar columnas
columnas_ordenadas = ['CC', 'Nombre Cliente', 'Tipo Modelo', 'Variable'] + todas_columnas_meses
df = df[columnas_ordenadas]

df.to_excel('datos_prueba.xlsx', index=False)
print("Archivo 'datos_prueba.xlsx' actualizado con éxito para pruebas de estacionalidad.")
