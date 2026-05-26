# -*- coding: utf-8 -*-
"""
Script para corregir la indentacion rota en app.py (lineas 363-444)
"""
import os

fpath = r'c:\Users\ivan.morales\Desktop\Proyecto Modelacion Comparativa\app.py'

# Buscar el archivo con acento
for fname in os.listdir(r'c:\Users\ivan.morales\Desktop\Proyecto Modelacion Comparativa'):
    if 'app' in fname.lower():
        fpath = os.path.join(r'c:\Users\ivan.morales\Desktop\Proyecto Modelacion Comparativa', fname)
        break

# Intentar con acento en la ruta
fpath = 'c:\\Users\\ivan.morales\\Desktop\\Proyecto Modelaci\u00f3n Comparativa\\app.py'

with open(fpath, encoding='utf-8') as f:
    lines = f.readlines()

print(f"Total lineas: {len(lines)}")
print(f"Linea 363: {lines[362].rstrip()}")
print(f"Linea 364: {lines[363].rstrip()}")
print(f"Linea 444: {lines[443].rstrip()}")
print(f"Linea 445: {lines[444].rstrip()}")
