import pandas as pd
import sys

# Redirigir salida a archivo
output_file = open("analisis_resultado.txt", "w", encoding="utf-8")
sys.stdout = output_file

print("=" * 80)
print("ANÁLISIS DETALLADO DE ARCHIVOS EXCEL")
print("=" * 80)

# ============================================================================
# ARCHIVO 1: validacion.xlsx
# ============================================================================
print("\n📋 ARCHIVO 1: validacion.xlsx")
print("-" * 80)
df_val = pd.read_excel("entrada/validacion.xlsx")
print(f"\n📊 Total de registros: {len(df_val)}")
print(f"\n📝 Columnas encontradas ({len(df_val.columns)}):")
for i, col in enumerate(df_val.columns, 1):
    print(f"  {i}. {col}")

print(f"\n🔍 Primeras 5 filas de ejemplo:")
with pd.option_context('display.max_columns', None, 'display.width', None):
    print(df_val.head(5))

# ============================================================================
# ARCHIVO 2: pagos.xlsx
# ============================================================================
print("\n\n" + "=" * 80)
print("📋 ARCHIVO 2: pagos.xlsx")
print("-" * 80)
df_pagos = pd.read_excel("entrada/pagos.xlsx")
print(f"\n📊 Total de registros: {len(df_pagos)}")
print(f"\n📝 Columnas encontradas ({len(df_pagos.columns)}):")
for i, col in enumerate(df_pagos.columns, 1):
    print(f"  {i}. {col}")

print(f"\n🔍 Primeras 5 filas de ejemplo:")
with pd.option_context('display.max_columns', None, 'display.width', None):
    print(df_pagos.head(5))

# ============================================================================
# RESUMEN
# ============================================================================
print("\n\n" + "=" * 80)
print("📌 RESUMEN")
print("=" * 80)
print(f"validacion.xlsx: {len(df_val)} filas × {len(df_val.columns)} columnas")
print(f"pagos.xlsx:      {len(df_pagos)} filas × {len(df_pagos.columns)} columnas")
print("=" * 80)

output_file.close()
print("✅ Análisis guardado en: analisis_resultado.txt", file=sys.__stdout__)
