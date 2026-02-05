import pandas as pd
import os

print("=" * 70)
print("ANÁLISIS DE ARCHIVOS EXCEL")
print("=" * 70)

# Analizar pagos.xlsx
print("\n📊 ARCHIVO: pagos.xlsx")
print("-" * 70)
df_pagos = pd.read_excel("entrada/pagos.xlsx")
print(f"Total de filas: {len(df_pagos)}")
print(f"Columnas: {list(df_pagos.columns)}")
print("\nPrimeras 3 filas:")
print(df_pagos.head(3))
print("\nTipos de datos:")
print(df_pagos.dtypes)

# Analizar validacion.xlsx
print("\n" + "=" * 70)
print("📊 ARCHIVO: validacion.xlsx")
print("-" * 70)
df_val = pd.read_excel("entrada/validacion.xlsx")
print(f"Total de filas: {len(df_val)}")
print(f"Columnas: {list(df_val.columns)}")
print("\nPrimeras 3 filas:")
print(df_val.head(3))
print("\nTipos de datos:")
print(df_val.dtypes)

# Exportar muestra
print("\n" + "=" * 70)
print("💾 Guardando muestras...")
df_pagos.head(5).to_excel("muestra_pagos.xlsx", index=False)
df_val.head(5).to_excel("muestra_validacion.xlsx", index=False)
print("✅ Archivos creados: muestra_pagos.xlsx y muestra_validacion.xlsx")
print("=" * 70)
