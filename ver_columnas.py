import pandas as pd
import sys

try:
    # Leer pagos.xlsx
    print("Leyendo pagos.xlsx...")
    df_pagos = pd.read_excel("entrada/pagos.xlsx", nrows=3)
    print("\n=== PAGOS.XLSX ===")
    print(f"Columnas: {list(df_pagos.columns)}")
    print("\n" + str(df_pagos))
    
    # Leer validacion.xlsx
    print("\n\nLeyendo validacion.xlsx...")
    df_val = pd.read_excel("entrada/validacion.xlsx", nrows=3)
    print("\n=== VALIDACION.XLSX ===")
    print(f"Columnas: {list(df_val.columns)}")
    print("\n" + str(df_val))
    
except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()
