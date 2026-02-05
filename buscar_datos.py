import pandas as pd
import sys

try:
    # Leer el Excel de pagos
    print("Leyendo Excel de pagos...")
    df = pd.read_excel("entrada/pagos.xlsx")
    
    print(f"Total registros: {len(df)}")
    print(f"Total columnas: {len(df.columns)}")
    
    # Normalizar código para buscar
    df['Codigo_norm'] = df['Código de aprobación'].astype(str).str.strip()
    
    # Buscar 971736
    print("\nBuscando código de autorización: 971736")
    registro = df[df['Codigo_norm'] == '971736']
    
    if registro.empty:
        print("No encontrado, intentando como número...")
        registro = df[df['Código de aprobación'] == 971736]
    
    if not registro.empty:
        print(f"\n✅ ENCONTRADO! (Fila {registro.index[0]})")
        print("\nTODAS LAS COLUMNAS Y VALORES:")
        print("=" * 80)
        
        row = registro.iloc[0]
        for col in df.columns:
            valor = row[col]
            if pd.notna(valor):
                print(f"{col:50s} | {valor}")
        
    else:
        print("\n❌ No se encontró el código 971736")
        print("\nPrimeros códigos de aprobación en el archivo:")
        print(df['Código de aprobación'].head(20))

except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()
