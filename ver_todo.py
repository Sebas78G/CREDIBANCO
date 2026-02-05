import pandas as pd

df = pd.read_excel("entrada/pagos.xlsx")

# Buscar 971736
df['Codigo_norm'] = df['Código de aprobación'].astype(str).str.strip()
registros = df[df['Codigo_norm'] == '971736']

print(f"Registros encontrados: {len(registros)}")
print()

# Ver TODAS las columnas numéricas de estos registros
for idx, (i, row) in enumerate(registros.iterrows(), 1):
    print(f"=== REGISTRO {idx} ===")
    
    # Columnas numericas importantes
    cols_ver = ['Importe', 'Valor sin IVA', 'IVA', 'Valor', 'Código de aprobación']
    
    for col in df.columns:
        val = row[col]
        if pd.notna(val):
            # Mostrar si es numero o si contiene info relevante
            if isinstance(val, (int, float)):
                if val != 0:
                    print(f"  {col}: {val}")
            elif isinstance(val, str) and ('airline' in val.lower() or 'amount' in val.lower() or 'tax' in val.lower()):
                # Mostrar parte de parametros
                if len(val) > 200:
                    print(f"  {col}: {val[:200]}...")
                else:
                    print(f"  {col}: {val}")
    print()
