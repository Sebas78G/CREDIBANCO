import pandas as pd

print("=" * 80)
print("ANALISIS DE ESTRUCTURA: AUTORIZACION 971736")
print("=" * 80)

# Leer Excel
df = pd.read_excel("entrada/pagos.xlsx")

# Normalizar código
df['Codigo_norm'] = df['Código de aprobación'].astype(str).str.strip()

# Buscar TODOS los registros con código 971736
registros = df[df['Codigo_norm'] == '971736']

print(f"\nTotal de registros encontrados: {len(registros)}")

if not registros.empty:
    print("\n" + "=" * 80)
    print("REGISTROS PARA AUTORIZACION 971736:")
    print("=" * 80)
    
    for idx, (i, row) in enumerate(registros.iterrows(), 1):
        print(f"\n--- REGISTRO {idx} (Fila {i}) ---")
        
        # Buscar columnas clave
        cols_importantes = [
            'Código de aprobación',
            'Número de orden',
            'Importe',
            'Valor sin IVA',
            'Parámetros adicionales de pedido'
        ]
        
        for col in cols_importantes:
            if col in df.columns:
                valor = row[col]
                if pd.notna(valor):
                    # Si es parámetros, buscar airline
                    if col == 'Parámetros adicionales de pedido':
                        params_str = str(valor)
                        print(f"\n  {col}:")
                        
                        # Buscar airlineName
                        if 'airlineName' in params_str:
                            import re
                            airline_match = re.search(r'airlineName:([^,\]]+)', params_str)
                            if airline_match:
                                print(f"    -> AEROLINEA: {airline_match.group(1)}")
                        else:
                            print(f"    -> No tiene airlineName (probablemente AGENCIA)")
                        
                        # Buscar valores
                        amount_match = re.search(r'amount:(\d+)', params_str)
                        if amount_match:
                            print(f"    -> Valor: {amount_match.group(1)}")
                            
                        tax_match = re.search(r'tax:(\d+)', params_str)
                        if tax_match:
                            print(f"    -> Impuesto/Tasa: {tax_match.group(1)}")
                    else:
                        print(f"  {col}: {valor}")

print("\n" + "=" * 80)
