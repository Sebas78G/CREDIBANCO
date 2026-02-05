import pandas as pd
import re

print("ANALISIS SIMPLE - Autorizacion 971736")
print("=" * 60)

df = pd.read_excel("entrada/pagos.xlsx")
df['Codigo_norm'] = df['Código de aprobación'].astype(str).str.strip()

registros = df[df['Codigo_norm'] == '971736']

print(f"Total registros: {len(registros)}\n")

for idx, (i, row) in enumerate(registros.iterrows(), 1):
    print(f"REGISTRO {idx}:")
    
    # Columna de parámetros
    if 'Parámetros adicionales de pedido' in df.columns:
        params = str(row.get('Parámetros adicionales de pedido', ''))
        
        # Buscar airlineName
        if 'airlineName' in params:
            airline = re.search(r'airlineName:([^,\]]+)', params)
            if airline:
                print(f"  TIPO: AEROLINEA ({airline.group(1)})")
        else:
            print(f"  TIPO: AGENCIA")
        
        # Buscar amount
        amount = re.search(r'amount:(\d+)', params)
        if amount:
            print(f"  VALOR: {int(amount.group(1)):,}")
            
        # Buscar airportTax
        if 'airportTax' in params:
            tax = re.search(r'airportTax:(\d+)', params)
            if tax:
                print(f"  TASA AEROPORTUARIA: {int(tax.group(1)):,}")
    
    # Ver importe
    if 'Importe' in df.columns:
        importe = row.get('Importe')
        if pd.notna(importe):
            print(f"  IMPORTE COLUMNA: {importe}")
    
    print()
