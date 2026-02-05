import pandas as pd
import re

# Test the extraction logic on code 971736
df_pagos = pd.read_excel("entrada/pagos.xlsx")

# Normalize
df_pagos['MATCH_KEY'] = df_pagos['Código de aprobación'].astype(str).str.strip()

# Find 971736
matches = df_pagos[df_pagos['MATCH_KEY'] == '971736']

print(f"Registros encontrados para 971736: {len(matches)}")
print()

for idx, (i, row) in enumerate(matches.iterrows(), 1):
    print(f"=== REGISTRO {idx} ===")
    params = str(row.get('Parámetros adicionales de pedido', ''))
    valor_total = row.get('Valor total', 0)
    
    print(f"Valor total: {valor_total:,.0f}")
    
    # Check if airline
    es_aerolinea = 'airlineName' in params
    print(f"Tipo: {'AEROLINEA' if es_aerolinea else 'AGENCIA'}")
    
    if es_aerolinea:
        # Extract airline name
        name_match = re.search(r'airlineName:([^,\]]+)', params)
        if name_match:
            print(f"  Nombre: {name_match.group(1)}")
        
        # Extract airline ID
        id_match = re.search(r'airlineId:(\d+)', params)
        if id_match:
            print(f"  ID: {id_match.group(1)}")
        
        # Extract tax
        tax_match = re.search(r'airportTax:(\d+)', params)
        if tax_match:
            tax = float(tax_match.group(1))
            print(f"  Tasa aeroportuaria: {tax:,.0f}")
            print(f"  Valor base: {valor_total - tax:,.0f}")
    
    print()

print("\nVALORES ESPERADOS DEL VOUCHER ORIGINAL:")
print("  AEROLINEA: 892,400 + 36,400 tasa = 928,800")
print("  AGENCIA: 83,000")
print("  TOTAL: 1,011,800")
