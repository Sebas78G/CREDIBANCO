import pandas as pd
import re

# Verificar extracción con el patrón corregido
df_pagos = pd.read_excel("entrada/pagos.xlsx")
df_pagos['MATCH_KEY'] = df_pagos['Código de aprobación'].astype(str).str.strip()

matches = df_pagos[df_pagos['MATCH_KEY'] == '971736']

print("="*70)
print("VERIFICACIÓN CON PATRÓN CORREGIDO - airTax.amount")
print("="*70)
print(f"\nRegistros encontrados: {len(matches)}\n")

for idx, (i, row) in enumerate(matches.iterrows(), 1):
    params = str(row.get('Parámetros adicionales de pedido', ''))
    valor_total = row.get('Valor total', 0)
    
    es_aerolinea = 'airlineName' in params
    
    print(f"Registro {idx}: {'AEROLINEA' if es_aerolinea else 'AGENCIA'}")
    print(f"  Valor total: {valor_total:,.0f} COP")
    
    if es_aerolinea:
        # Usar el patrón CORRECTO
        tax_match = re.search(r'airTax\.amount:([\d.]+)', params)
        
        if tax_match:
            tax = float(tax_match.group(1))
            print(f"  ✅ airTax.amount encontrado: {tax:,.2f} COP")
            print(f"  Valor base calculado: {(valor_total - tax):,.2f} COP")
            print(f"  Total AEROLINEA: {valor_total:,.2f} COP")
        else:
            print(f"  ❌ airTax.amount NO encontrado")
            print(f"  Parámetros: {params[:200]}...")
    else:
        print(f"  Total AGENCIA: {valor_total:,.2f} COP")
    
    print()

print("="*70)
print("VALORES ESPERADOS:")
print("  AEROLINEA: 892,400.00 (base) + 36,400.00 (tax) = 928,800.00")
print("  AGENCIA: 83,000.00")
print("="*70)
