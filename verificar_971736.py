import pandas as pd
import re

# Verificar extracción para código 971736
df_pagos = pd.read_excel("entrada/pagos.xlsx")
df_pagos['MATCH_KEY'] = df_pagos['Código de aprobación'].astype(str).str.strip()

matches = df_pagos[df_pagos['MATCH_KEY'] == '971736']

print("="*70)
print("VERIFICACIÓN DE EXTRACCIÓN - CÓDIGO 971736")
print("="*70)
print(f"\nRegistros encontrados: {len(matches)}\n")

total_aerolinea = 0
total_agencia = 0
tax_aerolinea = 0

for idx, (i, row) in enumerate(matches.iterrows(), 1):
    params = str(row.get('Parámetros adicionales de pedido', ''))
    valor_total = row.get('Valor total', 0)
    
    es_aerolinea = 'airlineName' in params
    
    print(f"Registro {idx}: {'AEROLINEA' if es_aerolinea else 'AGENCIA'}")
    print(f"  Valor total: {valor_total:,.0f} COP")
    
    if es_aerolinea:
        # Extract airline info
        name_match = re.search(r'airlineName:([^,\]]+)', params)
        id_match = re.search(r'airlineId:(\d+)', params)
        tax_match = re.search(r'airportTax:(\d+)', params)
        
        if name_match:
            print(f"  Aerolínea: {name_match.group(1)}")
        if id_match:
            print(f"  ID: {id_match.group(1)}")
        if tax_match:
            tax = float(tax_match.group(1))
            tax_aerolinea = tax
            print(f"  Tasa aeroportuaria: {tax:,.0f} COP")
            print(f"  Valor base: {(valor_total - tax):,.0f} COP")
        
        total_aerolinea = valor_total
    else:
        total_agencia = valor_total
    
    print()

print("="*70)
print("COMPARACIÓN CON VOUCHER ORIGINAL")
print("="*70)

print("\n📋 VALORES EXTRAÍDOS:")
print(f"  AEROLINEA:")
print(f"    - Valor base: {(total_aerolinea - tax_aerolinea):,.0f} COP")
print(f"    - Tasa: {tax_aerolinea:,.0f} COP")
print(f"    - Total: {total_aerolinea:,.0f} COP")
print(f"  AGENCIA:")
print(f"    - Total: {total_agencia:,.0f} COP")
print(f"  TOTAL GENERAL: {(total_aerolinea + total_agencia):,.0f} COP")

print("\n✅ VALORES ESPERADOS (del voucher original):")
print(f"  AEROLINEA:")
print(f"    - Valor base: 892,400 COP")
print(f"    - Tasa: 36,400 COP")
print(f"    - Total: 928,800 COP")
print(f"  AGENCIA:")
print(f"    - Total: 83,000 COP")
print(f"  TOTAL GENERAL: 1,011,800 COP")

print("\n" + "="*70)
# Verificar si coinciden
if total_aerolinea == 928800 and total_agencia == 83000:
    print("✅ ¡PERFECTO! Los valores coinciden exactamente")
elif total_aerolinea > 0:
    print("⚠️  Valores encontrados pero pueden no coincidir exactamente")
    print("    Esto puede deberse a que el código 971736 tiene diferentes")
    print("    valores en el Excel de pagos vs el voucher original mostrado")
else:
    print("❌ No se encontraron valores")
print("="*70)
