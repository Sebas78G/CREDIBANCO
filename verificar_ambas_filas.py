import pandas as pd
import re

# Verificar que AMBAS filas se están leyendo
df_pagos = pd.read_excel("entrada/pagos.xlsx")
df_pagos['MATCH_KEY'] = df_pagos['Código de aprobación'].astype(str).str.strip()

matches = df_pagos[df_pagos['MATCH_KEY'] == '971736']

print("="*70)
print("VERIFICACIÓN: ¿Se leen AMBAS filas?")
print("="*70)
print(f"\nTotal de registros encontrados para 971736: {len(matches)}")

if len(matches) == 0:
    print("\n❌ NO se encontró ningún registro")
elif len(matches) == 1:
    print("\n⚠️  SOLO se encontró 1 registro (debería haber 2)")
    params = str(matches.iloc[0]['Parámetros adicionales de pedido'])
    if 'airlineName' in params:
        print("   Es registro de AEROLINEA")
    else:
        print("   Es registro de AGENCIA")
else:
    print(f"\n✅ Se encontraron {len(matches)} registros")
    
    tiene_aerolinea = False
    tiene_agencia = False
    
    for idx, (i, row) in enumerate(matches.iterrows(), 1):
        params = str(row.get('Parámetros adicionales de pedido', ''))
        valor = row.get('Valor total', 0)
        
        es_aerolinea = 'airlineName' in params
        
        if es_aerolinea:
            tiene_aerolinea = True
            print(f"\n  Registro {idx}: AEROLINEA")
            print(f"    Valor total: {valor:,.0f} COP")
        else:
            tiene_agencia = True
            print(f"\n  Registro {idx}: AGENCIA")
            print(f"    Valor total: {valor:,.0f} COP")
    
    print("\n" + "="*70)
    print("RESUMEN:")
    print(f"  ✅ Tiene AEROLINEA: {tiene_aerolinea}")
    print(f"  ✅ Tiene AGENCIA: {tiene_agencia}")
    
    if tiene_aerolinea and tiene_agencia:
        print("\n✅ PERFECTO: Ambos tipos de registro están presentes")
    else:
        print("\n⚠️  PROBLEMA: Falta uno de los tipos de registro")
