import pandas as pd

df_pagos = pd.read_excel("entrada/pagos.xlsx")
df_pagos['MATCH_KEY'] = df_pagos['Código de aprobación'].astype(str).str.strip()

print("="*70)
print("ANÁLISIS: ¿Cuántos códigos tienen múltiples registros?")
print("="*70)

# Contar registros por código
conteo = df_pagos['MATCH_KEY'].value_counts()

# Códigos con más de 1 registro
multiples = conteo[conteo > 1]

print(f"\nTotal códigos únicos: {len(conteo)}")
print(f"Códigos con 1 solo registro: {len(conteo[conteo == 1])}")
print(f"Códigos con múltiples registros: {len(multiples)}")

if len(multiples) > 0:
    print(f"\nPrimeros 10 códigos con múltiples registros:")
    for codigo, cantidad in multiples.head(10).items():
        print(f"  {codigo}: {cantidad} registros")
        
        # Ver qué tipo de registros son
        registros = df_pagos[df_pagos['MATCH_KEY'] == codigo]
        tipos = []
        for _, row in registros.iterrows():
            params = str(row.get('Parámetros adicionales de pedido', ''))
            if 'airlineName' in params:
                tipos.append('AEROLINEA')
            else:
                tipos.append('AGENCIA')
        print(f"    Tipos: {', '.join(tipos)}")
else:
    print("\n❌ NO hay códigos con múltiples registros")
    print("   Cada código de autorización tiene UN SOLO registro")
    
print("\n" + "="*70)
print("CONCLUSIÓN:")
if len(multiples) > 0:
    print("  Algunos códigos SÍ tienen múltiples registros")
else:
    print("  TODOS los códigos tienen UN SOLO registro")
    print("  La estructura NO es multi-registro como pensábamos")
