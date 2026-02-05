import pandas as pd
import re

# Verificar la estructura con los ejemplos del usuario
df_pagos = pd.read_excel("entrada/pagos.xlsx")

print("="*70)
print("VERIFICACIÓN: Estructura de Número de pedido")
print("="*70)

# Buscar los números de pedido mencionados
ejemplos = ['224915_38257063', '224915_38257053', '224914_38257023', '224914_38257013']

for num_pedido in ejemplos:
    registros = df_pagos[df_pagos['Número de pedido'] == num_pedido]
    
    if len(registros) > 0:
        row = registros.iloc[0]
        codigo_auth = row.get('Código de aprobación', '')
        valor = row.get('Valor total', 0)
        params = str(row.get('Parámetros adicionales de pedido', ''))
        
        es_aerolinea = 'airlineName' in params
        tipo = 'AEROLINEA' if es_aerolinea else 'AGENCIA'
        
        print(f"\n{num_pedido}:")
        print(f"  Código auth: {codigo_auth}")
        print(f"  Tipo: {tipo}")
        print(f"  Valor: {valor:,.0f} COP")

# Extraer base de número de pedido
print("\n" + "="*70)
print("AGRUPACIÓN POR BASE DE NÚMERO DE PEDIDO")
print("="*70)

# Función para extraer base
def extraer_base_pedido(num_pedido):
    """Extrae la parte base del número de pedido (antes del _)"""
    if pd.isna(num_pedido):
        return ''
    num_str = str(num_pedido)
    if '_' in num_str:
        return num_str.split('_')[0]
    return num_str

# Agrupar
df_pagos['Base_Pedido'] = df_pagos['Número de pedido'].apply(extraer_base_pedido)

# Ver cuántos registros tiene cada base
for base in ['224915', '224914']:
    registros = df_pagos[df_pagos['Base_Pedido'] == base]
    print(f"\nBase {base}: {len(registros)} registros")
    
    if len(registros) > 0:
        for idx, (i, row) in enumerate(registros.iterrows(), 1):
            params = str(row.get('Parámetros adicionales de pedido', ''))
            es_aerolinea = 'airlineName' in params
            tipo = 'AEROLINEA' if es_aerolinea else 'AGENCIA'
            valor = row.get('Valor total', 0)
            
            print(f"  {idx}. {tipo}: {valor:,.0f} COP")
