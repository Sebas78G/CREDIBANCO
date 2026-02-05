import pandas as pd

df_pagos = pd.read_excel("entrada/pagos.xlsx")

print("="*70)
print("BÚSQUEDA: ¿Dónde está el registro de AGENCIA con 83,000?")
print("="*70)

# Buscar registros con valor 83000
registros_83000 = df_pagos[df_pagos['Valor total'] == 83000]

print(f"\nRegistros con Valor total = 83,000: {len(registros_83000)}")

if len(registros_83000) > 0:
    print(f"\nPrimeros 5 registros:")
    for idx, (i, row) in enumerate(registros_83000.head().iterrows(), 1):
        codigo = row.get('Código de aprobación', '')
        numero_orden = row.get('Número de pedido', '')
        params = str(row.get('Parámetros adicionales de pedido', ''))
        
        es_aerolinea = 'airlineName' in params
        tipo = 'AEROLINEA' if es_aerolinea else 'AGENCIA'
        
        print(f"\n  {idx}. Código: {codigo}")
        print(f"     Número de orden: {numero_orden}")
        print(f"     Tipo: {tipo}")
        print(f"     Valor: 83,000")

# También buscar por número de orden similar a 971736
print("\n" + "="*70)
print("BÚSQUEDA: Códigos de autorización cercanos a 971736")
print("="*70)

# Buscar códigos entre 971730 y 971740
for codigo_test in range(971730, 971745):
    registros = df_pagos[df_pagos['Código de aprobación'].astype(str).str.strip() == str(codigo_test)]
    if len(registros) > 0:
        for _, row in registros.iterrows():
            valor = row.get('Valor total', 0)
            params = str(row.get('Parámetros adicionales de pedido', ''))
            es_aerolinea = 'airlineName' in params
            tipo = 'AEROLINEA' if es_aerolinea else 'AGENCIA'
            
            print(f"  {codigo_test}: {tipo} - {valor:,.0f} COP")
