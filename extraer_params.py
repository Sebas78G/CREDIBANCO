import pandas as pd
import json
import re

df = pd.read_excel("entrada/pagos.xlsx")

# Buscar 971736
df['Codigo_norm'] = df['Código de aprobación'].astype(str).str.strip()
registro = df[df['Codigo_norm'] == '971736'].iloc[0]

print("CODIGO: 971736")
print("=" * 70)
print(f"Valor total: {registro['Valor total']}")
print(f"Numero de orden: {registro['Número de pedido']}")
print()

# Ver parametros adicionales
params_str = str(registro['Parámetros adicionales de pedido'])
print("PARAMETROS ADICIONALES:")
print("=" * 70)
print(params_str)
print()

# Intentar parsear como JSON-like
print("\nEXTRACCION DE VALORES:")
print("=" * 70)

# Buscar valores clave
patterns = {
    'airlineName': r'airlineName:([^,\]]+)',
    'airlineId': r'airlineId:(\d+)',
    'amount': r'(?:^|,\s*)amount:(\d+)',
    'airportTax': r'airportTax:(\d+)',
    'IVA.amount': r'IVA\.amount:([\d.]+)',
}

for key, pattern in patterns.items():
    matches = re.findall(pattern, params_str)
    if matches:
        print(f"{key:20s}: {matches}")

# Detectar si hay multiples "amount" (uno para airline, otro para agency)
all_amounts = re.findall(r'amount:(\d+)', params_str)
print(f"\nTodos los 'amount':    {all_amounts}")

# Ver si hay info de products o items
if 'products' in params_str or 'items' in params_str:
    print("\n¡CONTIENE INFO DE PRODUCTS O ITEMS!")
