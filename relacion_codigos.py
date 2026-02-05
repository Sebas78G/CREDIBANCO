import pandas as pd
import re

df_pagos = pd.read_excel("entrada/pagos.xlsx")

print("="*70)
print("RELACIÓN ENTRE 971736 (AEROLINEA) y 971739 (AGENCIA)")
print("="*70)

# Obtener ambos registros
reg_971736 = df_pagos[df_pagos['Código de aprobación'].astype(str).str.strip() == '971736'].iloc[0]
reg_971739 = df_pagos[df_pagos['Código de aprobación'].astype(str).str.strip() == '971739'].iloc[0]

print("\nCódigo 971736 (AEROLINEA):")
print(f"  Número de pedido: {reg_971736.get('Número de pedido', '')}")
print(f"  Valor total: {reg_971736.get('Valor total', 0):,.0f}")

# Extraer merchantOrderNumber de parámetros
params_736 = str(reg_971736.get('Parámetros adicionales de pedido', ''))
merchant_736 = re.search(r'merchantOrderNumber:(\d+)', params_736)
if merchant_736:
    print(f"  merchantOrderNumber: {merchant_736.group(1)}")

print("\nCódigo 971739 (AGENCIA):")
print(f"  Número de pedido: {reg_971739.get('Número de pedido', '')}")
print(f"  Valor total: {reg_971739.get('Valor total', 0):,.0f}")

params_739 = str(reg_971739.get('Parámetros adicionales de pedido', ''))
merchant_739 = re.search(r'merchantOrderNumber:(\d+)', params_739)
if merchant_739:
    print(f"  merchantOrderNumber: {merchant_739.group(1)}")

print("\n" + "="*70)
print("CONCLUSIÓN:")
if merchant_736 and merchant_739 and merchant_736.group(1) == merchant_739.group(1):
    print(f"  ✅ Ambos comparten merchantOrderNumber: {merchant_736.group(1)}")
    print("  → Se pueden relacionar por este campo")
else:
    num_pedido_736 = str(reg_971736.get('Número de pedido', ''))
    num_pedido_739 = str(reg_971739.get('Número de pedido', ''))
    
    if num_pedido_736 == num_pedido_739:
        print(f"  ✅ Ambos comparten Número de pedido: {num_pedido_736}")
        print("  → Se pueden relacionar por este campo")
    else:
        # Verificar si los códigos son consecutivos
        if int('971739') - int('971736') == 3:
            print("  ⚠️  Los códigos son consecutivos (+3)")
            print("  → Pueden estar relacionados por proximidad")
