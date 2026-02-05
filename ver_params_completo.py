import pandas as pd
import ast

df = pd.read_excel("entrada/pagos.xlsx")
df['Codigo_norm'] = df['Código de aprobación'].astype(str).str.strip()

# Buscar 971736
registro = df[df['Codigo_norm'] == '971736'].iloc[0]

params_str = str(registro['Parámetros adicionales de pedido'])

print("PARAMETROS COMPLETOS (971736):")
print(params_str)
print()
print("=" * 70)

# El formato parece ser tipo Python list con dicts
# Intentar parsear
try:
    # Los parametros parecen estar en formato: [key:value, key:value, ...]
    # Convertir a dict
    params_dict = {}
    
    # Quitar corchetes
    params_clean = params_str.strip('[]')
    
    # Separar por comas pero cuidado con comas dentro de valores
    items = []
    current = ""
    depth = 0
    
    for char in params_clean:
        if char in '[{':
            depth += 1
        elif char in ']}':
            depth -= 1
        elif char == ',' and depth == 0:
            items.append(current.strip())
            current = ""
            continue
        current += char
    
    if current:
        items.append(current.strip())
    
    print("\nITEMS ENCONTRADOS:")
    for item in items:
        if ':' in item:
            parts = item.split(':', 1)
            key = parts[0].strip()
            value = parts[1].strip()
            params_dict[key] = value
            print(f"  {key:30s} = {value}")

except Exception as e:
    print(f"Error parseando: {e}")
    
# Buscar especificamente los valores que necesitamos
print("\n" + "=" * 70)
print("VALORES CLAVE PARA VOUCHER:")
print("=" * 70)

# Del voucher original sabemos:
# AEROLINEA: 892,400 + 36,400 tasa = 928,800  
# AGENCIA: 83,000
# TOTAL: 1,011,800

print(f"Valor total del registro: {registro['Valor total']}")
print("\nNecesitamos encontrar:")
print("  - Valor AEROLINEA: 892,400 o 892400")
print("  - Tasa aeroportuaria: 36,400 o 36400")  
print("  - Valor AGENCIA: 83,000 o 83000")
