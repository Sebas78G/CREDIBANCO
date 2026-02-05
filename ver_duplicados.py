import pandas as pd

df = pd.read_excel("entrada/pagos.xlsx")
print("Analizando registros duplicados por codigo de autorizacion...\n")

# Ver codigos que aparecen mas de una vez
df['Codigo_norm'] = df['Código de aprobación'].astype(str).str.strip()
codigos_duplicados = df['Codigo_norm'].value_counts()
codigos_duplicados = codigos_duplicados[codigos_duplicados > 1]

print(f"Codigos de autorizacion que aparecen mas de una vez: {len(codigos_duplicados)}\n")

if len(codigos_duplicados) > 0:
    print("Primeros 5 codigos duplicados:")
    print(codigos_duplicados.head())
    
    # Tomar el primer codigo duplicado como ejemplo
    cod_ejemplo = codigos_duplicados.index[0]
    print(f"\n\n=== EJEMPLO: Codigo {cod_ejemplo} ({codigos_duplicados[cod_ejemplo]} registros) ===\n")
    
    registros_ej = df[df['Codigo_norm'] == cod_ejemplo]
    
    for idx, (i, row) in enumerate(registros_ej.iterrows(), 1):
        print(f"REGISTRO {idx}:")
        print(f"  Valor total: {row['Valor total']}")
        print(f"  Numero de orden: {row['Número de pedido']}")
        
        params = str(row['Parámetros adicionales de pedido'])
        if 'airlineName' in params:
            print(f"  TIPO: AEROLINEA")
            # Extraer airline name
            import re
            match = re.search(r'airlineName:([^,\]]+)', params)
            if match:
                print(f"    - Nombre: {match.group(1)}")
        else:
            print(f"  TIPO: AGENCIA")
        
        # Ver si hay amounts en parametros
        if 'amount' in params:
            amounts = re.findall(r'amount:(\d+\.?\d*)', params)
            if amounts:
                print(f"    - Amounts encontrados: {amounts}")
        
        print()
else:
    print("No hay codigos duplicados.")
    print("\nBuscando codigo 971736 especificamente...")
    reg = df[df['Codigo_norm'] == '971736']
    if len(reg) > 0:
        print(f"Encontrado 1 registro con 971736")
        print(f"Valor total: {reg.iloc[0]['Valor total']}")
        print(f"\nParametros:")
        print(reg.iloc[0]['Parámetros adicionales de pedido'])
