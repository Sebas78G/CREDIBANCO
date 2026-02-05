import pandas as pd

df = pd.read_excel("entrada/pagos.xlsx")

# Buscar por valor total = 1011800 (el total del voucher)
print("Buscando registros con valor total 1,011,800...")
registros_total = df[df['Valor total'] == 1011800]

if len(registros_total) > 0:
    print(f"\nEncontrados {len(registros_total)} registros")
    for idx, (i, row) in enumerate(registros_total.iterrows(), 1):
        print(f"\nREGISTRO {idx}:")
        print(f"  Codigo de aprobacion: {row['Código de aprobación']}")
        print(f"  Numero de orden: {row['Número de pedido']}")
        print(f"  Valor total: {row['Valor total']}")
else:
    print("No encontrado con valor total 1,011,800")
    
    # Buscar todos los registros con código 971736
    print("\nBuscando TODOS los registros con codigo 971736...")
    df['Codigo_norm'] = df['Código de aprobación'].astype(str).str.strip()
    todos_971736 = df[df['Codigo_norm'] == '971736']
    
    print(f"Encontrados: {len(todos_971736)} registros")
    for idx, (i, row) in enumerate(todos_971736.iterrows(), 1):
        print(f"\n  Registro {idx}:")
        print(f"    Valor total: {row['Valor total']}")
        print(f"    Numero de orden: {row['Número de pedido']}")
