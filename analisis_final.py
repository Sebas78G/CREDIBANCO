import pandas as pd
import re

# Leer archivos
validador = pd.read_excel("entrada/validacion.xlsx")
pagos = pd.read_excel("entrada/pagos.xlsx")

# Normalizar nombres de columnas (quitar espacios)
validador.columns = [c.strip() for c in validador.columns]
pagos.columns = [c.strip() for c in pagos.columns]

print("COLUMNAS VALIDADOR:", validador.columns.tolist())
print()

# Buscar el primer registro con AUT = 971736
ejemplo = validador[validador['AUT'].astype(str).str.strip() == '971736']

if len(ejemplo) > 0:
    print("REGISTRO DEL VALIDADOR (AUT 971736):")
    print("=" * 70)
    for col in validador.columns:
        print(f"  {col:15s}: {ejemplo.iloc[0][col]}")
    
    # Ahora buscar en pagos
    print("\n" + "=" * 70)
    print("BUSCANDO EN PAGOS...")
    print("=" * 70)
    
    # Intentar buscar por código de aprobación   
    pagos_match = pagos[pagos['Código de aprobación'].astype(str).str.strip() == '971736']
    
    print(f"\nRegistros encontrados: {len(pagos_match)}")
    
    for idx, (i, row) in enumerate(pagos_match.iterrows(), 1):
        print(f"\nREGISTRO PAGOS #{idx}:")
        print(f"  Número de pedido: {row['Número de pedido']}")
        print(f"  Valor total: {row['Valor total']}")
        
        # Analizar parámetros
        params = str(row['Parámetros adicionales de pedido'])
        
        # Buscar airline
        if 'airlineName' in params:
            match_airline = re.search(r'airlineName:([^,\]]+)', params)
            if match_airline:
                print(f"  TIPO: AEROLINEA - {match_airline.group(1)}")
        
        # Buscar todos los "amount"
        amounts = re.findall(r'(\w+\.?amount):([\d.]+)', params)
        if amounts:
            print(f"  AMOUNTS encontrados:")
            for name, value in amounts:
                print(f"    {name}: {value}")
        
        # Buscar tax
        if 'airportTax' in params or 'tax' in params:
            tax = re.search(r'(?:airport)?Tax:([\d.]+)', params)
            if tax:
                print(f"  TAX: {tax.group(1)}")

else:
    print("No encontrado 971736 en validador")
