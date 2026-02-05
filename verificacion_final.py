import pandas as pd
import re

# Verificar que ahora se lean AMBOS registros para 971736
df_pagos = pd.read_excel("entrada/pagos.xlsx")
df_pagos['MATCH_KEY'] = df_pagos['Código de aprobación'].astype(str).str.strip()

# Buscar el primer match
match_inicial = df_pagos[df_pagos['MATCH_KEY'] == '971736']

if not match_inicial.empty:
    num_pedido = match_inicial.iloc[0]['Número de pedido']
    print(f"Código 971736 encontrado")
    print(f"Número de pedido: {num_pedido}")
    
    # Extraer base
    if pd.notna(num_pedido):
        num_pedido_str = str(num_pedido)
        if '_' in num_pedido_str:
            base_pedido = num_pedido_str.split('_')[0]
        else:
            base_pedido = num_pedido_str
        
        print(f"Base de pedido: {base_pedido}")
        
        # Buscar TODOS con esa base
        matches = df_pagos[df_pagos['Número de pedido'].astype(str).str.startswith(base_pedido + '_')]
        
        if matches.empty:
            matches = df_pagos[df_pagos['Número de pedido'].astype(str) == base_pedido]
        
        print(f"\n✅ Total registros encontrados: {len(matches)}")
        print("="*70)
        
        total_aerolinea = 0
        total_agencia = 0
        
        for idx, (i, row) in enumerate(matches.iterrows(), 1):
            params = str(row.get('Parámetros adicionales de pedido', ''))
            valor = row.get('Valor total', 0)
            codigo = row.get('Código de aprobación', '')
            
            es_aerolinea = 'airlineName' in params
            
            if es_aerolinea:
                print(f"\nRegistro {idx}: AEROLINEA")
                print(f"  Código auth: {codigo}")
                print(f"  Valor total: {valor:,.0f} COP")
                
                # Extraer tax
                tax_match = re.search(r'airTax\.amount:([\d.]+)', params)
                if tax_match:
                    tax = float(tax_match.group(1))
                    print(f"  Tasa: {tax:,.0f} COP")
                    print(f"  Base: {(valor - tax):,.0f} COP")
                
                total_aerolinea = valor
            else:
                print(f"\nRegistro {idx}: AGENCIA")
                print(f"  Código auth: {codigo}")
                print(f"  Valor total: {valor:,.0f} COP")
                total_agencia = valor
        
        print("\n" + "="*70)
        print("TOTALES:")
        print(f"  AEROLINEA: {total_aerolinea:,.0f} COP")
        print(f"  AGENCIA: {total_agencia:,.0f} COP")
        print(f"  TOTAL GENERAL: {(total_aerolinea + total_agencia):,.0f} COP")
        print("="*70)
        
        if total_aerolinea == 928800 and total_agencia == 83000:
            print("\n✅ ¡PERFECTO! Los valores coinciden exactamente")
        else:
            print("\n⚠️  Los valores no coinciden con el esperado")
