import pandas as pd

# Leer ambos archivos
validador = pd.read_excel("entrada/validacion.xlsx")
pagos = pd.read_excel("entrada/pagos.xlsx")

print("=" * 70)
print("ANALISIS DE RELACION ENTRE ARCHIVOS")
print("=" * 70)

print("\nVALIDADOR:")
print(f"  Columnas: {validador.columns.tolist()}")
print(f"  Total: {len(validador)} registros")

# Ver primer registro del validador
print("\n  Primer registro:")
for col in validador.columns:
    print(f"    {col:15s}: {validador.iloc[0][col]}")

print("\n" + "=" * 70)
print("PAGOS:")
print(f"  Columnas relevantes: Número de pedido, Código de aprobación, Valor total")
print(f"  Total: {len(pagos)} registros")

# Ver el primer registro de pagos
print("\n  Primer registro pagos:")
print(f"    Número de pedido: {pagos.iloc[0]['Número de pedido']}")
print(f"    Código de aprobación: {pagos.iloc[0]['Código de aprobación']}")
print(f"    Valor total: {pagos.iloc[0]['Valor total']}")

print("\n" + "=" * 70)
print("BUSQUEDA DE COINCIDENCIAS:")
print("=" * 70)

# El validador tiene AUT (código de autorización)
# El pagos tiene "Código de aprobación"
# Intentar match

# Normalizar
validador['AUT_norm'] = validador['AUT'].astype(str).str.strip()
pagos['Codigo_norm'] = pagos['Código de aprobación'].astype(str).str.strip()

# Buscar primer registro del validador en pagos
primer_aut = validador.iloc[0]['AUT_norm']
print(f"\nBuscando AUT '{primer_aut}' del validador en pagos...")

match_pago = pagos[pagos['Codigo_norm'] == primer_aut]

if len(match_pago) > 0:
    print(f"  ✓ ENCONTRADO! ({len(match_pago)} registro(s))")
    for idx, (i, row) in enumerate(match_pago.iterrows(), 1):
        print(f"\n  Match {idx}:")
        print(f"    Número de pedido: {row['Número de pedido']}")
        print(f"    Valor total: {row['Valor total']}")
        
        # Ver si tiene info de airline/agency en parámetros    
        params = str(row['Parámetros adicionales de pedido'])
        if 'airlineName' in params:
            print(f"    Contiene: AEROLINEA")
        else:
            print(f"    Contiene: Posiblemente AGENCIA")
else:
    print(f"  ✗ NO ENCONTRADO")
    print(f"\n  Códigos disponibles en pagos (primeros 10):")
    print(f"  {pagos['Código de aprobación'].head(10).tolist()}")
