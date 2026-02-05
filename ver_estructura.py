import pandas as pd

# El código 971736 corresponde al voucher de ejemplo
# Vamos a buscar TODOS los registros y sus columnas relevantes

print("Analizando estructura del Excel de pagos...\n")

df = pd.read_excel("entrada/pagos.xlsx")

# Ver estructura
print("=" * 80)
print("COLUMNAS DEL EXCEL (52 total):")
print("=" * 80)
for i, col in enumerate(df.columns, 1):
    print(f"{i:2d}. {col}")

#Buscar columnas que puedan contener info de aerolínea/agencia
print("\n" + "=" * 80)
print("BÚSQUEDA DE PARÁMETROS ADICIONALES:")
print("=" * 80)

# Ver un registro ejemplo
if 'Parámetros adicionales de pedido' in df.columns:
    ejemplo = df['Parámetros adicionales de pedido'].iloc[0]
    print(f"\nEjemplo de parámetros:\n{ejemplo}")
