import pandas as pd

# Read both files
df_val = pd.read_excel("entrada/validacion.xlsx")
df_pagos = pd.read_excel("entrada/pagos.xlsx")

print("="*60)
print("ANÁLISIS DE MATCHING")
print("="*60)

# Normalize column names
df_val.columns = [c.strip() for c in df_val.columns]
df_pagos.columns = [c.strip() for c in df_pagos.columns]

# Get auth columns
col_aut_val = 'AUT'
col_aut_pagos = 'Código de aprobación'

# Normalize codes
df_val['MATCH_KEY'] = df_val[col_aut_val].astype(str).str.strip()
df_pagos['MATCH_KEY'] = df_pagos[col_aut_pagos].astype(str).str.strip()

print(f"\nTotal en validador: {len(df_val)}")
print(f"Total en pagos: {len(df_pagos)}")

# Count unique auth codes
unique_val = df_val['MATCH_KEY'].nunique()
unique_pagos = df_pagos['MATCH_KEY'].nunique()

print(f"\nCódigos únicos en validador: {unique_val}")
print(f"Códigos únicos en pagos: {unique_pagos}")

# Find matches
matches = 0
no_matches = 0
multi_record = 0

for code in df_val['MATCH_KEY'].unique():
    found = df_pagos[df_pagos['MATCH_KEY'] == code]
    if len(found) > 0:
        matches += 1
        if len(found) > 1:
            multi_record += 1
    else:
        no_matches += 1

print(f"\n✅ Códigos con match: {matches}")
print(f"❌ Códigos sin match: {no_matches}")
print(f"📊 Códigos con múltiples registros: {multi_record}")

# Show some examples of no-match codes
print(f"\nEjemplos de códigos SIN match (primeros 10):")
no_match_codes = []
for code in df_val['MATCH_KEY'].unique():
    if len(df_pagos[df_pagos['MATCH_KEY'] == code]) == 0:
        no_match_codes.append(code)
        if len(no_match_codes) >= 10:
            break

for code in no_match_codes:
    print(f"  {code}")

# Check if 971736 is in the validator
if '971736' in df_val['MATCH_KEY'].values:
    print(f"\n✅ Código 971736 ESTÁ en el validador")
else:
    print(f"\n❌ Código 971736 NO está en el validador")
