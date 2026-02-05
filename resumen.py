import os

# Contar archivos
ok_dir = "vouchers_ok"
error_dir = "vouchers_error"

ok_count = len([f for f in os.listdir(ok_dir) if f.endswith('.pdf')])
error_count = len([f for f in os.listdir(error_dir) if f.endswith('.pdf')])

print("=" * 70)
print("RESUMEN DE VOUCHERS GENERADOS")
print("=" * 70)
print(f"✅ Vouchers OK (encontrados):     {ok_count}")
print(f"❌ Vouchers ERROR (no encontr.):  {error_count}")
print(f"📊 TOTAL PROCESADO:               {ok_count + error_count}")
print("=" * 70)
print(f"\n📁 Ubicaciones:")
print(f"   ✅ OK:    {ok_dir}/")
print(f"   ❌ ERROR: {error_dir}/")
print("=" * 70)
