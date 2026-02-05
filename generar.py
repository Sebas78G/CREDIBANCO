import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import os
import glob
import re

# ======================================================
# CONFIGURACIÓN
# ======================================================
CARPETA_EXCEL_ENTRADA = "entrada"
CARPETA_VOUCHERS_OK = "vouchers_ok"
CARPETA_VOUCHERS_ERROR = "vouchers_error"

# ======================================================
# CREAR CARPETAS
# ======================================================
for c in [CARPETA_EXCEL_ENTRADA, CARPETA_VOUCHERS_OK, CARPETA_VOUCHERS_ERROR]:
    os.makedirs(c, exist_ok=True)

# ======================================================
# AUTO-DETECCIÓN DE ARCHIVOS EXCEL
# ======================================================
def detectar_exceles():
    """
    Auto-detecta cuál Excel es el validador y cuál es el de datos
    basándose en el número de columnas
    """
    archivos_excel = glob.glob(os.path.join(CARPETA_EXCEL_ENTRADA, "*.xlsx"))
    
    if len(archivos_excel) < 2:
        raise Exception(f"❌ ERROR: Se necesitan al menos 2 archivos Excel en '{CARPETA_EXCEL_ENTRADA}/'")
    
    info_exceles = []
    for archivo in archivos_excel:
        df = pd.read_excel(archivo, nrows=0)
        num_cols = len(df.columns)
        info_exceles.append({
            'ruta': archivo,
            'nombre': os.path.basename(archivo),
            'columnas': num_cols
        })
    
    # Ordenar por número de columnas (menor a mayor)
    info_exceles.sort(key=lambda x: x['columnas'])
    
    validador = info_exceles[0]
    datos = info_exceles[-1]
    
    print(f"✅ Validador detectado: {validador['nombre']} ({validador['columnas']} columnas)")
    print(f"✅ Datos detectado:     {datos['nombre']} ({datos['columnas']} columnas)")
    
    return validador['ruta'], datos['ruta']

# ======================================================
# NORMALIZACIÓN
# ======================================================
def normalizar_codigo(codigo):
    if pd.isna(codigo): return ''
    if isinstance(codigo, float): codigo = int(codigo)
    codigo = str(codigo).strip().replace(' ', '')
    if codigo.endswith('.0'): codigo = codigo[:-2]
    return codigo

def limpiar_valor(valor):
    if pd.isna(valor): return 0.0
    if isinstance(valor, (int, float)): return float(valor)
    valor = str(valor).replace('$', '').replace('COP', '').replace(' ', '')
    valor = valor.replace('.', '').replace(',', '.')
    try: return float(valor)
    except: return 0.0

def formatear_moneda(valor):
    return f"{valor:,.2f} COP".replace(',', 'X').replace('.', ',').replace('X', '.')

def formatear_tarjeta(t):
    if pd.isna(t): return "**** **** **** ****"
    t = str(t).replace('*', '').replace(' ', '')
    if len(t) >= 4: return f"**** **** **** {t[-4:]}"
    return "**** **** **** ****"

def obtener_franquicia(t):
    t = str(t).upper().replace('*', '').replace(' ', '')
    if 'VI' in t or t.startswith('4'): return "VISA"
    if 'MC' in t or t.startswith('5'): return "MASTERCARD"
    if t.startswith('3'): return "AMERICAN EXPRESS"
    return "VISA"

# ======================================================
# EXTRACCIÓN DE DATOS DE PARÁMETROS
# ======================================================
def extraer_info_transaccion(matches, aut_code):
    """
    Analiza las filas del Excel de pagos para extraer info
    de Aerolínea y Agencia separadamente
    """
    info = {
        'aerolinea': {
            'existe': False,
            'nombre': 'SATENA',
            'id': '53',
            'valor_base': 0.0,
            'impuesto': 0.0,
            'total': 0.0,
            'aut': aut_code,
            'comercio': '011029774'
        },
        'agencia': {
            'existe': False,
            'valor_base': 0.0,
            'impuesto': 0.0,
            'total': 0.0,
            'aut': aut_code, # Suele ser el mismo o +3
            'comercio': '011029774'
        },
        'general': {
            'nuevo_aut': aut_code # Por si cambia
        }
    }
    
    # Iterar sobre las filas encontradas
    for idx, row in matches.iterrows():
        params = str(row.get('Parámetros adicionales de pedido', ''))
        valor_fila = limpiar_valor(row.get('Valor total', 0))
        
        # Datos generales del primer registro
        if 'titular' not in info['general']:
            info['general']['titular'] = row.get('Titular de la tarjeta', '')
            info['general']['ip'] = row.get('IP', '34.232.176.163')
            info['general']['fecha'] = str(row.get('Fecha de pago', '')).split('.')[0] # Limpiar
        
        # Determinar si es AEROLINEA o AGENCIA
        es_aerolinea = 'airlineName' in params
        
        if es_aerolinea:
            info['aerolinea']['existe'] = True
            info['aerolinea']['total'] = valor_fila
            
            
            # Extraer Tax (el campo es airTax.amount, no airportTax)
            tax_match = re.search(r'airTax\.amount:([\d.]+)', params)
            if tax_match:
                info['aerolinea']['impuesto'] = float(tax_match.group(1))
            
            # Base = Total - Tax
            info['aerolinea']['valor_base'] = info['aerolinea']['total'] - info['aerolinea']['impuesto']
            
            # Extraer Nombre Aerolinea
            name_match = re.search(r'airlineName:([^,\]]+)', params)
            if name_match:
                info['aerolinea']['nombre'] = name_match.group(1)
                
            # Extraer ID Aerolinea
            id_match = re.search(r'airlineId:(\d+)', params)
            if id_match:
                info['aerolinea']['id'] = id_match.group(1)
                
        else:
            # Es AGENCIA
            info['agencia']['existe'] = True
            info['agencia']['total'] = valor_fila
            info['agencia']['valor_base'] = valor_fila # Asumimos IVA 0 por defecto
            
            # A veces la agencia tiene un aut diferente (ej: 971739 vs 971736)
            # Intentar ver si el excel tiene columna 'Código de aprobación' diferente en esta fila
            aut_fila = str(row.get('Código de aprobación', '')).strip()
            if aut_fila and aut_fila != 'nan':
                 info['agencia']['aut'] = aut_fila

    return info

# ======================================================
# GENERADOR PDF
# ======================================================
def generar_voucher_pdf(datos_validador, info_pago, nombre, carpeta):
    ruta = os.path.join(carpeta, nombre)
    c = canvas.Canvas(ruta, pagesize=letter)
    w, h = letter
    
    x_left = 80
    x_right = w - 80
    y = h - 60
    
    # 1. ENCABEZADO (fuente más delgada para coincidir con original)
    c.setFont("Helvetica", 24)
    c.drawString(x_left, y, "credibanco")
    y -= 50
    
    # 2. CAJA VERDE (colores exactos del original)
    c.setFillColor(colors.HexColor("#E8F5E9"))  # Verde claro fondo
    c.roundRect(x_left, y - 35, x_right - x_left, 50, 10, fill=1, stroke=0)
    c.setFillColor(colors.HexColor("#4CAF50"))  # Verde texto
    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(w/2, y - 10, "Pago exitoso")
    c.setFont("Helvetica", 12)
    c.drawCentredString(w/2, y - 28, "¡Gracias!")
    c.setFillColor(colors.black)
    y -= 70
    
    # 3. INFO COMERCIO
    c.setFont("Helvetica", 9)
    c.setFillColor(colors.HexColor("#757575"))  # Gris más claro
    c.drawString(x_left, y, "EXPRESO VIAJES Y TURISMO")
    ip = info_pago.get('general', {}).get('ip', '34.232.176.163')
    c.drawString(x_left, y - 12, f"IP {ip}")
    c.setFillColor(colors.black)
    y -= 35
    
    # 4. INFO PAGO HEADER
    c.setFont("Helvetica-Bold", 14)
    c.drawString(x_left, y, "Información del pago")
    y -= 30
    
    def fila(label, valor, color=colors.black, bold=True):
        nonlocal y
        c.setFont("Helvetica", 10)
        c.setFillColor(colors.HexColor("#757575"))  # Gris labels
        c.drawString(x_left, y, label)
        font = "Helvetica-Bold" if bold else "Helvetica"
        c.setFont(font, 10)
        c.setFillColor(color)
        c.drawRightString(x_right, y, str(valor))
        c.setFillColor(colors.black)
        y -= 15
        
    fila("Estado", "Aprobado", colors.HexColor("#4CAF50"))
    
    # Usar datos del validador preferiblemente, o info extraída
    fecha = str(datos_validador.get('FECHA', '')).replace('.', '/')
    fila("Fecha y hora", fecha)
    fila("Número de orden", str(datos_validador.get('TKT', '')))
    fila("Número de terminal", "00006760") # Fijo por ahora
    fila("Franquicia", obtener_franquicia(datos_validador.get('TARJETA', '')))
    fila("Número de tarjeta", formatear_tarjeta(datos_validador.get('TARJETA', '')))
    
    titular = info_pago.get('general', {}).get('titular', 'NUEVA EPS SA')
    fila("Titular de la Tarjeta", titular)
    
    # SECCIÓN AEROLINEA
    y -= 15
    data_air = info_pago.get('aerolinea', {})
    if data_air.get('existe'):
        c.setFont("Helvetica-Bold", 12)
        c.drawString(x_left, y, "AEROLINEA")
        y -= 20
        
        fila("Número de autorización", data_air['aut'])
        fila("Código de comercio", data_air['comercio'])
        fila("Nombre de la aerolínea", data_air['nombre'])
        fila("ID de aerolínea", data_air['id'])
        fila("Número de Cuotas", "1")
        fila("Valor a Pagar", formatear_moneda(data_air['valor_base']))
        fila("IVA", formatear_moneda(0))
        fila("Tasa aeroportuaria", formatear_moneda(data_air['impuesto']))
        
        c.setFont("Helvetica-Bold", 11)
        c.drawString(x_left, y, "Total")
        c.drawRightString(x_right, y, formatear_moneda(data_air['total']))
        y -= 25

    # SECCIÓN AGENCIA
    data_agency = info_pago.get('agencia', {})
    if data_agency.get('existe'):
        c.setFont("Helvetica-Bold", 12)
        c.drawString(x_left, y, "AGENCIA")
        y -= 20
        
        fila("Número de autorización", data_agency['aut'])
        fila("Código de comercio", data_agency['comercio'])
        fila("Número de Cuotas", "1")
        fila("Valor a Pagar", formatear_moneda(data_agency['valor_base']))
        fila("IVA", formatear_moneda(0))
        
        c.setFont("Helvetica-Bold", 11)
        c.drawString(x_left, y, "Total")
        c.drawRightString(x_right, y, formatear_moneda(data_agency['total']))
        y -= 25
        
    # TOTAL GENERAL
    total_gral = data_air.get('total', 0) + data_agency.get('total', 0)
    # Si no hubo desglose, usar valor original del validador para no mostrar 0
    if total_gral == 0:
        total_gral = limpiar_valor(datos_validador.get('VALOR', 0))

    c.setFont("Helvetica-Bold", 14)
    c.drawString(x_left, y, "Total")
    c.drawRightString(x_right, y, formatear_moneda(total_gral))
    y -= 40
    
    # FOOTER
    c.setFont("Helvetica", 7)
    c.setFillColor(colors.HexColor("#999999"))
    texto = ("Comprobante de pago venta no presencial ( * ) sujeto a verificación de la DIAN "
             "pagará incondicionalmente y a la orden del acreedor, el valor total de este pagaré "
             "junto con los intereses a las tasas máximas permitidas por la ley.")
    
    # Texto multilínea simple
    lines = [texto[i:i+95] for i in range(0, len(texto), 95)]
    for line in lines:
        c.drawString(x_left, y, line)
        y -= 9
        
    c.save()
    print(f"  📄 Generado: {nombre}")

# ======================================================
# EJECUCIÓN
# ======================================================
def procesar_vouchers():
    print("="*60)
    print("GENERADOR VOUCHERS - MODO EXTRACTOR EXACTO")
    print("="*60)
    
    try:
        r_val, r_dat = detectar_exceles()
    except Exception as e:
        print(e)
        return

    print("📖 Leyendo archivos...")
    df_val = pd.read_excel(r_val)
    df_pagos = pd.read_excel(r_dat)
    
    # Detectar columna AUT en validador
    col_aut_val = next((c for c in df_val.columns if 'AUT' in str(c).upper()), None)
    if not col_aut_val:
        print("❌ No se encontró columna AUT en validador")
        return
        
    # Detectar columna AUT en pagos
    col_aut_pagos = next((c for c in df_pagos.columns if 'APROBACI' in str(c).upper()), None)
    if not col_aut_pagos:
        print("❌ No se encontró columna Aprobación en pagos")
        return

    print(f"🔍 Columnas clave: Validador='{col_aut_val}' | Pagos='{col_aut_pagos}'")
    
    # Normalizar para búsquedas
    df_val['MATCH_KEY'] = df_val[col_aut_val].apply(normalizar_codigo)
    df_pagos['MATCH_KEY'] = df_pagos[col_aut_pagos].apply(normalizar_codigo)
    
    # Indexar pagos para búsqueda rápida (puede haber duplicados, así que no unique)
    print("⚙️  Procesando...")
    
    conteos = {'ok': 0, 'error': 0}
    
    for _, row_val in df_val.iterrows():
        key = row_val['MATCH_KEY']
        
        # Datos básicos del validador para el PDF
        datos_basic = {
            'TKT': row_val.get('TKT', ''),
            'FECHA': row_val.get('FECHA ', row_val.get('FECHA', '')), # Intentar con/sin espacio
            'TARJETA': row_val.get('TARJETA  ', row_val.get('TARJETA', '')),
            'VALOR': row_val.get('VALOR', 0),
            'AUT': row_val.get(col_aut_val, ''),
            'PNR': row_val.get('PNR ', row_val.get('PNR', ''))
        }
        
        nombre_archivo = f"TKT_{datos_basic['TKT']}_AUT_{key}.pdf".replace(' ', '_')
        
        # Buscar en pagos (primer match para obtener base de pedido)
        match_inicial = df_pagos[df_pagos['MATCH_KEY'] == key]
        
        if match_inicial.empty:
            # Error - no encontrado
            generar_voucher_pdf(datos_basic, {}, nombre_archivo, CARPETA_VOUCHERS_ERROR)
            conteos['error'] += 1
        else:
            # Obtener número de pedido del primer match
            num_pedido = match_inicial.iloc[0]['Número de pedido']
            
            # Extraer base del número de pedido (parte antes del _)
            if pd.notna(num_pedido):
                num_pedido_str = str(num_pedido)
                if '_' in num_pedido_str:
                    base_pedido = num_pedido_str.split('_')[0]
                else:
                    base_pedido = num_pedido_str
                
                # Buscar TODOS los registros con la misma base de pedido
                matches = df_pagos[df_pagos['Número de pedido'].astype(str).str.startswith(base_pedido + '_')]
                
                # Si no hay matches con _, intentar match exacto
                if matches.empty:
                    matches = df_pagos[df_pagos['Número de pedido'].astype(str) == base_pedido]
            else:
                # Si no hay número de pedido, usar solo el match inicial
                matches = match_inicial
            
            # OK - Extraer info detallada de TODOS los registros relacionados
            info_completa = extraer_info_transaccion(matches, key)
            generar_voucher_pdf(datos_basic, info_completa, nombre_archivo, CARPETA_VOUCHERS_OK)
            conteos['ok'] += 1
            
    print("\n" + "="*60)
    print(f"RESUMEN: OK={conteos['ok']} | ERROR={conteos['error']}")
    print("="*60)

if __name__ == "__main__":
    procesar_vouchers()
