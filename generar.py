import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import os
import glob
import re
import sys

# Configurar encoding para Windows
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# ======================================================
# CONFIGURACIÓN
# ======================================================
CARPETA_EXCEL_ENTRADA = "entrada"
CARPETA_VOUCHERS_OK = "vouchers_ok"

# ======================================================
# CREAR CARPETAS
# ======================================================
for c in [CARPETA_EXCEL_ENTRADA, CARPETA_VOUCHERS_OK]:
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

def limpiar_numero(valor):
    """Limpia .0 de números flotantes para mostrarlos como enteros"""
    if pd.isna(valor): return ''
    valor_str = str(valor)
    if valor_str.endswith('.0'):
        return valor_str[:-2]
    return valor_str


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
    t_str = str(t).replace(' ', '')  # Eliminar espacios pero mantener asteriscos
    
    # Extraer solo los dígitos del final (después de los asteriscos)
    # Buscar los últimos dígitos contiguos
    digitos = ''
    for char in reversed(t_str):
        if char.isdigit():
            digitos = char + digitos
        else:
            break  # Detenerse al encontrar un asterisco u otro carácter
    
    if len(digitos) >= 4:
        return f"**** **** **** {digitos[-4:]}"
    return "**** **** **** ****"

def obtener_franquicia(t):
    t = str(t).upper().replace('*', '').replace(' ', '')
    if 'VI' in t or t.startswith('4'): return "VISA"
    if 'MC' in t or t.startswith('5'): return "MASTERCARD"
    if t.startswith('3'): return "AMERICAN EXPRESS"
    return "VISA"

# ======================================================
# DETECCIÓN FLEXIBLE DE COLUMNAS
# ======================================================
def detectar_columna(df, patrones):
    """
    Busca una columna en el DataFrame que coincida con alguno de los patrones dados.
    
    Args:
        df: DataFrame de pandas
        patrones: lista de strings a buscar (ej: ['TKT', 'TICKET', 'NUMERO_TKT'])
    
    Returns:
        nombre exacto de la columna encontrada o None
    """
    for columna in df.columns:
        # Normalizar nombre de columna: quitar espacios, mayúsculas
        col_normalizada = str(columna).strip().upper().replace(' ', '_')
        
        # Buscar coincidencia con algún patrón
        for patron in patrones:
            patron_norm = patron.strip().upper().replace(' ', '_')
            if patron_norm in col_normalizada or col_normalizada in patron_norm:
                return columna
    
    return None

def mapear_columnas_validador(df):
    """
    Detecta y mapea las columnas del Excel validador de forma robusta.
    
    Returns:
        dict: {'TKT': 'nombre_real_columna', 'FECHA': ..., etc}
    """
    mapa = {}
    
    # Definir patrones para cada columna esperada
    patrones_columnas = {
        'TKT': ['TKT', 'TICKET', 'NUMERO', 'NUMBER', 'NUM_TKT'],
        'FECHA': ['FECHA', 'DATE', 'HORA', 'DATETIME', 'TIMESTAMP'],
        'TARJETA': ['TARJETA', 'CARD', 'NUMERO_TARJETA', 'CARD_NUMBER'],
        'VALOR': ['VALOR', 'TOTAL', 'MONTO', 'AMOUNT', 'IMPORTE'],
        'AUT': ['AUT', 'AUTORIZACION', 'CODIGO', 'APROBACION', 'AUTHORIZATION', 'AUTH', 'APPROVAL'],
        'PNR': ['PNR', 'LOCALIZADOR', 'BOOKING', 'RESERVA']
    }
    
    # Detectar cada columna
    for clave, patrones in patrones_columnas.items():
        columna_detectada = detectar_columna(df, patrones)
        if columna_detectada:
            mapa[clave] = columna_detectada
            print(f"  ✓ {clave}: '{columna_detectada}'")
        else:
            print(f"  ⚠ {clave}: No detectada (será omitida)")
    
    return mapa

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
            info['general']['tarjeta'] = row.get('Número de tarjeta', '') # Del Excel pagos
        
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
    
    x_left = 140
    x_right = w - 140
    y = h - 40
    
    # 1. LOGO - Visible arriba a la izquierda
    logo_path = os.path.join("img", "credibanco.png")
    if os.path.exists(logo_path):
        c.drawImage(logo_path, x_left, y - 20, width=120, height=30, preserveAspectRatio=True, mask='auto')
    else:
        c.setFont("Helvetica", 16)
        c.drawString(x_left, y, "credibanco")
    y -= 60
    
    # 2. CAJA "Pago exitoso" con marco aparte
    c.setStrokeColor(colors.HexColor("#D1D5DB"))
    c.setLineWidth(1.5)
    c.setFillColor(colors.HexColor("#F3F4F6"))
    c.roundRect(x_left + 10, y - 45, x_right - x_left - 20, 60, 8, fill=1, stroke=1)
    c.setStrokeColor(colors.black)
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 18)
    c.drawCentredString(w/2, y - 15, "Pago exitoso")
    c.setFont("Helvetica", 10)
    c.setFillColor(colors.HexColor("#6B7280"))
    c.drawCentredString(w/2, y - 32, "¡Gracias!")
    c.setFillColor(colors.black)
    y -= 80
    
    # INICIO DEL MARCO PRINCIPAL - Antes de EXPRESO para incluirlo dentro
    marco_inicio_y = y
    
    y -= 20 # Padding superior para que el texto no quede pegado al borde
    
    c.setFont("Helvetica-Bold", 9)
    c.setFillColor(colors.HexColor("#6B7280"))
    c.drawString(x_left + 10, y, "EXPRESO VIAJES Y TURISMO")
    y -= 12
    
    # Línea separadora completa
    c.setStrokeColor(colors.HexColor("#E5E7EB"))
    c.setLineWidth(1)
    c.line(x_left + 10, y, x_right - 10, y)
    c.setStrokeColor(colors.black)
    y -= 20
    
    # 4. INFORMACIÓN DEL PAGO
    c.setFont("Helvetica-Bold", 13)
    c.drawString(x_left + 10, y, "Información del pago")
    y -= 20
    
    def fila(label, valor, color_valor=colors.black, bold=True):
        nonlocal y
        c.setFont("Helvetica", 9)
        c.setFillColor(colors.HexColor("#6B7280"))
        c.drawString(x_left + 10, y, label)
        font = "Helvetica-Bold" if bold else "Helvetica"
        c.setFont(font, 9)
        c.setFillColor(color_valor)
        c.drawRightString(x_right - 10, y, str(valor))
        c.setFillColor(colors.black)
        y -= 14
        
    fila("Estado", "Aprobado", colors.HexColor("#6DC4E8"))
    
    fecha = str(datos_validador.get('FECHA', '')).replace('.', '/')
    fila("Fecha y hora", fecha)
    fila("Número de orden", limpiar_numero(datos_validador.get('TKT', '')))
    fila("Número de terminal", "00006760")
    
    # Usar número de tarjeta del Excel de pagos (tiene asteriscos), fallback al validador
    num_tarjeta = info_pago.get('general', {}).get('tarjeta', datos_validador.get('TARJETA', ''))
    fila("Franquicia", obtener_franquicia(num_tarjeta))
    fila("Número de tarjeta", formatear_tarjeta(num_tarjeta))
    
    titular = info_pago.get('general', {}).get('titular', 'NUEVA EPS SA')
    fila("Titular de la Tarjeta", titular)
    y -= 8
    
    # SECCIÓN AEROLINEA
    data_air = info_pago.get('aerolinea', {})
    if data_air.get('existe'):
        # Línea separadora
        c.setStrokeColor(colors.HexColor("#E5E7EB"))
        c.setLineWidth(1)
        c.line(x_left + 10, y, x_right - 10, y)
        c.setStrokeColor(colors.black)
        y -= 18
        
        c.setFont("Helvetica-Bold", 11)
        c.drawString(x_left + 10, y, "AEROLINEA")
        y -= 18
        
        fila("Número de autorización", data_air['aut'])
        fila("Código de comercio", data_air['comercio'])
        fila("Nombre de la aerolínea", data_air['nombre'])
        fila("ID de aerolínea", data_air['id'])
        fila("Número de Cuotas", "1")
        fila("Valor a Pagar", formatear_moneda(data_air['valor_base']))
        fila("IVA", formatear_moneda(0))
        fila("Tasa aeroportuaria", formatear_moneda(data_air['impuesto']))
        
        c.setFont("Helvetica-Bold", 10)
        c.setFillColor(colors.HexColor("#6B7280"))
        c.drawString(x_left + 10, y, "Total")
        c.setFillColor(colors.black)
        c.drawRightString(x_right - 10, y, formatear_moneda(data_air['total']))
        y -= 20

    # SECCIÓN AGENCIA
    data_agency = info_pago.get('agencia', {})
    if data_agency.get('existe'):
        # Línea separadora
        c.setStrokeColor(colors.HexColor("#E5E7EB"))
        c.setLineWidth(1)
        c.line(x_left + 10, y, x_right - 10, y)
        c.setStrokeColor(colors.black)
        y -= 18
        
        c.setFont("Helvetica-Bold", 11)
        c.drawString(x_left + 10, y, "AGENCIA")
        y -= 18
        
        fila("Número de autorización", data_agency['aut'])
        fila("Código de comercio", data_agency['comercio'])
        fila("Número de Cuotas", "1")
        fila("Valor a Pagar", formatear_moneda(data_agency['valor_base']))
        fila("IVA", formatear_moneda(0))
        
        c.setFont("Helvetica-Bold", 10)
        c.setFillColor(colors.HexColor("#6B7280"))
        c.drawString(x_left + 10, y, "Total")
        c.setFillColor(colors.black)
        c.drawRightString(x_right - 10, y, formatear_moneda(data_agency['total']))
        y -= 20
        
    # TOTAL GENERAL
    total_gral = data_air.get('total', 0) + data_agency.get('total', 0)
    if total_gral == 0:
        total_gral = limpiar_valor(datos_validador.get('VALOR', 0))

    # Línea separadora final
    c.setStrokeColor(colors.HexColor("#E5E7EB"))
    c.setLineWidth(1)
    c.line(x_left + 10, y, x_right - 10, y)
    c.setStrokeColor(colors.black)
    y -= 20

    c.setFont("Helvetica-Bold", 13)
    c.drawString(x_left + 10, y, "Total")
    c.drawRightString(x_right - 10, y, formatear_moneda(total_gral))
    y -= 25
    
    # MARCO GENERAL con bordes redondeados
    marco_fin_y = y
    marco_altura = marco_inicio_y - marco_fin_y
    c.setStrokeColor(colors.HexColor("#D1D5DB"))
    c.setLineWidth(1.5)
    c.roundRect(x_left, marco_fin_y, x_right - x_left, marco_altura, 10, fill=0, stroke=1)
    c.setStrokeColor(colors.black)
    y -= 15
    
    # FOOTER
    c.setFont("Helvetica", 7)
    c.setFillColor(colors.HexColor("#9CA3AF"))
    texto = ("Comprobante de pago venta no presencial ( * ) sujeto a verificación de la DIAN "
             "pagaré incondicionalmente y a la orden del acreedor, el valor total de este pagaré "
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
    
    # MAPEO ROBUSTO DE COLUMNAS DEL VALIDADOR
    print("\n🔍 Detectando columnas del validador...")
    mapa_val = mapear_columnas_validador(df_val)
    
    # Verificar que se detectaron las columnas críticas
    if 'AUT' not in mapa_val:
        print("❌ ERROR: No se pudo detectar la columna de autorización en el validador")
        return
    if 'TKT' not in mapa_val:
        print("❌ ERROR: No se pudo detectar la columna de TKT en el validador")
        return
        
    # Detectar columna AUT en pagos
    col_aut_pagos = next((c for c in df_pagos.columns if 'APROBACI' in str(c).upper()), None)
    if not col_aut_pagos:
        print("❌ No se encontró columna Aprobación en pagos")
        return

    print(f"\n🔍 Columnas clave: Validador AUT='{mapa_val['AUT']}' | Pagos='{col_aut_pagos}'")
    
    # Normalizar para búsquedas
    df_val['MATCH_KEY'] = df_val[mapa_val['AUT']].apply(normalizar_codigo)
    df_pagos['MATCH_KEY'] = df_pagos[col_aut_pagos].apply(normalizar_codigo)
    
    # Indexar pagos para búsqueda rápida (puede haber duplicados, así que no unique)
    print("⚙️  Procesando...")
    
    conteos = {'ok': 0, 'error': 0}
    errores_detallados = []  # Lista para rastrear errores
    
    for _, row_val in df_val.iterrows():
        key = row_val['MATCH_KEY']
        
        # Datos básicos del validador para el PDF (usar mapeo dinámico)
        datos_basic = {
            'TKT': row_val.get(mapa_val.get('TKT'), '') if 'TKT' in mapa_val else '',
            'FECHA': row_val.get(mapa_val.get('FECHA'), '') if 'FECHA' in mapa_val else '',
            'TARJETA': row_val.get(mapa_val.get('TARJETA'), '') if 'TARJETA' in mapa_val else '',
            'VALOR': row_val.get(mapa_val.get('VALOR'), 0) if 'VALOR' in mapa_val else 0,
            'AUT': row_val.get(mapa_val.get('AUT'), '') if 'AUT' in mapa_val else '',
            'PNR': row_val.get(mapa_val.get('PNR'), '') if 'PNR' in mapa_val else ''
        }
        
        nombre_archivo = f"TKT_{limpiar_numero(datos_basic['TKT'])}_AUT_{key}.pdf".replace(' ', '_')
        
        # Buscar en pagos (primer match para obtener base de pedido)
        match_inicial = df_pagos[df_pagos['MATCH_KEY'] == key]
        
        if match_inicial.empty:
            # Error - no encontrado, solo registrar en el reporte Excel
            observacion = f"No se encontró el código de autorización '{key}' en el Excel de pagos"
            errores_detallados.append({
                'Número de Autorización': key,
                'TKT': limpiar_numero(datos_basic['TKT']),
                'Fecha': datos_basic['FECHA'],
                'Valor': limpiar_valor(datos_basic['VALOR']),
                'Observación': observacion
            })
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
    
    # Generar reporte de errores en Excel
    if errores_detallados:
        print("\n📊 Generando reporte de errores...")
        df_errores = pd.DataFrame(errores_detallados)
        archivo_reporte = "reporte_errores_vouchers.xlsx"
        df_errores.to_excel(archivo_reporte, index=False, sheet_name='Errores')
        print(f"✅ Reporte generado: {archivo_reporte}")
        print(f"   Total de errores: {len(errores_detallados)}")
            
    print("\n" + "="*60)
    print(f"RESUMEN: OK={conteos['ok']} | ERROR={conteos['error']}")
    print("="*60)

if __name__ == "__main__":
    procesar_vouchers()
