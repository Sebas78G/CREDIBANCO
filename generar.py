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


CARPETA_EXCEL_ENTRADA = "excel"
CARPETA_VOUCHERS_OK = "vouchers_ok"


for c in [CARPETA_EXCEL_ENTRADA, CARPETA_VOUCHERS_OK]:
    os.makedirs(c, exist_ok=True)

def detectar_exceles():
    # Buscar archivos Excel (.xlsx) y OpenDocument (.ods)
    archivos_excel = []
    archivos_excel.extend(glob.glob(os.path.join(CARPETA_EXCEL_ENTRADA, "*.xlsx")))
    archivos_excel.extend(glob.glob(os.path.join(CARPETA_EXCEL_ENTRADA, "*.ods")))
    
    if len(archivos_excel) < 2:
        raise Exception(f"❌ ERROR: Se necesitan al menos 2 archivos Excel en '{CARPETA_EXCEL_ENTRADA}/'")
    
    info_exceles = []
    for archivo in archivos_excel:
        try:
            # Detectar el motor apropiado según la extensión
            motor = None
            if archivo.endswith('.ods'):
                motor = 'odf'  # Requiere instalar: pip install odfpy
            
            # Leer solo la primera fila para detectar columnas
            df = pd.read_excel(archivo, nrows=0, engine=motor)
            num_cols = len(df.columns)
            nombre = os.path.basename(archivo)
            
           
            tipo_detectado = "Desconocido"
            if num_cols <= 10: 
                tipo_detectado = "VOUCHER (pocas columnas)"
            elif num_cols >= 15:  
                tipo_detectado = "Transacciones completas (muchas columnas)"
            
            info_exceles.append({
                'ruta': archivo,
                'nombre': nombre,
                'columnas': num_cols,
                'tipo': tipo_detectado
            })
            
        except Exception as e:
            print(f"⚠️  No se pudo leer {os.path.basename(archivo)}: {e}")
            continue
    
    if len(info_exceles) < 2:
        raise Exception(f"❌ ERROR: No se pudieron leer suficientes archivos Excel válidos en '{CARPETA_EXCEL_ENTRADA}/'")
    
    info_exceles.sort(key=lambda x: x['columnas'])

    validador = info_exceles[0]
    
    datos = info_exceles[-1]
    
    print(f"✅ VOUCHER detectado:      {validador['nombre']}")
    print(f"   └─ Tipo: {validador['tipo']}")
    print(f"   └─ Columnas: {validador['columnas']}")
    print(f"")
    print(f"✅ Transacciones detectado: {datos['nombre']}")
    print(f"   └─ Tipo: {datos['tipo']}")
    print(f"   └─ Columnas: {datos['columnas']}")
    
    if abs(validador['columnas'] - datos['columnas']) < 5:
        print(f"")
        print(f"⚠️  ADVERTENCIA: Los archivos tienen números de columnas similares")
        print(f"   Validador: {validador['columnas']} | Datos: {datos['columnas']}")
        print(f"   Verifica que los archivos sean correctos.")
    
    return validador['ruta'], datos['ruta']


def normalizar_codigo(codigo):
    if pd.isna(codigo): return ''
    if isinstance(codigo, float): codigo = int(codigo)
    codigo = str(codigo).strip().replace(' ', '')
    if codigo.endswith('.0'): codigo = codigo[:-2]
    return codigo

def limpiar_numero(valor):
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
    t_str = str(t).replace(' ', '')  
    
    digitos = ''
    for char in reversed(t_str):
        if char.isdigit():
            digitos = char + digitos
        else:
            break  
    
    if len(digitos) >= 4:
        return f"**** **** **** {digitos[-4:]}"
    return "**** **** **** ****"

def obtener_franquicia(t):
    t = str(t).upper().replace('*', '').replace(' ', '')
    if 'VI' in t or t.startswith('4'): return "VISA"
    if 'MC' in t or t.startswith('5'): return "MASTERCARD"
    if t.startswith('3'): return "AMERICAN EXPRESS"
    return "VISA"


def detectar_columna(df, patrones):
   
    for columna in df.columns:
        col_normalizada = str(columna).strip().upper().replace(' ', '_')
        
        for patron in patrones:
            patron_norm = patron.strip().upper().replace(' ', '_')
            if patron_norm in col_normalizada or col_normalizada in patron_norm:
                return columna
    
    return None

def mapear_columnas_validador(df):
    
    mapa = {}
    

    patrones_columnas = {
        'TKT': ['TKT', 'TICKET', 'NUMERO', 'NUMBER', 'NUM_TKT'],
        'FECHA': ['FECHA', 'DATE', 'HORA', 'DATETIME', 'TIMESTAMP'],
        'TARJETA': ['TARJETA', 'CARD', 'NUMERO_TARJETA', 'CARD_NUMBER'],
        'VALOR': ['VALOR', 'TOTAL', 'MONTO', 'AMOUNT', 'IMPORTE'],
        'AUT': ['AUT', 'AUTORIZACION', 'CODIGO', 'APROBACION', 'AUTHORIZATION', 'AUTH', 'APPROVAL'],
        'PNR': ['PNR', 'LOCALIZADOR', 'BOOKING', 'RESERVA']
    }
    
    
    for clave, patrones in patrones_columnas.items():
        columna_detectada = detectar_columna(df, patrones)
        if columna_detectada:
            mapa[clave] = columna_detectada
            print(f"  ✓ {clave}: '{columna_detectada}'")
        else:
            print(f"  ⚠ {clave}: No detectada (será omitida)")
    
    return mapa

def consolidar_filas_por_autorizacion(df, mapa_columnas):
    """
    Consolida filas del Excel validador que tienen el mismo código de autorización.
    
    Problema: El Excel validador puede tener múltiples filas con el mismo código AUT
    pero con información diferente (ej: una tiene TKT, otra tiene PNR).
    
    Solución: Agrupar por código de autorización y combinar la información:
    - Elegir valores no vacíos de cada columna
    - Combinar múltiples TKTs separados por coma
    - Generar UNA sola fila por código de autorización único
    
    Returns:
        DataFrame consolidado con una fila por código de autorización único
    """
    if 'AUT' not in mapa_columnas:
        print("⚠️  No se puede consolidar: columna AUT no detectada")
        return df
    
    col_aut = mapa_columnas['AUT']
    
    # Normalizar códigos de autorización para agrupación
    df_temp = df.copy()
    df_temp['AUT_NORMALIZADO'] = df_temp[col_aut].apply(normalizar_codigo)
    
    # Identificar duplicados
    duplicados = df_temp[df_temp.duplicated(subset=['AUT_NORMALIZADO'], keep=False)]
    num_duplicados = len(duplicados)
    
    if num_duplicados == 0:
        print("✓ No se encontraron códigos de autorización duplicados")
        return df
    
    print(f"🔍 Detectados {num_duplicados} registros con códigos de autorización duplicados")
    
    # Función auxiliar para consolidar valores de un grupo
    def consolidar_grupo(grupo):
        """Combina información de múltiples filas con el mismo código AUT"""
        resultado = {}
        
        for col in grupo.columns:
            if col == 'AUT_NORMALIZADO':
                continue
                
            valores = grupo[col].dropna()
            valores = valores[valores.astype(str).str.strip() != '']
            
            if len(valores) == 0:
                resultado[col] = pd.NA
            elif col == mapa_columnas.get('TKT'):
                # Para TKT, combinar TODOS los valores únicos
                tkts_unicos = valores.apply(limpiar_numero).unique()
                # Filtrar valores vacíos
                tkts_unicos = [t for t in tkts_unicos if t and str(t).strip()]
                
                if len(tkts_unicos) == 0:
                    resultado[col] = pd.NA
                elif len(tkts_unicos) == 1:
                    resultado[col] = tkts_unicos[0]
                else:
                    # Guardar todos los TKTs separados por guión bajo para el nombre del archivo
                    resultado[col] = '_'.join(map(str, tkts_unicos))
            else:
                # Para otros campos, tomar el primer valor no vacío
                resultado[col] = valores.iloc[0]
        
        return pd.Series(resultado)
    
    # Consolidar grupos
    df_consolidado = df_temp.groupby('AUT_NORMALIZADO', as_index=False).apply(
        consolidar_grupo, include_groups=False
    )
    
    # Eliminar la columna temporal
    if 'AUT_NORMALIZADO' in df_consolidado.columns:
        df_consolidado = df_consolidado.drop(columns=['AUT_NORMALIZADO'])
    
    num_consolidados = len(df) - len(df_consolidado)
    print(f"✓ Consolidados {num_consolidados} registros duplicados")
    print(f"📊 Total a procesar: {len(df_consolidado)} vouchers únicos (de {len(df)} filas originales)")
    
    return df_consolidado


def extraer_info_transaccion(matches, aut_code):
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
            'aut': aut_code, 
            'comercio': '011029774'
        },
        'general': {
            'nuevo_aut': aut_code 
        }
    }
    
    
    for idx, row in matches.iterrows():
        params = str(row.get('Parámetros adicionales de pedido', ''))
        valor_fila = limpiar_valor(row.get('Valor total', 0))
        
        if 'titular' not in info['general']:
            info['general']['titular'] = row.get('Titular de la tarjeta', '')
            info['general']['ip'] = row.get('IP', '34.232.176.163')
            info['general']['fecha'] = str(row.get('Fecha de pago', '')).split('.')[0] 
            info['general']['tarjeta'] = row.get('Número de tarjeta', '') 
        
        es_aerolinea = 'airlineName' in params
        
        if es_aerolinea:
            info['aerolinea']['existe'] = True
            info['aerolinea']['total'] = valor_fila
            
            
            tax_match = re.search(r'airTax\.amount:([\d.]+)', params)
            if tax_match:
                info['aerolinea']['impuesto'] = float(tax_match.group(1))
            
            info['aerolinea']['valor_base'] = info['aerolinea']['total'] - info['aerolinea']['impuesto']
            
            name_match = re.search(r'airlineName:([^,\]]+)', params)
            if name_match:
                info['aerolinea']['nombre'] = name_match.group(1)
                
            id_match = re.search(r'airlineId:(\d+)', params)
            if id_match:
                info['aerolinea']['id'] = id_match.group(1)
                
        else:
            info['agencia']['existe'] = True
            info['agencia']['total'] = valor_fila
            info['agencia']['valor_base'] = valor_fila 
            
            aut_fila = str(row.get('Código de aprobación', '')).strip()
            if aut_fila and aut_fila != 'nan':
                 info['agencia']['aut'] = aut_fila

    return info


def generar_voucher_pdf(datos_validador, info_pago, nombre, carpeta):
    ruta = os.path.join(carpeta, nombre)
    c = canvas.Canvas(ruta, pagesize=letter)
    w, h = letter
    
    x_left = 140
    x_right = w - 140
    y = h - 40
    
    logo_path = os.path.join("img", "credibanco.png")
    if os.path.exists(logo_path):
        c.drawImage(logo_path, x_left, y - 20, width=120, height=30, preserveAspectRatio=True, mask='auto')
    else:
        c.setFont("Helvetica", 16)
        c.drawString(x_left, y, "credibanco")
    y -= 60
    
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
    
    marco_inicio_y = y
    
    y -= 20 
    
    c.setFont("Helvetica-Bold", 9)
    c.setFillColor(colors.HexColor("#6B7280"))
    c.drawString(x_left + 10, y, "EXPRESO VIAJES Y TURISMO")
    y -= 12
    
    c.setStrokeColor(colors.HexColor("#E5E7EB"))
    c.setLineWidth(1)
    c.line(x_left + 10, y, x_right - 10, y)
    c.setStrokeColor(colors.black)
    y -= 20
    
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
    
    num_tarjeta = info_pago.get('general', {}).get('tarjeta', datos_validador.get('TARJETA', ''))
    fila("Franquicia", obtener_franquicia(num_tarjeta))
    fila("Número de tarjeta", formatear_tarjeta(num_tarjeta))
    
    titular = info_pago.get('general', {}).get('titular', 'NUEVA EPS SA')
    fila("Titular de la Tarjeta", titular)
    y -= 8
    
    data_air = info_pago.get('aerolinea', {})
    if data_air.get('existe'):
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

    data_agency = info_pago.get('agencia', {})
    if data_agency.get('existe'):
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
        
    total_gral = data_air.get('total', 0) + data_agency.get('total', 0)
    if total_gral == 0:
        total_gral = limpiar_valor(datos_validador.get('VALOR', 0))
    c.setStrokeColor(colors.HexColor("#E5E7EB"))
    c.setLineWidth(1)
    c.line(x_left + 10, y, x_right - 10, y)
    c.setStrokeColor(colors.black)
    y -= 20

    c.setFont("Helvetica-Bold", 13)
    c.drawString(x_left + 10, y, "Total")
    c.drawRightString(x_right - 10, y, formatear_moneda(total_gral))
    y -= 25
    
    marco_fin_y = y
    marco_altura = marco_inicio_y - marco_fin_y
    c.setStrokeColor(colors.HexColor("#D1D5DB"))
    c.setLineWidth(1.5)
    c.roundRect(x_left, marco_fin_y, x_right - x_left, marco_altura, 10, fill=0, stroke=1)
    c.setStrokeColor(colors.black)
    y -= 15
    
    c.setFont("Helvetica", 7)
    c.setFillColor(colors.HexColor("#9CA3AF"))
    texto = ("Comprobante de pago venta no presencial ( * ) sujeto a verificación de la DIAN "
             "pagaré incondicionalmente y a la orden del acreedor, el valor total de este pagaré "
             "junto con los intereses a las tasas máximas permitidas por la ley.")
    
    lines = [texto[i:i+95] for i in range(0, len(texto), 95)]
    for line in lines:
        c.drawString(x_left, y, line)
        y -= 9
        
    c.save()
    print(f"  📄 Generado: {nombre}")



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
    # Detectar si el validador es .ods
    motor_val = 'odf' if r_val.endswith('.ods') else None
    motor_dat = 'odf' if r_dat.endswith('.ods') else None
    
    df_val = pd.read_excel(r_val, engine=motor_val)
    df_pagos = pd.read_excel(r_dat, engine=motor_dat)
    
    print("\n🔍 Detectando columnas del validador...")
    mapa_val = mapear_columnas_validador(df_val)
    
    if 'AUT' not in mapa_val:
        print("❌ ERROR: No se pudo detectar la columna de autorización en el validador")
        return
    if 'TKT' not in mapa_val:
        print("❌ ERROR: No se pudo detectar la columna de TKT en el validador")
        return
    
    # CONSOLIDAR FILAS DUPLICADAS POR CÓDIGO DE AUTORIZACIÓN
    print("\n🔄 Consolidando vouchers duplicados...")
    df_val = consolidar_filas_por_autorizacion(df_val, mapa_val)
        
    col_aut_pagos = next((c for c in df_pagos.columns if 'APROBACI' in str(c).upper()), None)
    if not col_aut_pagos:
        print("❌ No se encontró columna Aprobación en pagos")
        return

    print(f"\n🔍 Columnas clave: Validador AUT='{mapa_val['AUT']}' | Pagos='{col_aut_pagos}'")
    
    df_val['MATCH_KEY'] = df_val[mapa_val['AUT']].apply(normalizar_codigo)
    df_pagos['MATCH_KEY'] = df_pagos[col_aut_pagos].apply(normalizar_codigo)
    
    print("⚙️  Procesando...")
    
    conteos = {'ok': 0, 'error': 0}
    errores_detallados = []  
    
    for _, row_val in df_val.iterrows():
        key = row_val['MATCH_KEY']
        
        datos_basic = {
            'TKT': row_val.get(mapa_val.get('TKT'), '') if 'TKT' in mapa_val else '',
            'FECHA': row_val.get(mapa_val.get('FECHA'), '') if 'FECHA' in mapa_val else '',
            'TARJETA': row_val.get(mapa_val.get('TARJETA'), '') if 'TARJETA' in mapa_val else '',
            'VALOR': row_val.get(mapa_val.get('VALOR'), 0) if 'VALOR' in mapa_val else 0,
            'AUT': row_val.get(mapa_val.get('AUT'), '') if 'AUT' in mapa_val else '',
            'PNR': row_val.get(mapa_val.get('PNR'), '') if 'PNR' in mapa_val else ''
        }
        
        nombre_archivo = f"TKT_{limpiar_numero(datos_basic['TKT'])}_AUT_{key}.pdf".replace(' ', '_')
        
        match_inicial = df_pagos[df_pagos['MATCH_KEY'] == key]
        
        if match_inicial.empty:
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
            num_pedido = match_inicial.iloc[0]['Número de pedido']
            
            if pd.notna(num_pedido):
                num_pedido_str = str(num_pedido)
                if '_' in num_pedido_str:
                    base_pedido = num_pedido_str.split('_')[0]
                else:
                    base_pedido = num_pedido_str
                
                matches = df_pagos[df_pagos['Número de pedido'].astype(str).str.startswith(base_pedido + '_')]
                
                if matches.empty:
                    matches = df_pagos[df_pagos['Número de pedido'].astype(str) == base_pedido]
            else:
                matches = match_inicial
            
            info_completa = extraer_info_transaccion(matches, key)
            generar_voucher_pdf(datos_basic, info_completa, nombre_archivo, CARPETA_VOUCHERS_OK)
            conteos['ok'] += 1
    
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
