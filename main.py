import os
import sys
import subprocess

# =====================================================================
# 🛠️ VERIFICADOR E INSTALADOR AUTOMÁTICO DE DEPENDENCIAS
# =====================================================================
def verificar_e_instalar_dependencias():
    # Lista de librerías necesarias mapeadas a sus nombres en pip
    librerias = {
        "webview": "pywebview",
        "flask": "Flask",
        "pandas": "pandas",
        "openpyxl": "openpyxl",
        "xlrd": "xlrd"
    }
    
    faltantes = []
    for modulo, paquete in librerias.items():
        try:
            __import__(modulo)
        except ImportError:
            faltantes.append(paquete)
            
    if faltantes:
        print(f"📦 Detectadas librerías faltantes: {faltantes}")
        print("⏳ Instalando dependencias en segundo plano, por favor espera...")
        try:
            # Ejecuta pip usando el mismo intérprete de Python que está corriendo el script
            subprocess.check_call([sys.executable, "-m", "pip", "install", *faltantes])
            print("✅ ¡Todas las dependencias se instalaron correctamente!")
        except Exception as e:
            print(f"❌ Error crítico instalando las dependencias: {e}")
            sys.exit(1)

# Ejecutamos la verificación ANTES de importar las librerías pesadas
verificar_e_instalar_dependencias()

# =====================================================================
# 🚀 ARRANQUE DE LA APLICACIÓN
# =====================================================================
import webview
import threading
from flask import Flask, jsonify, render_template, request

app = Flask(__name__, static_folder='frontend', template_folder='frontend')

@app.route('/')
def index():
    return render_template('index.html')

import sqlite3
import pandas as pd

@app.route('/procesar', methods=['POST'])
def procesar():
    try:
        # Verificar si se recibió el archivo
        if 'archivo' not in request.files:
            return jsonify({'error': 'No se recibió ningún archivo en la petición'}), 400
            
        file = request.files['archivo']
        if file.filename == '':
            return jsonify({'error': 'Nombre de archivo vacío'}), 400
            
        # 1. Leer el archivo Excel usando Pandas
        # file.stream permite leer el archivo directamente desde memoria sin guardarlo en disco
        df = pd.read_excel(file.stream, engine='openpyxl')
        
        # Normalizar nombres de columnas a mayúsculas y quitar espacios para evitar errores de tipeo
        df.columns = df.columns.astype(str).str.strip().str.upper()
        
        # Columnas requeridas para filtrar (en mayúsculas)
        columnas_filtro = ['DISPONIBLE', 'EN PLANTA', 'PAGADO']
        
        # Columnas requeridas para guardar (clave en mayúsculas, valor es el nombre final en la base de datos)
        columnas_guardar_map = {
            'NRO. CONTENEDOR': 'NUM_CONTENEDOR',
            'PTO. DESCARGA': 'PTO_DESCARGA',
            'PLANO': 'PLANO',
            'DESC. PLANO': 'DESC_PLANO',
            'CANTIDAD': 'CANTIDAD'
        }
        
        # Verificar que existan todas las columnas requeridas en el DataFrame
        columnas_requeridas = columnas_filtro + list(columnas_guardar_map.keys())
        columnas_faltantes = [col for col in columnas_requeridas if col not in df.columns]
        
        if columnas_faltantes:
            # Reconvertir a título para mostrar un mensaje amigable
            faltantes_formateados = [col.title() for col in columnas_faltantes]
            return jsonify({'error': f'El archivo no contiene las columnas necesarias: {", ".join(faltantes_formateados)}'}), 400
            
        # Limpiar y normalizar los valores de las columnas de filtro (eliminar espacios y pasar a mayúsculas)
        for col in columnas_filtro:
            df[col] = df[col].astype(str).str.strip().str.upper()
            
        # 2. Filtrar donde DISPONIBLE == "SI" Y EN PLANTA == "SI" Y PAGADO == "SI"
        df_filtrado = df[
            (df['DISPONIBLE'] == 'SI') &
            (df['EN PLANTA'] == 'SI') &
            (df['PAGADO'] == 'SI')
        ]
        
        # 3. Seleccionar las columnas requeridas y renombrarlas a lo pedido para la base de datos
        df_final = df_filtrado[list(columnas_guardar_map.keys())].rename(columns=columnas_guardar_map)
        
        # 4. Guardar en SQLite
        # Esto crea automáticamente el archivo 'datos_inventario.db' si no existe
        db_path = os.path.join(os.getcwd(), 'datos_inventario.db')
        conexion = sqlite3.connect(db_path)
        
        # to_sql con if_exists='replace' elimina la tabla 'inventario' existente y crea una nueva con los datos actuales
        # index=False evita guardar el índice numérico de pandas como columna
        df_final.to_sql('inventario', conexion, if_exists='replace', index=False)
        
        # Ejecutar VACUUM para compactar el archivo de la base de datos en disco y liberar espacio
        conexion.execute("VACUUM")
        conexion.close()
        
        return jsonify({
            'success': True,
            'registros': len(df_final),
            'mensaje': f'Se guardaron {len(df_final)} registros filtrados en la base de datos.'
        })
        
    except Exception as e:
        return jsonify({'error': f'Error al procesar el Excel: {str(e)}'}), 500

@app.route('/procesar_piezas', methods=['POST'])
def procesar_piezas():
    try:
        if 'archivo' not in request.files:
            return jsonify({'error': 'No se recibio ningun archivo'}), 400

        file = request.files['archivo']
        if file.filename == '':
            return jsonify({'error': 'Nombre de archivo vacio'}), 400

        from datetime import date, timedelta, datetime
        import io
        import openpyxl

        file_bytes = file.read()

        hoy = date.today()
        manana = hoy + timedelta(days=1)

        def parse_as_date(col_val):
            if isinstance(col_val, (datetime, date)):
                d = col_val.date() if isinstance(col_val, datetime) else col_val
                if 2000 <= d.year <= 2100:
                    return d
                return None
            
            col_str = str(col_val).strip()
            if "." in col_str:
                col_str_cleaned = col_str.split(".")[0]
            else:
                col_str_cleaned = col_str

            try:
                dt = pd.to_datetime(col_str_cleaned)
                if 2000 <= dt.year <= 2100:
                    return dt.date()
            except:
                pass

            for fmt in ["%d/%m", "%d-%m", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"]:
                try:
                    dt = datetime.strptime(col_str_cleaned, fmt)
                    if "%Y" not in fmt:
                        dt = dt.replace(year=date.today().year)
                    if 2000 <= dt.year <= 2100:
                        return dt.date()
                except:
                    pass
            return None

        # Prioritize sheet names
        try:
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
            sheets = wb.sheetnames
            wb.close()
        except Exception as e:
            return jsonify({'error': f'No se pudo abrir el archivo Excel: {str(e)}'}), 400

        def get_sheet_priority(name):
            n = name.lower()
            if n in ['visão geral', 'visao geral']:
                return 0
            if 'geral' in n or 'general' in n:
                return 1
            if 'base' in n:
                return 2
            if 'inventario' in n or 'inventário' in n:
                return 3
            return 4
            
        sheets_sorted = sorted(sheets, key=get_sheet_priority)

        mejor_sheet = None
        mejor_fila = None
        mejor_score = -1
        mejor_col_mapping = {}

        for sheet in sheets_sorted:
            for header_row in range(16):
                try:
                    df_test = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, header=header_row, nrows=2)
                    cols = list(df_test.columns)

                    ref_col = None
                    des_col = None
                    flujo_espe_col = None
                    darsena_col = None
                    mag_bdl_col = None
                    hoy_col = None
                    manana_col = None

                    # Ref matching (exact first, then case-insensitive prefixes)
                    for c in cols:
                        c_str = str(c).strip().lower()
                        if c_str == "ref":
                            ref_col = c
                            break
                    if ref_col is None:
                        for c in cols:
                            c_str = str(c).strip().lower()
                            if c_str == "refs":
                                ref_col = c
                                break
                    if ref_col is None:
                        for c in cols:
                            c_str = str(c).strip().lower()
                            if "ref" in c_str and "ref." in c_str:
                                ref_col = c
                                break
                        if ref_col is None:
                            for c in cols:
                                c_str = str(c).strip().lower()
                                if "ref" in c_str:
                                    ref_col = c
                                    break

                    # Designación matching
                    for c in cols:
                        c_str = str(c).strip().lower()
                        if c_str in ["designação", "designacao", "desig", "designação nfc", "designação bhm"]:
                            if des_col is None or len(c_str) < len(str(des_col)):
                                des_col = c
                    if des_col is None:
                        for c in cols:
                            c_str = str(c).strip().lower()
                            if "designa" in c_str or "desig" in c_str:
                                des_col = c
                                break

                    # Flujo Espe matching
                    for c in cols:
                        c_str = str(c).strip().lower()
                        if c_str == "flujo espe" or c_str == "fluxo espe":
                            flujo_espe_col = c
                            break
                    if flujo_espe_col is None:
                        for c in cols:
                            c_str = str(c).strip().lower()
                            if "flujo esp" in c_str or "fluxo esp" in c_str:
                                flujo_espe_col = c
                                break

                    # Darsena matching
                    for c in cols:
                        c_str = str(c).strip().lower()
                        if c_str == "darsena" or c_str == "dársena":
                            darsena_col = c
                            break
                    if darsena_col is None:
                        for c in cols:
                            c_str = str(c).strip().lower()
                            if "darsena" in c_str or "sgr (darsena)" in c_str:
                                darsena_col = c
                                break

                    # MAG + BDL matching
                    for c in cols:
                        c_str = str(c).strip().lower()
                        if c_str == "mag + bdl" or c_str == "mag+bdl":
                            mag_bdl_col = c
                            break
                    if mag_bdl_col is None:
                        for c in cols:
                            c_str = str(c).strip().lower()
                            if "mag" in c_str and "bdl" in c_str:
                                mag_bdl_col = c
                                break

                    # Today & Tomorrow date matching
                    # First try exact match with today & manana
                    for c in cols:
                        d = parse_as_date(c)
                        if d == hoy and hoy_col is None:
                            hoy_col = c
                        elif d == manana and manana_col is None:
                            manana_col = c

                    # Fallback date detection if exact not found: find any date columns closest to today
                    if hoy_col is None or manana_col is None:
                        parsed_dates = []
                        for c in cols:
                            d = parse_as_date(c)
                            if d is not None:
                                parsed_dates.append((c, d))
                        
                        # Sort by difference to today
                        parsed_dates_sorted = sorted(parsed_dates, key=lambda x: abs((x[1] - hoy).days))
                        if len(parsed_dates_sorted) >= 2:
                            two_closest = sorted(parsed_dates_sorted[:2], key=lambda x: x[1])
                            if hoy_col is None:
                                hoy_col = two_closest[0][0]
                            if manana_col is None:
                                manana_col = two_closest[1][0]

                    score = sum(1 for x in [ref_col, des_col, flujo_espe_col, darsena_col, mag_bdl_col, hoy_col, manana_col] if x is not None)

                    is_better = score > mejor_score
                    if score == mejor_score and mejor_sheet is not None:
                        if sheet.lower() in ['visão geral', 'visao geral'] and mejor_sheet.lower() not in ['visão geral', 'visao geral']:
                            is_better = True

                    if is_better:
                        mejor_score = score
                        mejor_sheet = sheet
                        mejor_fila = header_row
                        mejor_col_mapping = {
                            'Ref': ref_col,
                            'Designacao': des_col,
                            'Flujo Espe': flujo_espe_col,
                            'Darsena': darsena_col,
                            'MAG + BDL': mag_bdl_col,
                            'flujo_hoy': hoy_col,
                            'flujo_manana': manana_col
                        }

                        if score == 7:
                            break
                except Exception:
                    continue
            if mejor_score == 7:
                break

        if mejor_score < 4 or mejor_sheet is None:
            muestra_sheets = ", ".join(sheets)
            df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheets[0], header=None, nrows=10)
            muestra_raw = df_raw.to_string()
            return jsonify({'error': f'No se pudieron detectar columnas suficientes (mínimo 4). '
                                     f'Hojas disponibles: {muestra_sheets}. '
                                     f'Muestra de la primera hoja ({sheets[0]}):\n{muestra_raw}'}), 400

        # Check for missing columns
        missing_cols = [k for k, v in mejor_col_mapping.items() if v is None]
        if missing_cols:
            df_all_cols = pd.read_excel(io.BytesIO(file_bytes), sheet_name=mejor_sheet, header=mejor_fila, nrows=1)
            disponibles = ", ".join([str(c) for c in df_all_cols.columns])
            return jsonify({'error': f'Columnas no encontradas en hoja "{mejor_sheet}" fila {mejor_fila + 1}: '
                                     f'{", ".join(missing_cols)}. '
                                     f'Columnas disponibles: {disponibles}'}), 400

        # Read the full sheet with the chosen header row
        df_full = pd.read_excel(io.BytesIO(file_bytes), sheet_name=mejor_sheet, header=mejor_fila)

        # Select and rename to canonical columns
        selected_cols = [mejor_col_mapping[c] for c in ['Ref', 'Designacao', 'Flujo Espe', 'Darsena', 'MAG + BDL', 'flujo_hoy', 'flujo_manana']]
        df_final = df_full[selected_cols].copy()
        
        rename_dict = {
            mejor_col_mapping['Ref']: 'Ref',
            mejor_col_mapping['Designacao']: 'Designacao',
            mejor_col_mapping['Flujo Espe']: 'Flujo Espe',
            mejor_col_mapping['Darsena']: 'Darsena',
            mejor_col_mapping['MAG + BDL']: 'MAG + BDL',
            mejor_col_mapping['flujo_hoy']: 'flujo_hoy',
            mejor_col_mapping['flujo_manana']: 'flujo_manana'
        }
        df_final = df_final.rename(columns=rename_dict)

        # Clean Ref column: make sure it is string, no decimal suffix, stripped
        df_final['Ref'] = df_final['Ref'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        
        # Drop rows where Ref is empty, NaN or 'nan'
        df_final = df_final[df_final['Ref'] != '']
        df_final = df_final[df_final['Ref'].notna()]
        df_final = df_final[df_final['Ref'].str.lower() != 'nan']

        # Save to database
        db_path = os.path.join(os.getcwd(), 'datos_inventario.db')
        conexion = sqlite3.connect(db_path)
        df_final.to_sql('piezas', conexion, if_exists='replace', index=False)
        conexion.execute("VACUUM")
        conexion.close()

        # Format today and tomorrow dates nicely for response log message
        def get_display_date_name(col_val):
            d = parse_as_date(col_val)
            if d is not None:
                return f"{d.day}/{d.month}"
            return str(col_val)

        col_hoy_str = get_display_date_name(mejor_col_mapping['flujo_hoy'])
        col_man_str = get_display_date_name(mejor_col_mapping['flujo_manana'])

        return jsonify({
            'success': True,
            'registros': len(df_final),
            'col_hoy': col_hoy_str,
            'col_man': col_man_str,
            'mensaje': f'Se guardaron {len(df_final)} registros de la hoja "{mejor_sheet}" en la tabla piezas.'
        })

    except Exception as e:
        return jsonify({'error': f'Error al procesar el archivo de piezas: {str(e)}'}), 500


@app.route('/contenedores', methods=['GET'])
def get_contenedores():
    try:
        db_path = os.path.join(os.getcwd(), 'datos_inventario.db')
        if not os.path.exists(db_path):
            return jsonify({'contenedores': []})
            
        conexion = sqlite3.connect(db_path)
        cursor = conexion.cursor()
        # Query unique containers and their discharge point.
        cursor.execute("SELECT NUM_CONTENEDOR, MAX(PTO_DESCARGA) FROM inventario GROUP BY NUM_CONTENEDOR ORDER BY NUM_CONTENEDOR")
        rows = cursor.fetchall()
        conexion.close()
        
        contenedores = [{'num_contenedor': r[0], 'pto_descarga': r[1]} for r in rows if r[0]]
        return jsonify({'contenedores': contenedores})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/contenedores/<num_contenedor>/piezas', methods=['GET'])
def get_piezas_contenedor(num_contenedor):
    try:
        db_path = os.path.join(os.getcwd(), 'datos_inventario.db')
        if not os.path.exists(db_path):
            return jsonify({'piezas': []})
            
        conexion = sqlite3.connect(db_path)
        cursor = conexion.cursor()
        cursor.execute("SELECT PLANO, DESC_PLANO, CANTIDAD FROM inventario WHERE NUM_CONTENEDOR = ? ORDER BY PLANO", (num_contenedor,))
        rows = cursor.fetchall()
        conexion.close()
        
        piezas = [{'plano': r[0], 'desc_plano': r[1], 'cantidad': int(r[2])} for r in rows]
        return jsonify({'piezas': piezas})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/piezas', methods=['GET'])
def get_piezas():
    try:
        db_path = os.path.join(os.getcwd(), 'datos_inventario.db')
        if not os.path.exists(db_path):
            return jsonify({'piezas': [], 'columnas': []})

        conexion = sqlite3.connect(db_path)
        cursor   = conexion.cursor()

        # Verificar si la tabla piezas existe
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='piezas'")
        if not cursor.fetchone():
            conexion.close()
            return jsonify({'piezas': [], 'columnas': []})

        cursor.execute("SELECT * FROM piezas ORDER BY Ref")
        rows    = cursor.fetchall()
        columnas = [description[0] for description in cursor.description]
        conexion.close()

        piezas = [dict(zip(columnas, row)) for row in rows]
        return jsonify({'piezas': piezas, 'columnas': columnas})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':

    # 1. Creamos la ventana apuntando a la dirección local del puerto 7777
    window = webview.create_window('Mi Aplicación Ejecutable', 'http://127.0.0.1:7777', fullscreen=True)
    
    # 2. Corremos Flask en un hilo separado para que no bloquee la ventana de la app
    threading.Thread(target=lambda: app.run(port=7777, debug=False, use_reloader=False), daemon=True).start()
    
    # 3. Arrancamos la interfaz gráfica
    webview.start()