import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.axis import ChartLines 
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.drawing.text import CharacterProperties
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import tkinter as tk
from tkinter import filedialog
import os
from tkinter import simpledialog, messagebox # <--- Añadimos messagebox

# --- CONFIGURACIÓN AUTOMATIZADA ---

# 1. Preparar la ventana emergente
root = tk.Tk()
root.withdraw() 
root.attributes('-topmost', True) 

# --- CONFIGURACIÓN AUTOMATIZADA ---

# 1. Preparar la ventana emergente
root = tk.Tk()
root.withdraw() 
root.attributes('-topmost', True) 

# --- NUEVO: PESTAÑA DE BIENVENIDA Y GUÍA (ONBOARDING) ---
mensaje_guia = (
    "¡Bienvenido al Limpiador de Registros!\n\n"
    "Esta herramienta procesa los datos brutos del vatímetro para extraer las métricas "
    "eléctricas clave, generar gráficas analíticas y clasificar el equipo.\n\n"
    "📚 A TENER EN CUENTA ANTES DE PROCEDER:\n\n"
    "👉 1. UMBRAL DE POTENCIA (W):\n"
    "Define el instante de inicio (t0). El algoritmo escanea el registro y establece el "
    "punto cero en el momento exacto en el que el consumo CAE por debajo de este valor "
    "(el programa ignorará automáticamente los 0W iniciales de la máquina apagada).\n\n"
    "👉 2. VENTANA DE CÁLCULO (Segundos):\n"
    "Establece el periodo de integración a partir del t0. Este parámetro es crítico, "
    "ya que el programa calcula matemáticamente la Energía Disipada (Julios) multiplicando "
    "la Potencia Media de esta franja por el tiempo de cálculo especificado.\n\n"
    "👉 3. FILTRADO ESTADÍSTICO DE RUIDO:\n"
    "Las métricas de Media, Mediana y Mínimo pueden verse falseadas por el ruido "
    "instrumental del equipo o periodos de inactividad a 0W. El filtro opcional permite "
    "purgar estos valores para que la estadística refleje únicamente el consumo real.\n\n"
    "Haz clic en Aceptar para importar el registro .CSV."
)

# Mostramos la ventana de información ANTES de hacer nada más
messagebox.showinfo("Guía de Usuario - Instrucciones", mensaje_guia)
# ---------------------------------------------------------

print("Esperando a que selecciones un archivo...")

# 2. Abrir el explorador de archivos para elegir el CSV
archivos_seleccionados = filedialog.askopenfilenames(
    title="Selecciona el registro del vatímetro",
    filetypes=[("Archivos CSV", "*.csv;*.CSV"), ("Todos los archivos", "*.*")]
)

if not archivos_seleccionados:
    print("Operación cancelada. No se seleccionó ningún archivo.")
    exit()

for archivo_csv in archivos_seleccionados:
    # 3. Crear el nombre dinámico del Excel
    nombre_base = os.path.splitext(os.path.basename(archivo_csv))[0]
    carpeta_origen = os.path.dirname(archivo_csv)
    archivo_excel = os.path.join(carpeta_origen, f"{nombre_base}_potencia.xlsx")

    # ---------------------

        # --- NUEVO: PREGUNTAR UMBRALES Y TIEMPO ---
    umbral_1 = simpledialog.askfloat("Umbral", f"Archivo: {nombre_base}\nUmbral de potencia (W) para Elemento 1:", parent=root, minvalue=0.0)
    umbral_2 = simpledialog.askfloat("Umbral", f"Archivo: {nombre_base}\nUmbral de potencia (W) para Elemento 2:", parent=root, minvalue=0.0)
    umbral_3 = simpledialog.askfloat("Umbral", f"Archivo: {nombre_base}\nUmbral de potencia (W) para Elemento 3:", parent=root, minvalue=0.0)
    tiempo_calc = simpledialog.askfloat("Tiempo", f"Archivo: {nombre_base}\nTiempo para el cálculo en segundos:", parent=root, minvalue=0.0)
    
    # Si el usuario cierra alguna ventana o le da a cancelar
    if None in (umbral_1, umbral_2, umbral_3, tiempo_calc):
        print("Operación cancelada para {nombre_base}. ")
        continue

    #preguntar si se ignoran los ceros
    ignorar_ceros = messagebox.askyesno("Filtro de Ceros", f"Archivo: {nombre_base}\n\n¿Deseas IGNORAR los valores '0' para que no alteren la Media, Mediana y Mínimo?\n\n(Sí = Filtrar los ceros / No = Tenerlos en cuenta)")

    umbrales = {1: umbral_1, 2: umbral_2, 3: umbral_3}
    # ----------------------------------------

    print(f"Leyendo el archivo: {archivo_csv}...")

    # 1. LEER Y SEPARAR LAS PARTES DEL CSV
    info_data = []
    settings_data = []
    data_start_idx = 0

    with open(archivo_csv, 'r', encoding='latin1') as f:
        lines = f.readlines()
        section = "info"
        for i, line in enumerate(lines):
            line = line.strip()
            if not line: continue
            
            if line == '"<Settings>"': section = "settings"; continue
            if line == '"<Settings (Harmonics)>"': section = "harmonics"; continue
            
            if line.startswith('"Store No."') or line.startswith('Store No.'):
                data_start_idx = i
                break
                
            parts = [p.strip('"') for p in line.split('",')]
            if len(parts) >= 2:
                if section == "info": info_data.append((parts[0].strip('"'), parts[1].strip('"')))
                elif section == "settings": settings_data.append([p.strip('"') for p in parts])

    # Leer la tabla de datos principal
    df = pd.read_csv(archivo_csv, skiprows=data_start_idx, encoding='latin1')

    # --- NUEVO: 2. CREAR LAS 4 TABLAS DE DATOS (MAX, MIN, MEDIO, MEDIANA) ---
    summary_max_data = []
    summary_min_data = []
    summary_mean_data = []
    summary_median_data = [] # <--- NUEVO
    
    franjas_indices = {} 
    avisos_falta_datos = []
    
    col_tiempo = None
    for c in df.columns:
        if 'time' in str(c).lower():
            col_tiempo = c
            break
            
    if col_tiempo:
        # Le añadimos format='mixed' para que no se queje al adivinar el formato
        tiempos = pd.to_datetime(df[col_tiempo], errors='coerce', format='mixed')
        
        # --- NUEVO: Convertimos los textos a formato "Reloj puro" ---
        # Esto soluciona de golpe que Excel no te mostrara los números del eje X
        df[col_tiempo] = tiempos.dt.time
    else:
        tiempos = None

    for i in [1, 2, 3]:
        cols = {
            "Voltaje Urms (V)": f"Urms-E{i}",
            "Corriente Irms (A)": f"Irms-E{i}",
            "Potencia Activa P (W)": f"P-E{i}",
            "Potencia Aparente S (VA)": f"S-E{i}"
        }
        
        stats_max = {"Parámetro": f"Elemento {i}"}
        stats_min = {"Parámetro": f"Elemento {i}"}
        stats_mean = {"Parámetro": f"Elemento {i}"}
        stats_median = {"Parámetro": f"Elemento {i}"} # <--- NUEVO
        
        col_potencia = f"P-E{i}"
        start_idx = df.index[0] if not df.empty else 0 
        
        if col_potencia in df.columns:
            vals_potencia = pd.to_numeric(df[col_potencia], errors='coerce').abs()
            
            # --- NUEVO ESCUDO: IGNORAR LOS 0W INICIALES AL BUSCAR EL UMBRAL ---
            # Le decimos que busque valores menores al umbral, pero EXIGIENDO que 
            # sean mayores a 0.0001W. Así evitamos que empiece el ensayo en un 0W
            # de cuando la máquina estaba apagada.
            validos = vals_potencia[(vals_potencia < umbrales[i]) & (vals_potencia > 0.0001)]
            
            if not validos.empty:
                start_idx = validos.index[0] 
            # ------------------------------------------------------------------
                
        end_idx = start_idx
        if tiempos is not None and pd.notna(tiempos.loc[start_idx]):
            t_inicio = tiempos.loc[start_idx]
            dif_segundos = (tiempos.loc[start_idx:] - t_inicio).dt.total_seconds()
            dif_segundos = dif_segundos.apply(lambda x: x + 86400 if x < 0 else x)
            dentro_del_tiempo = dif_segundos[dif_segundos <= tiempo_calc]
            
            if not dentro_del_tiempo.empty:
                end_idx = dentro_del_tiempo.index[-1] 
            
            if dif_segundos.max() < (tiempo_calc - 1):
                # En lugar de mostrar la ventana aquí, lo guardamos en la lista
                avisos_falta_datos.append(f"- Elemento {i}: Se apagó a los {round(dif_segundos.max(), 1)}s")
        else:
            end_idx = start_idx + int(tiempo_calc) - 1
            if end_idx > df.index[-1]: end_idx = df.index[-1]
            
        franjas_indices[i] = (start_idx, end_idx) 
        
        for metric_name, col_name in cols.items():
            if col_name in df.columns:
                vals = pd.to_numeric(df[col_name], errors='coerce').abs()
                vals_finales = vals.loc[start_idx:end_idx] 
                
                # --- NUEVO: SI EL USUARIO DIJO "SÍ", VOLVEMOS INVISIBLES LOS CEROS ---
                if ignorar_ceros:
                    # Volvemos a ignorar el ruido microscópico (valores menores a 0.0001)
                    vals[vals < 0.0001] = np.nan
                    vals_finales[vals_finales < 0.0001] = np.nan
                # ---------------------------------------------------------------------

                stats_max[metric_name] = round(vals.max(), 4)
                stats_min[metric_name] = round(vals.min(),4)
                stats_mean[metric_name] = round(vals_finales.mean(), 4)
                stats_median[metric_name] = round(vals.median(), 4) # <--- NUEVO
            else:
                stats_max[metric_name] = "No detectado"
                stats_min[metric_name] = "No detectado"
                stats_mean[metric_name] = "No detectado"
                stats_median[metric_name] = "No detectado" # <--- NUEVO
                
        if col_potencia in df.columns:
            vals_pot_franja = vals_potencia.loc[start_idx:end_idx]
            stats_max["Energía Disipada (J)"] = "N/A"
            stats_min["Energía Disipada (J)"] = "N/A"
            stats_median["Energía Disipada (J)"] = "N/A" # <--- NUEVO
            stats_mean["Energía Disipada (J)"] = round(vals_pot_franja.mean() * tiempo_calc, 4)
        else:
            stats_max["Energía Disipada (J)"] = "No detectado"
            stats_min["Energía Disipada (J)"] = "No detectado"
            stats_median["Energía Disipada (J)"] = "No detectado" # <--- NUEVO
            stats_mean["Energía Disipada (J)"] = "No detectado"
                
        summary_max_data.append(stats_max)
        summary_min_data.append(stats_min)
        summary_mean_data.append(stats_mean)
        summary_median_data.append(stats_median) # <--- NUEVO

    # --- NUEVO: MOSTRAR UN SOLO AVISO SI FALTARON DATOS ---
    if avisos_falta_datos:
        mensaje = f"Has pedido un ensayo de {tiempo_calc}s, pero faltan datos en el archivo {nombre_base}:\n\n" + "\n".join(avisos_falta_datos)
        messagebox.showwarning("¡Atención: Faltan datos!", mensaje)
    # -----------------------------------------------------

    df_max = pd.DataFrame(summary_max_data).set_index("Parámetro").T
    df_min = pd.DataFrame(summary_min_data).set_index("Parámetro").T
    df_mean = pd.DataFrame(summary_mean_data).set_index("Parámetro").T
    df_median = pd.DataFrame(summary_median_data).set_index("Parámetro").T # <--- NUEVO

    # 3. CREAR EL EXCEL Y DARLE FORMATO
    wb = Workbook()
    ws_resumen = wb.active
    ws_resumen.title = "Resumen"

    # Estilos visuales
    header_font = Font(bold=True, color="FFFFFF")
    align_center = Alignment(horizontal="center", vertical="center")
    borde_fino = Border(
        left=Side(border_style="thin", color="000000"),
        right=Side(border_style="thin", color="000000"),
        top=Side(border_style="thin", color="000000"),
        bottom=Side(border_style="thin", color="000000")
    )

    def escribir_tabla(ws, datos, start_row, start_col, titulo=None, color_hex="4F81BD", formato_numero='0.00E+00', usar_absoluto=False):
        if titulo:
            ws.cell(row=start_row, column=start_col, value=titulo).font = Font(bold=True, size=12)
            start_row += 1
        
        if isinstance(datos, pd.DataFrame):
            filas = dataframe_to_rows(datos, index=False, header=True)
        else:
            filas = datos
            
        fondo = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
        fondo_gris = PatternFill(start_color="F3F3F3", end_color="EAEAEA", fill_type="solid")
           
        for r_idx, row in enumerate(filas, start=start_row):
            for c_idx, value in enumerate(row, start=start_col):
                if value is None: continue
                
                if usar_absoluto and isinstance(value, (int, float, np.number)):
                    value = abs(value)
                
                celda = ws.cell(row=r_idx, column=c_idx, value=value)
                celda.border = borde_fino
                
                if r_idx == start_row: 
                    celda.font = header_font
                    celda.fill = fondo
                    celda.alignment = align_center
                else:
                    if c_idx == start_col:
                        celda.fill = fondo_gris

                    if isinstance(value, (int, float, np.number)):
                        celda.number_format = formato_numero


    # --- NUEVO: REORGANIZACIÓN DEL EXCEL ---

    # 1. Pegamos la Info en su sitio original (Fila 2, Columna 2)
    info_df = pd.DataFrame(info_data, columns=["Atributo", "Valor"])
    escribir_tabla(ws_resumen, info_df, 2, 2, "Información del Equipo", color_hex="92C98F")

    # 2. Plantamos la LEYENDA arriba a la derecha (Fila 2, Columna 5)
    color_ps1 = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid") 
    color_ps2 = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid") 
    color_ps3 = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid") 
    
    ws_resumen.cell(row=2, column=5, value="Clasificación (Potencia):").font = Font(bold=True, size=12)
    ws_resumen.cell(row=3, column=5, value="PS1: <= 15W").fill = color_ps1
    ws_resumen.cell(row=4, column=5, value="PS2: > 15W y <= 100W").fill = color_ps2
    ws_resumen.cell(row=5, column=5, value="PS3: > 100W").fill = color_ps3
    
    nota_texto = "*Nota: Criterio orientativo. Decidir finalmente por potencia disponible medida según 62368-1, no solo por potencia nominal. Tener en cuenta la duración temporal y el peor caso de ensayo."
    celda_nota = ws_resumen.cell(row=7, column=5, value=nota_texto)
    celda_nota.font = Font(italic=True, size=9, color="595959")

    # 3. Pegamos las 4 TABLAS en formato cuadrícula 2x2
    # Fila 10 (Arriba): Máximos y Medios
    escribir_tabla(ws_resumen, df_max.reset_index(), 10, 2, "Valores Máximos", color_hex="D693DB", formato_numero='0.0000', usar_absoluto=True)
    escribir_tabla(ws_resumen, df_mean.reset_index(), 10, 7, "Valores Medios en franja umbral (Promedios)", color_hex="D693DB", formato_numero='0.0000', usar_absoluto=True)
    
    # Fila 18 (Abajo): Mínimos y Mediana
    escribir_tabla(ws_resumen, df_min.reset_index(), 18, 2, "Valores Mínimos", color_hex="D693DB", formato_numero='0.0000', usar_absoluto=True)
    escribir_tabla(ws_resumen, df_median.reset_index(), 18, 7, "Valores Medianos (Mediana)", color_hex="D693DB", formato_numero='0.0000', usar_absoluto=True)

    # 4. COLOREAR LA POTENCIA EN LAS 4 TABLAS
    # Le pasamos las coordenadas de inicio exactas de nuestras 4 tablas (Fila, Columna)
    ubicaciones_tablas = [(10, 2), (10, 7), (18, 2), (18, 7)] 
    
    for f_inicio, c_inicio in ubicaciones_tablas:
        fila_potencia = 0
        # Buscamos la "Potencia Activa P (W)" dinámicamente debajo del título de cada tabla
        for r in range(f_inicio + 1, f_inicio + 8):
            if ws_resumen.cell(row=r, column=c_inicio).value == "Potencia Activa P (W)":
                fila_potencia = r
                break
                
        if fila_potencia != 0:
            for col in range(c_inicio + 1, c_inicio + 4): # Evaluamos los 3 elementos
                valor = ws_resumen.cell(row=fila_potencia, column=col).value
                if isinstance(valor, (int, float, np.number)):
                    p_activa = abs(valor) 
                    if p_activa <= 15:
                        ws_resumen.cell(row=fila_potencia, column=col).fill = color_ps1
                    elif p_activa <= 100:
                        ws_resumen.cell(row=fila_potencia, column=col).fill = color_ps2
                    else:
                        ws_resumen.cell(row=fila_potencia, column=col).fill = color_ps3

    # 5. BAJAMOS LA CONFIGURACIÓN (Lo bajamos a la fila 28 para que quepan las tablas nuevas)
    escribir_tabla(ws_resumen, settings_data, 28, 2, "Configuración (Settings)", color_hex="A5CDF0")

    # Creamos la segunda pestaña
    ws_datos = wb.create_sheet(title="Datos Registro")
    escribir_tabla(ws_datos, df, 1, 1, None, color_hex="A5CED6")
    
    # --- NUEVO: COLOREAR LA FRANJA DE DATOS VÁLIDOS ---
    color_franja = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # Verde clarito
    
    # 1. Identificar a qué elemento pertenece cada columna de Excel
    cols_elemento = {1: [], 2: [], 3: []}
    for c in range(1, ws_datos.max_column + 1):
        val = ws_datos.cell(row=1, column=c).value
        if val and isinstance(val, str):
            if "-E1" in val: cols_elemento[1].append(c)
            elif "-E2" in val: cols_elemento[2].append(c)
            elif "-E3" in val: cols_elemento[3].append(c)

    # 2. Pintar las celdas EXTREMADAMENTE exactas de la franja de tiempo
    for num_elemento, franja in franjas_indices.items():
        s_idx, e_idx = franja
        
        # En Excel la fila 1 es la cabecera, así que el dato 0 es la fila 2
        fila_inicio = s_idx + 2 
        fila_fin = e_idx + 2
        
        # Le ponemos un tope para que no intente pintar fuera del documento si hay algún error
        fila_fin = min(fila_fin, ws_datos.max_row)
        
        for r in range(fila_inicio, fila_fin + 1):
            for c_idx in cols_elemento[num_elemento]:
                ws_datos.cell(row=r, column=c_idx).fill = color_franja
    # --------------------------------------------------
         
    # Ajustar el ancho de las columnas
    for ws in wb.worksheets:
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                
                # --- NUEVO: Ignoramos la celda de la nota larga (Fila 7, Columna 5) al medir
                if ws.title == "Resumen" and cell.row == 7 and cell.column == 5:
                    continue
                    
                try:
                    if cell.value is not None:
                        if isinstance(cell.value, (int, float, np.number)):
                            longitud_celda = 10
                        else:
                            longitud_celda = len(str(cell.value))
                            
                        if longitud_celda > max_length:
                            max_length = longitud_celda
                except:
                    pass
            
            ws.column_dimensions[column].width = min(max_length + 3, 40)

    # --- NUEVO: CREAR PESTAÑA DE GRÁFICAS ---
    ws_graficas = wb.create_sheet(title="Gráficas")

    # 1. Buscar en qué columna exacta de Excel cayó el Tiempo y las Potencias
    col_idx_tiempo = None
    cols_potencia = {1: None, 2: None, 3: None}
    
    for c_idx, col_name in enumerate(df.columns, start=1):
        if col_tiempo and col_name == col_tiempo:
            col_idx_tiempo = c_idx
        elif col_name == 'P-E1': cols_potencia[1] = c_idx
        elif col_name == 'P-E2': cols_potencia[2] = c_idx
        elif col_name == 'P-E3': cols_potencia[3] = c_idx

   # 2. Dibujar una gráfica de curva continua por cada elemento
    fila_destino = 2 
    max_r = ws_datos.max_row
    
    for i in [1, 2, 3]:
        if cols_potencia[i] is not None:
            chart = ScatterChart()
            chart.scatterStyle = 'line' 
            chart.title = f"Potencia Activa P (W) - Elemento {i}"
            
            # --- NUEVO: HACEMOS EL TÍTULO MÁS GRANDE ---
            # sz=1400 es tamaño 14, b=True es Negrita
            chart.title.tx.rich.p[0].pPr.defRPr = CharacterProperties(sz=1400, b=True) 
            
            # Tamaño de la gráfica
            chart.width = 18  
            chart.height = 10 
            
            # Cuadrícula Gris
            chart.x_axis.majorGridlines = ChartLines()
            chart.y_axis.majorGridlines = ChartLines()
            
            # Encender las casillas de los ejes
            chart.x_axis.delete = False 
            chart.y_axis.delete = False 
            
            # Mandar los ejes a los bordes
            chart.x_axis.crosses = "min" 
            chart.y_axis.crosses = "min" 
            
            # Formato de los números
            chart.x_axis.number_format = 'hh:mm:ss'
            chart.x_axis.number_format_linked = False 
            chart.y_axis.number_format = 'General'
            chart.y_axis.number_format_linked = False 
            
            chart.x_axis.tickLblPos = "nextTo" 
            chart.y_axis.tickLblPos = "nextTo" 

            # --- EL TRUCO INFALIBLE: ALEJAR TEXTOS CON SALTOS DE LÍNEA ---
            # Los \n son "Intros". Obligan a Excel a crear espacio vacío.
            chart.x_axis.title = "\n\nTiempo" 
            chart.y_axis.title = "Potencia (W)\n\n\n" 
            # -------------------------------------------------------------
            
            # Ponemos los ejes en negrita para que destaquen más
            chart.x_axis.title.tx.rich.p[0].pPr.defRPr = CharacterProperties(b=True)
            chart.y_axis.title.tx.rich.p[0].pPr.defRPr = CharacterProperties(b=True)
            
            datos_y = Reference(ws_datos, min_col=cols_potencia[i], min_row=2, max_row=max_r)
            
            if col_idx_tiempo is not None:
                datos_x = Reference(ws_datos, min_col=col_idx_tiempo, min_row=2, max_row=max_r)
                serie = Series(datos_y, xvalues=datos_x, title_from_data=False)
            else:
                serie = Series(datos_y, title_from_data=False)
                
            chart.legend = None 
            
            # Línea Azul Profesional
            serie.marker.symbol = "none" 
            serie.graphicalProperties.line.solidFill = "0070C0" 
            serie.graphicalProperties.line.width = 20000 
            
            chart.series.append(serie)
            
            ws_graficas.add_chart(chart, f"B{fila_destino}")
            
            fila_destino += 22 
    # ----------------------------------------
            
    # Guardar
    wb.save(archivo_excel)
    print(f"¡Listo! Se ha guardado el archivo bonito como: {archivo_excel}")