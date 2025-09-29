import pandas as pd
import os
from openpyxl.utils import get_column_letter

# --- Configuración ---
carpeta_reportes = 'reportes_mensuales'
archivo_salida = 'resumen_consolidado.xlsx'
nombre_de_la_hoja = 'REPORTE FACTURACIÓN' 

# --- AJUSTES DE LECTURA DE LA TABLA ---
filas_a_saltar_al_inicio = 1 
numero_de_filas_a_leer = 29

# --- INICIO DE LA LÓGICA DEL SCRIPT ---
lista_de_datos = []
print(f"Buscando archivos de Excel en la carpeta: '{carpeta_reportes}'...")

for archivo in os.listdir(carpeta_reportes):
    if archivo.endswith(('.xlsx', '.xls')) and not archivo.startswith('~$'):
        ruta_completa = os.path.join(carpeta_reportes, archivo)
        print(f"Leyendo la tabla de datos del archivo: {archivo}...")
        try:
            # Leemos las 2 primeras filas como un encabezado jerárquico
            df_mensual = pd.read_excel(
                ruta_completa,
                sheet_name=nombre_de_la_hoja,
                skiprows=filas_a_saltar_al_inicio,
                nrows=numero_de_filas_a_leer,
                header=[0, 1]
            )
            
            # --- NUEVA LÓGICA DE RENOMBRADO A PRUEBA DE ERRORES ---
            # 1. Creamos nombres únicos preliminares a partir de la jerarquía
            df_mensual.columns = ['_'.join(map(str, col)).strip() for col in df_mensual.columns.values]

            # 2. Creamos un "mapa" para traducir esos nombres a un formato final y limpio
            # Esto maneja las inconsistencias como 'Unnamed', espacios o errores de tipeo.
            nombres_finales = {}
            columnas_esperadas = [
                ('N°', 'N°'), ('Ciudad', 'Ciudad'), ('AGENCIA', 'Agencia'), ('Ubicación', 'Ubicación'),
                ('Departamento', 'Departamento'), ('Modelo', 'Modelo'), ('Serial', 'Serial'), ('IP', 'IP'),
                ('IMPRESIONES B/N_Contador Inicial', 'Inicial Imp. B/N'), ('IMPRESIONES B/N_Contador Final', 'Final Imp. B/N'), ('IMPRESIONES B/N_TOTAL', 'Total Imp. B/N'),
                ('COPIAS B/N_Contador Inicial', 'Inicial Copias B/N'), ('COPIAS B/N_Contador Final', 'Final Copias B/N'), ('COPIAS B/N_TOTAL', 'Total Copias B/N'),
                ('IMPRESIONES COLOR_Contador Inicial', 'Inicial Imp. Color'), ('IMPRESIONES COLOR_Contador Final', 'Final Imp. Color'), ('IMPRESIONES COLOR_TOTAL', 'Total Imp. Color'),
                ('COPIAS COLOR_Contador Inicial', 'Inicial Copias Color'), ('COPIAS COLOR_Contador Final', 'Final Copias Color'), ('COPIAS COLOR_TOTAL', 'Total Copias Color')
            ]
            
            # Asignamos los nombres limpios basados en la posición, lo que es más robusto
            for i, (nombre_esperado, nombre_limpio) in enumerate(columnas_esperadas):
                if i < len(df_mensual.columns):
                    nombres_finales[df_mensual.columns[i]] = nombre_limpio
            
            df_mensual.rename(columns=nombres_finales, inplace=True)
            # --- FIN DE LA NUEVA LÓGICA ---

            nombre_periodo = os.path.splitext(archivo)[0] 
            df_mensual['Periodo'] = nombre_periodo
            lista_de_datos.append(df_mensual)
        
        except Exception as e:
            print(f"  -> ADVERTENCIA: No se pudo leer el archivo '{archivo}'. Error: {e}")

if not lista_de_datos:
    print("Error: No se procesó ningún archivo con éxito.")
else:
    print("Consolidando todos los datos...")
    df_consolidado = pd.concat(lista_de_datos, ignore_index=True)
    
    print("Ordenando los datos cronológicamente...")
    meses_es = {
        'Enero': 'January', 'Febrero': 'February', 'Marzo': 'March', 'Abril': 'April',
        'Mayo': 'May', 'Junio': 'June', 'Julio': 'July', 'Agosto': 'August',
        'Septiembre': 'September', 'Octubre': 'October', 'Noviembre': 'November', 'Diciembre': 'December'
    }
    periodo_traducido = df_consolidado['Periodo'].str.split().str[0].map(meses_es)
    año_str = df_consolidado['Periodo'].str.split().str[1]
    periodo_en_ingles = periodo_traducido + ' ' + año_str
    df_consolidado['FechaOrden'] = pd.to_datetime(periodo_en_ingles)
    df_consolidado = df_consolidado.sort_values('FechaOrden')
    df_consolidado = df_consolidado.drop(columns=['FechaOrden'])

    print("Limpiando datos y recalculando totales para asegurar la precisión...")
    columnas_contadores = [
        'Inicial Imp. B/N', 'Final Imp. B/N', 'Inicial Copias B/N', 'Final Copias B/N',
        'Inicial Imp. Color', 'Final Imp. Color', 'Inicial Copias Color', 'Final Copias Color'
    ]
    for col in columnas_contadores:
        if col in df_consolidado.columns:
            df_consolidado[col] = pd.to_numeric(df_consolidado[col], errors='coerce').fillna(0)
    
    df_consolidado['Total Imp. B/N'] = df_consolidado['Final Imp. B/N'] - df_consolidado['Inicial Imp. B/N']
    df_consolidado['Total Copias B/N'] = df_consolidado['Final Copias B/N'] - df_consolidado['Inicial Copias B/N']
    df_consolidado['Total Imp. Color'] = df_consolidado['Final Imp. Color'] - df_consolidado['Inicial Imp. Color']
    df_consolidado['Total Copias Color'] = df_consolidado['Final Copias Color'] - df_consolidado['Inicial Copias Color']
    
    print("Calculando totales anuales...")
    resumen_anual = {'Concepto': [], 'Total Anual': []}
    columnas_totales = ['Total Imp. B/N', 'Total Copias B/N', 'Total Imp. Color', 'Total Copias Color']
    for col in columnas_totales:
        resumen_anual['Concepto'].append(col.replace('Total', 'Total de'))
        resumen_anual['Total Anual'].append(df_consolidado[col].sum())
    
    df_resumen = pd.DataFrame(resumen_anual)

    print("Guardando el archivo final con dos hojas...")
    with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
        df_consolidado.to_excel(writer, index=False, sheet_name='Resumen Detallado')
        df_resumen.to_excel(writer, index=False, sheet_name='Totales Anuales')
        
        worksheet1 = writer.sheets['Resumen Detallado']
        for column_cells in worksheet1.columns:
            max_length = max(len(str(cell.value)) for cell in column_cells)
            adjusted_width = max_length + 2
            worksheet1.column_dimensions[column_cells[0].column_letter].width = adjusted_width

        worksheet2 = writer.sheets['Totales Anuales']
        for column_cells in worksheet2.columns:
            max_length = max(len(str(cell.value)) for cell in column_cells)
            adjusted_width = max_length + 2
            worksheet2.column_dimensions[column_cells[0].column_letter].width = adjusted_width
    
    print("-" * 30)
    print(f"✅ ¡Proceso completado! El archivo está guardado con dos hojas en: '{archivo_salida}'!")