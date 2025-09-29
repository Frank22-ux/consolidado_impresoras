import pandas as pd
from docx import Document
from docx.shared import Inches
from sklearn.linear_model import LinearRegression
import numpy as np

# --- Configuración ---
archivo_entrada = 'resumen_consolidado.xlsx'
archivo_salida_word = 'Reporte_Analisis_Consumo_Anual_Completo.docx'

# --- 1. Cargar y Preparar los Datos ---
print(f"Leyendo la hoja 'Resumen Detallado' del archivo: {archivo_entrada}...")
df = pd.read_excel(archivo_entrada, sheet_name='Resumen Detallado')

print("Convirtiendo periodos a fechas...")
meses_es = {
    'Enero': 'January', 'Febrero': 'February', 'Marzo': 'March', 'Abril': 'April',
    'Mayo': 'May', 'Junio': 'June', 'Julio': 'July', 'Agosto': 'August',
    'Septiembre': 'September', 'Octubre': 'October', 'Noviembre': 'November', 'Diciembre': 'December'
}
periodo_traducido = df['Periodo'].str.split().str[0].map(meses_es)
año_str = df['Periodo'].str.split().str[1]
periodo_en_ingles = periodo_traducido + ' ' + año_str
df['Fecha'] = pd.to_datetime(periodo_en_ingles)
df['Año'] = df['Fecha'].dt.year

# --- 2. Preparar el Documento de Word y el Título Dinámico ---
documento = Document()
años_en_datos = sorted(df['Año'].unique())
if len(años_en_datos) == 1:
    titulo_principal = f'Análisis de Consumo y Proyecciones {años_en_datos[0]}'
else:
    titulo_principal = f'Análisis de Consumo y Proyecciones {años_en_datos[0]}-{años_en_datos[-1]}'
documento.add_heading(titulo_principal, level=1)

columnas_analisis = ['Total Imp. B/N', 'Total Copias B/N', 'Total Imp. Color', 'Total Copias Color']

# --- 3. Bucle Principal para Analizar Cada Año por Separado ---
for año in años_en_datos:
    print(f"\n--- Procesando datos para el año: {año} ---")
    documento.add_heading(f'Resultados para el Año {año}', level=2)
    
    df_año = df[df['Año'] == año].copy()

    # Análisis de Promedios
    print(f"Calculando promedios para {año}...")
    analisis_detallado = df_año.groupby(['Agencia', 'Departamento'])[columnas_analisis].mean().round(0)
    analisis_detallado.reset_index(inplace=True)
    documento.add_paragraph(f"La siguiente tabla muestra el promedio mensual de consumo para {año}.")
    tabla_analisis = documento.add_table(rows=1, cols=len(analisis_detallado.columns))
    tabla_analisis.style = 'Table Grid'
    for i, nombre_columna in enumerate(analisis_detallado.columns):
        tabla_analisis.cell(0, i).text = str(nombre_columna)
    for index, fila in analisis_detallado.iterrows():
        celdas_fila = tabla_analisis.add_row().cells
        for i, valor in enumerate(fila):
            celdas_fila[i].text = f"{valor:,.0f}".replace(',', '.') if isinstance(valor, (int, float)) else str(valor)

    # Proyecciones
    print(f"Realizando proyecciones basadas en los datos de {año}...")
    documento.add_heading(f'Proyecciones a Futuro (Basado en {año})', level=3)
    
    # Proyección por Tipo de Consumo
    proyecciones_por_tipo = {}
    for columna in columnas_analisis:
        consumo_mensual = df_año.groupby('Fecha')[columna].sum().reset_index()
        if len(consumo_mensual) > 1:
            consumo_mensual['Mes_Num'] = np.arange(len(consumo_mensual))
            modelo = LinearRegression()
            modelo.fit(consumo_mensual[['Mes_Num']], consumo_mensual[columna])
            ultimo_mes_num = consumo_mensual['Mes_Num'].max()
            meses_futuros = np.array([[ultimo_mes_num + 6], [ultimo_mes_num + 12]])
            predicciones = modelo.predict(meses_futuros).round(0)
            if predicciones[0] >= 0:
                proyecciones_por_tipo[columna] = predicciones
    
    if proyecciones_por_tipo:
        documento.add_paragraph("Proyección por tipo de consumo (General):", style='Heading 4')
        tabla_proy_tipo = documento.add_table(rows=1, cols=3)
        tabla_proy_tipo.style = 'Table Grid'
        tabla_proy_tipo.cell(0, 0).text = 'Tipo de Consumo'
        tabla_proy_tipo.cell(0, 1).text = 'Proyección a 6 Meses'
        tabla_proy_tipo.cell(0, 2).text = 'Proyección a 1 Año'
        for tipo, proyecciones in proyecciones_por_tipo.items():
            celdas_fila = tabla_proy_tipo.add_row().cells
            celdas_fila[0].text = str(tipo)
            celdas_fila[1].text = f"{int(proyecciones[0]):,}".replace(',', '.')
            celdas_fila[2].text = f"{int(proyecciones[1]):,}".replace(',', '.')

    # --- LÓGICA REINCORPORADA: Proyección por Agencia y por Tipo ---
    proyecciones_por_agencia_detallado = {}
    for agencia in df_año['Agencia'].unique():
        df_agencia = df_año[df_año['Agencia'] == agencia]
        for columna in columnas_analisis:
            consumo_mensual = df_agencia.groupby('Fecha')[columna].sum().reset_index()
            if len(consumo_mensual) > 1:
                consumo_mensual['Mes_Num'] = np.arange(len(consumo_mensual))
                modelo = LinearRegression()
                modelo.fit(consumo_mensual[['Mes_Num']], consumo_mensual[columna])
                ultimo_mes_num = consumo_mensual['Mes_Num'].max()
                meses_futuros = np.array([[ultimo_mes_num + 6], [ultimo_mes_num + 12]])
                predicciones = modelo.predict(meses_futuros).round(0)
                if predicciones[0] >= 0:
                    if agencia not in proyecciones_por_agencia_detallado:
                        proyecciones_por_agencia_detallado[agencia] = {}
                    proyecciones_por_agencia_detallado[agencia][columna] = predicciones

    if proyecciones_por_agencia_detallado:
        documento.add_paragraph("Proyección por agencia y tipo de consumo:", style='Heading 4')
        tabla_proy_agencia = documento.add_table(rows=1, cols=4)
        tabla_proy_agencia.style = 'Table Grid'
        tabla_proy_agencia.cell(0, 0).text = 'Agencia'
        tabla_proy_agencia.cell(0, 1).text = 'Tipo de Consumo'
        tabla_proy_agencia.cell(0, 2).text = 'Proyección a 6 Meses'
        tabla_proy_agencia.cell(0, 3).text = 'Proyección a 1 Año'
        for agencia, proyecciones_tipos in proyecciones_por_agencia_detallado.items():
            for tipo, proyecciones in proyecciones_tipos.items():
                celdas_fila = tabla_proy_agencia.add_row().cells
                celdas_fila[0].text = str(agencia)
                celdas_fila[1].text = str(tipo)
                celdas_fila[2].text = f"{int(proyecciones[0]):,}".replace(',', '.')
                celdas_fila[3].text = f"{int(proyecciones[1]):,}".replace(',', '.')

# --- 4. Guardar el Documento ---
documento.save(archivo_salida_word)
print("-" * 30)
print(f"✅ ¡Reporte anual generado exitosamente en '{archivo_salida_word}'!")