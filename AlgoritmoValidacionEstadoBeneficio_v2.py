###
### Autor: Kramer con k de keppler 
### Fecha: 28/08/2024
###
import os
import pandas as pd

# Configuración de opciones de visualización
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

# Verifica la existencia del archivo en la ruta especificada
file_path = '/content/PIAM_UNICAUCA.xlsx'

if not os.path.isfile(file_path):
    raise FileNotFoundError(f"{file_path} no encontrado.")
else:
    print(f"Archivo {file_path} encontrado.")

# Abre el archivo en modo binario para verificar problemas de acceso
try:
    with open(file_path, 'rb') as f:
        print(f"Archivo {file_path} abierto satisfactoriamente en modo binario.")
except OSError as e:
    print(f"Error al abrir el archivo {file_path}: {e}")

# Función para cargar las hojas de trabajo desde el archivo Excel
def load_excel_sheets(file_path):
    xls = pd.ExcelFile(file_path, engine='openpyxl')
    sheet_names = xls.sheet_names
    print("Nombres de las hojas:", sheet_names)
    
    sheet_dic = {}
    for sheet_name in sheet_names:
        sheet_dic[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
    return sheet_dic

# Cargar los datos
sheet_dic = load_excel_sheets(file_path)

# Definir columnas por periodo
columnas_por_periodo = {
    'PIAM2021_1': ['BOLETA', 'RECURSOS APLICADOS'],
    'PIAM2021_2': ['BOLETA', 'ESTADO F', 'ESTADO'],
    'PIAM2022_1': ['BOLETA', 'ESTADO POLITICA', 'Criterio NO Acceso'],
    'PIAM2022_2': ['BOLETA', 'ESTADO', 'RESULTADO_VALIDACION'],
    'PIAM2023_1': ['RECIBO', 'ESTADO', 'NACIMIENTO'],
    'PIAM2023_2': ['RECIBO', 'ESTADO POLITICA', 'CRITERIOVAL21'],
    'PIAM2024_1DF': ['RECIBO', 'ESTADO CIVF', 'STT EJECUCION']
}

# Crear un DataFrame con todos los registros de consultaCaro
df_combined = sheet_dic.get('CONSULTACARO').copy()
df_combined['contador'] = 'No encontrado'
df_combined['estado_beneficio'] = None
df_combined['criterio_beneficio'] = None

# Función para realizar el cruce de datos
def realizar_cruce(consultaCaro, piam_df, columnas_a_traer, columna_boleta, cruce_id):
    consultaCaro['Documento'] = consultaCaro['Documento'].astype(str)
    piam_df[columna_boleta] = piam_df[columna_boleta].astype(str)
    
    # Realizar el cruce
    df_merged = pd.merge(consultaCaro, piam_df[columnas_a_traer], 
                         left_on='Documento', right_on=columna_boleta, how='left')
    
    # Actualizar la columna 'contador' con el identificador del cruce
    cruce_str = f'{cruce_id}'
    df_merged['contador'] = df_merged.apply(
        lambda row: cruce_str if pd.notna(row[columnas_a_traer[1]]) else row['contador'], 
        axis=1
    )
    
    # Actualizar las columnas 'estado_beneficio' y 'criterio_beneficio'
    if len(columnas_a_traer) > 1:
        if 'estado_beneficio' in df_merged.columns:
            df_merged['estado_beneficio'] = df_merged.apply(
                lambda row: f"{row['estado_beneficio']} {row[columnas_a_traer[1]]}".strip() 
                if pd.notna(row[columnas_a_traer[1]]) else row['estado_beneficio'], 
                axis=1
            )
        else:
            df_merged['estado_beneficio'] = df_merged[columnas_a_traer[1]].fillna('')
        
        if len(columnas_a_traer) > 2:
            df_merged['criterio_beneficio'] = df_merged.apply(
                lambda row: f"{row['criterio_beneficio']} {row[columnas_a_traer[2]]}".strip() 
                if pd.notna(row[columnas_a_traer[2]]) else row['criterio_beneficio'], 
                axis=1
            )
        else:
            df_merged['criterio_beneficio'] = df_merged[columnas_a_traer[2]].fillna('') if len(columnas_a_traer) > 2 else df_merged['criterio_beneficio']
    
    # Eliminar las columnas específicas del cruce, excepto las nuevas
    columns_to_drop = [columna_boleta] + columnas_a_traer[1:]
    df_merged = df_merged.drop(columns=columns_to_drop)
    
    return df_merged

# Realizar los cruces en cascada
cruce_id = 0
for nombre_hoja, columnas in columnas_por_periodo.items():
    if nombre_hoja in sheet_dic:
        piam_df = sheet_dic[nombre_hoja]
        columna_boleta = columnas[0]
        columnas_a_traer = [columna_boleta] + columnas[1:]
        
        cruce_id += 1  # Incrementar el ID del cruce
        
        # Realizar el cruce en cascada
        df_combined = realizar_cruce(df_combined, piam_df, columnas_a_traer, columna_boleta, cruce_id)
        print(f"Datos cruzados para {nombre_hoja} añadidos al DataFrame combinado.")
    else:
        print(f"DataFrame para la hoja {nombre_hoja} no encontrado o vacío.")

# Reemplazar None con una cadena vacía en las columnas 'estado_beneficio' y 'criterio_beneficio'
df_combined['estado_beneficio'].fillna('', inplace=True)
df_combined['criterio_beneficio'].fillna('', inplace=True)

# Eliminar la palabra "None" de las columnas 'estado_beneficio' y 'criterio_beneficio'
df_combined['estado_beneficio'] = df_combined['estado_beneficio'].str.replace('None', '').str.strip()
df_combined['criterio_beneficio'] = df_combined['criterio_beneficio'].str.replace('None', '').str.strip()

# Crear el archivo de salida
output_path = "/content/AUDITORIA_PAGOS_PIAM.xlsx"

with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    if not df_combined.empty:
        df_combined.to_excel(writer, sheet_name='Combina_Cruces', index=False)
        print(f"Datos combinados guardados en la hoja 'Combina_Cruces'.")
    else:
        print(f"No se encontraron datos para guardar en el archivo de salida.")

print(f"Archivo guardado en {output_path}")
