import os
import xlsxwriter
import pandas as pd


pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

sheet_dic = {}

# Verifica la existencia del archivo en la ruta especifica
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


# Obtiene los nombres de las hojas de trabajo

def load_excel_sheets(file_path):
  xls = pd.ExcelFile(file_path, engine='openpyxl')
  sheet_names = xls.sheet_names
  print("Nombres de las hojas:", sheet_names)

  for sheet_name in sheet_names:
    sheet_dic[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
  return sheet_dic, sheet_names

sheet_dic, sheet_names = load_excel_sheets(file_path)

for sheet_name in sheet_names:
    globals()[sheet_name] = sheet_dic[sheet_name]

if 'CONSULTACARO' in globals() and 'PIAM2021_1' in globals():
  print("Las hojas 'CONSULTACARO' y 'PIAM2021_1' existen en el diccionario.")
  consultaCaro = globals()['CONSULTACARO']
  piam2021_1 = globals()['PIAM2021_1']
  consultaCaro['Documento'] = consultaCaro['Documento'].astype(str)
  piam2021_1['BOLETA'] = piam2021_1['BOLETA'].astype(str)
  consultaCaroCruzado = pd.merge(consultaCaro, piam2021_1[['BOLETA','RECURSOS APLICADOS']], left_on='Documento', right_on='BOLETA', how='left')
if 'CONSULTACARO' in globals() and 'PIAM2021_2' in globals():
  print("Las hojas 'CONSULTACARO' y 'PIAM2021_2' existen en el diccionario.")
  consultaCaro = globals()['CONSULTACARO']
  piam2021_2 = globals()['PIAM2021_2']
  consultaCaro['Documento'] = consultaCaro['Documento'].astype(str)
  piam2021_2['BOLETA'] = piam2021_2['BOLETA'].astype(str)
  consultaCaroCruzado = pd.merge(consultaCaro, piam2021_2[['BOLETA','ESTADO F','ESTADO']], left_on='Documento', right_on='BOLETA', how='left')
else:
    print("Uno o ambos DataFrames no se han cargado correctamente.")


output_path = "/content/AUDITORIA_PAGOS_PIAM.xlsx"
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:

  consultaCaroCruzado.to_excel(writer, sheet_name='PIAM20221_SQCaro', index=False)

print(f"Archivo guardado en {output_path}")
