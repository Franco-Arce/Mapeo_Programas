import pandas as pd
import openpyxl

# Ruta del archivo original
archivo_origen = "ADMITIDOS VIRTUAL 2026-1  PREGRADO CIERRE.xlsx"
archivo_destino = "Admitidos_Consolidado.xlsx"

# Cargar el archivo Excel
wb = openpyxl.load_workbook(archivo_origen)
sheet_names = wb.sheetnames

print(f"Total de hojas encontradas: {len(sheet_names)}")
print("Hojas en el archivo:")
for i, name in enumerate(sheet_names):
    print(f"  {i}: {name}")

# Encontrar el índice de la hoja "ADMON"
try:
    indice_admon = sheet_names.index("ADMON")
    print(f"\nLa hoja 'ADMON' está en la posición {indice_admon}")
except ValueError:
    # Si no encuentra exactamente "ADMON", buscar hojas que contengan "ADMON"
    indice_admon = None
    for i, name in enumerate(sheet_names):
        if "ADMON" in name.upper():
            indice_admon = i
            print(f"\nSe encontró una hoja similar: '{name}' en la posición {i}")
            break
    
    if indice_admon is None:
        print("\nNo se encontró ninguna hoja con 'ADMON' en el nombre.")
        print("Por favor, verifica el nombre exacto de la hoja.")
        exit(1)

# Hojas a consolidar (desde ADMON en adelante)
hojas_a_consolidar = sheet_names[indice_admon:]
print(f"\nHojas a consolidar ({len(hojas_a_consolidar)}):")
for name in hojas_a_consolidar:
    print(f"  - {name}")

# Lista para almacenar los DataFrames
dfs = []

# Leer cada hoja y agregarla a la lista
for hoja in hojas_a_consolidar:
    print(f"\nProcesando hoja: {hoja}")
    df = pd.read_excel(archivo_origen, sheet_name=hoja)
    print(f"  Filas: {len(df)}, Columnas: {len(df.columns)}")
    
    # Agregar una columna para identificar de qué hoja proviene cada fila
    df['Hoja_Origen'] = hoja
    
    dfs.append(df)

# Concatenar todos los DataFrames
print("\nConsolidando todas las hojas...")
df_consolidado = pd.concat(dfs, ignore_index=True)

print(f"\nDataFrame consolidado:")
print(f"  Total de filas: {len(df_consolidado)}")
print(f"  Total de columnas: {len(df_consolidado.columns)}")

# Guardar en un nuevo archivo Excel
print(f"\nGuardando en '{archivo_destino}'...")
df_consolidado.to_excel(archivo_destino, index=False, sheet_name="Consolidado")

print(f"\nProceso completado exitosamente!")
print(f"Archivo guardado: {archivo_destino}")
