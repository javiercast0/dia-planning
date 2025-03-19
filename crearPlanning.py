import openpyxl
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import datetime

# Función para seleccionar un archivo
def seleccionar_archivo(titulo):
    root = Tk()
    root.withdraw()
    archivo = askopenfilename(title=titulo, filetypes=[("Excel files", "*.xlsm *.xlsx")])
    root.destroy()
    return archivo

# Cargar los archivos de origen de forma interactiva
source_file = seleccionar_archivo("Selecciona el archivo del día actual")
source_file2 = seleccionar_archivo("Selecciona el archivo del día anterior")
source_file3 = seleccionar_archivo("Selecciona el archivo de hace 2 días")
source_file4 = seleccionar_archivo("Selecciona el archivo de volumetría de Jaen")

# Cargamos las hojas de los archivos.
wb_source = openpyxl.load_workbook(source_file, data_only=True)
ws_source = wb_source['Resumen']

wb_source2 = openpyxl.load_workbook(source_file2, data_only=True)
ws_source2 = wb_source2['Resumen']  

wb_source3 = openpyxl.load_workbook(source_file3, data_only=True)
ws_source3 = wb_source3['Resumen']

wb_source4 = openpyxl.load_workbook(source_file4, data_only=True)
ws_source4 = wb_source4.active

# Crear un nuevo archivo Excel
new_wb = openpyxl.Workbook()
new_ws = new_wb.active

# Copiar los datos de la columna D7 hacia abajo de la volumetría del día actual.
# Basicamente conseguimos un listado de tiendas que se servirán al día siguiente.
row = 7
col = 4
while ws_source.cell(row=row, column=col).value is not None:
    new_ws.cell(row=row-6, column=1, value=ws_source.cell(row=row, column=col).value)
    row += 1

# Guardar el nuevo archivo Excel
fecha_actual = datetime.datetime.now().strftime("%d-%m").lstrip('0')
new_file = f'Volumetría {fecha_actual} (DRIVE).xlsx'
new_wb.save(new_file)

# Cargar el nuevo archivo Excel
wb_final = openpyxl.load_workbook(new_file)
ws_final = wb_final.active

# Iterar sobre todas las celdas de la columna A en el archivo final.
# Lo que hará será buscar el valor en la columna D del archivo de volumetrías y obtener el valor 13 celdas a la derecha (que coincide con el 11 del seco)
# Así con el area 91 (23 celdas a la derecha), 61 y 68 (21 y 22 celdas a la derecha, ya de la volumetría del dia anterior), 51 (19 valores a la derecha, ya de la volumetria
# de hace dos días) y con la Volumetría de Jaén.
for row in range(1, ws_final.max_row + 1):
    value_to_find = ws_final.cell(row=row, column=1).value
    if value_to_find is None:
        continue
    
    found_value = None
    for source_row in range(1, ws_source.max_row + 1):
        if ws_source.cell(row=source_row, column=4).value == value_to_find:
            found_value = ws_source.cell(row=source_row, column=17).value
            found_value2 = ws_source.cell(row=source_row, column=23).value
            if found_value is not None:
                ws_final.cell(row=row, column=2, value=found_value)
            if found_value2 is not None:
                ws_final.cell(row=row, column=3, value=found_value2)
         
    found_value = None
    for source2_row in range(1, ws_source2.max_row + 1):
        if ws_source2.cell(row=source2_row, column=4).value == value_to_find:
            found_value = ws_source2.cell(row=source2_row, column=21).value
            found_value2 = ws_source.cell(row=source2_row, column=22).value
            if found_value is not None:
                ws_final.cell(row=row, column=4, value=found_value)
            if found_value2 is not None:
                ws_final.cell(row=row, column=5, value=found_value2)
    
    found_value = None
    for source3_row in range(1, ws_source3.max_row + 1):
        if ws_source3.cell(row=source3_row, column=4).value == value_to_find:
            found_value = ws_source3.cell(row=source3_row, column=19).value
            if found_value is not None:
                ws_final.cell(row=row, column=6, value=found_value)
    
    found_value = None
    for source4_row in range(1, ws_source4.max_row + 1):
        if ws_source4.cell(row=source4_row, column=4).value == value_to_find:
            found_value = ws_source4.cell(row=source4_row, column=8).value
            if found_value is not None:
                ws_final.cell(row=row, column=7, value=found_value)
    
# Guardar los cambios en el archivo
wb_final.save(new_file)