import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime

# Funkcje obliczające wymiary plików DXF
def sze(filename):
    with open(filename, 'r', encoding='utf-8', errors='replace') as file_obj:
        file = file_obj.read()
        lista = file.split()
    min_index = lista.index('$EXTMIN')
    max_index = lista.index('$EXTMAX')
    minX = float(lista[min_index + 2])
    maxX = float(lista[max_index + 2])
    return int(maxX) - int(minX)

def wys(filename):
    with open(filename, 'r', encoding='utf-8', errors='replace') as file_obj:
        file = file_obj.read()
        lista = file.split()
    min_index = lista.index('$EXTMIN')
    max_index = lista.index('$EXTMAX')
    minY = float(lista[min_index + 4])
    maxY = float(lista[max_index + 4])
    return int(maxY) - int(minY)

# Przygotowanie DataFrame
excel_file = 'WYKAZ_DXF-FAST.xlsx'
try:
    df = pd.read_excel(excel_file)
except FileNotFoundError:
    df = pd.DataFrame(columns=['lp', 'ELEMENT_DXF', 'X-DXF', 'Y-DXF', 'UWAGI', 'DATA_UTWORZENIA', 'ZMIANY'])

# Skanowanie folderu i aktualizacja DataFrame
for filename in os.listdir():
    if filename.endswith('.dxf'):
        file_path = os.path.join(os.getcwd(), filename)
        creation_time = os.path.getctime(file_path)
        creation_date = datetime.datetime.fromtimestamp(creation_time).strftime('%Y-%m-%d %H:%M:%S')
        szerokosc = sze(filename)
        wysokosc = wys(filename)

        # Sprawdzenie, czy plik jest już w DataFrame
        if filename[:-4] in df['ELEMENT_DXF'].values:
            # Aktualizacja istniejącego wpisu
            row_index = df.index[df['ELEMENT_DXF'] == filename[:-4]].tolist()[0]
            if df.at[row_index, 'X-DXF'] != szerokosc or df.at[row_index, 'Y-DXF'] != wysokosc:
                df.at[row_index, 'ZMIANY'] = 'zmiana'
            df.at[row_index, 'X-DXF'] = szerokosc
            df.at[row_index, 'Y-DXF'] = wysokosc
            df.at[row_index, 'DATA_UTWORZENIA'] = creation_date
        else:
            # Dodanie nowego wpisu
            df = df.append({'lp': len(df) + 1, 'ELEMENT_DXF': filename[:-4], 'X-DXF': szerokosc, 'Y-DXF': wysokosc, 'UWAGI': '', 'DATA_UTWORZENIA': creation_date, 'ZMIANY': ''}, ignore_index=True)

# Sortowanie DataFrame według daty utworzenia
df.sort_values(by='DATA_UTWORZENIA', inplace=True)

# Zapis do Excel z użyciem openpyxl dla auto dopasowania szerokości kolumn
wb = Workbook()
ws = wb.active

for r in dataframe_to_rows(df, index=False, header=True):
    ws.append(r)

for col in ws.columns:
    max_length = max(len(str(cell.value)) for cell in col)
    ws.column_dimensions[col[0].column_letter].width = max_length + 2

wb.save(excel_file)
