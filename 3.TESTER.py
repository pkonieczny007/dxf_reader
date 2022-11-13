import os
import pandas as pd
import openpyxl
from shutil import copyfile


#tworzenie pliku porównania
print('tworzenie pliku porownania')
src = '2.PLIK_DO_SPRAWDZENIA.xlsx'
dst = '3.TESTER.xlsx'
copyfile(src, dst)


wb = openpyxl.load_workbook('3.TESTER.xlsx')
sheet = wb.active
sheet.title = 'porownanie'

#TWORZENIE DF WYKAZU 
df_wykaz = pd.read_excel('wykaz.xlsx')


#TWORZENIE LISTY Z WYKAZU
wykaz_lista = [
    [
        str(df_wykaz['NAZWA'][i]), 
        str(df_wykaz['Abmess_1'][i]), 
        str(df_wykaz['Abmes_2'][i]),
        str(df_wykaz['UWAGI'][i])
                                    ]
               
               for i in range(len(df_wykaz['Lp.']))]

#TWORZENIE LISTY Z DXF
df_dxf = pd.read_excel('2.PLIK_DO_SPRAWDZENIA.xlsx')
dxf_lista = [
    [
        str(df_dxf['ELEMENT_DXF'][i]), 
        str(df_dxf['X-DXF'][i]), 
        str(df_dxf['Y-DXF'][i]),
        str(df_dxf['UWAGI'][i])
                                    ]
               
               for i in range(len(df_dxf['lp']))]
#NAME_TEST - sprawdza nazwy z wykazu - czy jest na liscie dxf


#SORT_TEST - jeżeli nazwa zgadza się - sprawdzanie wymiarow - metoda sort


#TEST(+/-1)- jeżeli błąd w metodzie sort - sprawdzanie metodą (+/- 1)



#ZAPIS
print('zapis')
wb.save('3.TESTER.xlsx')
wb.close()
