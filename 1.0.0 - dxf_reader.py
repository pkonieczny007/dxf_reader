import os 
import openpyxl

#tworzenie pliku excel
wb = openpyxl.Workbook()
wb.save('1.WYKAZ_DXF.xlsx')
wb.close()

#funkcje 
def sze(filename):
    with open(filename) as file_obj:
        file=file_obj.read()
        lista = file.split()

    min_index= lista.index('$EXTMIN')
    max_index=lista.index('$EXTMAX')

    minx_index=min_index + 2
    maxx_index=max_index +2

    minX=float(lista[minx_index])
    maxX=float(lista[maxx_index])

    miny_index=min_index + 4
    maxy_index=max_index +4          

    minY=float(lista[miny_index])
    maxY=float(lista[maxy_index])

    szerokosc=int(maxX)-int(minX)

    return szerokosc




def wys(filename):
    with open(filename) as file_obj:
        file=file_obj.read()
        lista = file.split()

    min_index= lista.index('$EXTMIN')
    max_index=lista.index('$EXTMAX')

    minx_index=min_index + 2
    maxx_index=max_index +2

    minX=float(lista[minx_index])
    maxX=float(lista[maxx_index])

    miny_index=min_index + 4
    maxy_index=max_index +4          

    minY=float(lista[miny_index])
    maxY=float(lista[maxy_index])

    wysokosc=int(maxY)-int(minY)

    return wysokosc



#aplikacja tworzenia wykazu dxf


n=0


wb = openpyxl.load_workbook('1.WYKAZ_DXF.xlsx')
sheet = wb.active
sheet.title = 'wykaz_dxf'

#nagłówki tabeli
sheet['A1']='lp'
sheet['B1']='ELEMENT_DXF'
sheet['C1']='X-DXF'
sheet['D1']='Y-DXF'
sheet['e1']='UWAGI'

for filename in os.listdir():
    if filename.endswith('.dxf'):
        try:
            n+=1
            wys1=wys(filename)
            szer1=sze(filename)
            print(n, filename, wys1, szer1)

            a="A"+str(n+1)
            b="B"+str(n+1)
            c="C"+str(n+1)
            d="D"+str(n+1)
            
            sheet[a]=str(n)
            sheet[b]=str(filename[:-4])
            sheet[c]=str(wys1)
            sheet[d]=str(szer1)
        except:            
            a="A"+str(n+1)
            b="B"+str(n+1)
            c="C"+str(n+1)
            d="D"+str(n+1)
            e="E"+str(n+1)
            sheet[a]=str(n)
            sheet[b]=str(filename[:-4])
            sheet[c]=str(wys1)
            sheet[d]=str(szer1)
            sheet[e]='blad'

        wb.save('1.WYKAZ_DXF.xlsx')



wb.save('1.WYKAZ_DXF.xlsx')
wb.close()
