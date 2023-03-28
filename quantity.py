from openpyxl import load_workbook
from datetime import datetime


tride = load_workbook(filename='tride.xlsx')
ozon = load_workbook(filename='ozon.xlsx')


tride_sheets = tride.sheetnames
ozon_sheets = ozon.sheetnames

tride_sheet = tride[tride_sheets[0]]
ozon_sheet = ozon[ozon_sheets[1]]
count = 0



# приводим все занчения из столбца количество к 0 в остатках озона
def valueToZero():
    for i in range(1, ozon_sheet.max_row):
        if type(ozon_sheet.cell(row=i, column=5).value) is str:
            continue
        else:
            ozon_sheet.cell(row=i, column=5).value = 0


# соотносим артикли в прайсе поставщика с остатками озона
def articles():
    global count
    count = 0
    for i in range(1, ozon_sheet.max_row):
        x = typeObject(ozon_sheet.cell(row=i, column=3).value)

        for j in range(1, tride_sheet.max_row):
            y = typeObject(tride_sheet.cell(row=j, column=3).value)

            if y == x and ozon_sheet.cell(row=i, column=2).value is None:
                count += 1
                ozon_sheet.cell(row=i, column=2).value = tride_sheet.cell(row=j, column=1).value
                print(x, y)


 # проверка остатков
def quantity():
    global count

    # проверка на пустые строки
    def typeObject(word):
        if word is None:
            return word
        else:
            return word.replace(' ', '')
        
    
    for i in range(1, ozon_sheet.max_row):
        x = typeObject(ozon_sheet.cell(row=i, column=2).value)

        for j in range(1, tride_sheet.max_row):
            y = typeObject(tride_sheet.cell(row=j, column=1).value)


            if x == y:
                count += 1

                if tride_sheet.cell(row=j, column=4).value == 'Резерв':
                    ozon_sheet.cell(row=i, column=5).value = 0
                elif tride_sheet.cell(row=j, column=4).value == '2+':
                    ozon_sheet.cell(row=i, column=5).value = 3
                else:
                    if tride_sheet.cell(row=j, column=4).value == 'Склад':
                        continue
                    else:
                        ozon_sheet.cell(row=i, column=5).value = int(tride_sheet.cell(row=j, column=4).value)


quantity()

print(count, "совпадений")
ozon.save('ozon({}-{}).xlsx'.format(datetime.now().day, datetime.now().month))
