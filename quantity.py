from openpyxl import load_workbook
from datetime import date
from send_email import send_email


tride = load_workbook(filename='tride.xlsx')
ozon = load_workbook(filename='ozon.xlsx')


tride_sheets = tride.sheetnames
ozon_sheets = ozon.sheetnames

tride_sheet = tride[tride_sheets[0]]
ozon_sheet = ozon[ozon_sheets[1]]

    # проверка на пустые строки
def type_object(word):
    if word is None:
        return word
    else:
        return word.replace(' ', '')

# приводим все занчения из столбца количество к 0 в остатках озона
def valueToZero():
    for i in range(1, ozon_sheet.max_row):
        if type(ozon_sheet.cell(row=i, column=5).value) is str:
            continue
        else:
            ozon_sheet.cell(row=i, column=5).value = 0


# соотносим артикли в прайсе поставщика с остатками озона
def articles():
    count = 0
    for i in range(1, ozon_sheet.max_row):
        x = type_object(ozon_sheet.cell(row=i, column=3).value)

        for j in range(1, tride_sheet.max_row):
            y = type_object(tride_sheet.cell(row=j, column=3).value)

            if y == x and ozon_sheet.cell(row=i, column=2).value is None:
                count += 1
                ozon_sheet.cell(row=i, column=2).value = tride_sheet.cell(row=j, column=1).value
                print(x, y)
    print(f'Добавлено {count} новых артикулов')


 # проверка остатков
def quantity():
    count = 0
    
    for i in range(1, ozon_sheet.max_row):
        x = type_object(ozon_sheet.cell(row=i, column=2).value)

        for j in range(1, tride_sheet.max_row):
            y = type_object(tride_sheet.cell(row=j, column=1).value)


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

    ozon_sheet.delete_cols(2)
    print(f'Обновлено {count} товаров!')
    ozon.save(f'ozon_{date.today().strftime("%d-%b-%Y")}.xlsx')
