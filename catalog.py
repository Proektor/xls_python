# import os
import csv
import xlwt
import xlrd
# import openpyxl
# import xlutils

with open('../src/catalog.csv') as src_cat:             # открываем исходный каталог
    rdr = csv.DictReader(src_cat)                       # объявляем построчный итератор - читальщик
    fieldName = str.split(rdr.fieldnames[0], ';')       # выделяем список полей заголовков

    # настраиваем стиль будущего шаблона каталога
# ------------------------------------------------------
    aligment = xlwt.Alignment()
    aligment.wrap = 1
    aligment.horz = xlwt.Alignment.HORZ_LEFT
    aligment.vert = xlwt.Alignment.VERT_CENTER

    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN

    font0 = xlwt.Font()
    font0.name = 'Arial cyr'
    font0.bold = True
    style0 = xlwt.XFStyle()
    style0.font = font0
# ------------------------------------------------------

    # создаем шаблон каталога с пустым листом
    emptyBook = xlwt.Workbook()
    emptySheet = emptyBook.add_sheet('catalog')

    # записываем строку заголовков
    i = 0
    for field in fieldName:
        emptySheet.write(0, i, field, style0)
        i += 1

    # сохраняем шаблон
    emptyBook.save('../out/template_catalog.xls')

    postName = []

    # обрабатываем исходный каталог построчно
    # for row in rdr:
    srcRow = next(rdr)
    data = str.split(srcRow[rdr.fieldnames[0]], ';')

    if data[-1] == '':
        namePost = 'Поставщик не указан'
    else:
        namePost = data[-1]

    if postName.__contains__(namePost):
        workB = xlwt.Workbook('../out/catalog_'+namePost)
    else:
        emptyBook.save('../out/catalog_'+namePost+'.xls')
        postName.append(namePost)

    workB = xlwt.Workbook('../out/catalog_'+namePost+'.xls')
    workSh = workB.set_active_sheet(0)

    # emptyRow
    # i = 0
    # for fieldData in data:
    #     workSh.put_cell(emptyRow,i,0,fieldData)
    #     i += 1









    print('1')
    # firstRow = str.split(, ';')

