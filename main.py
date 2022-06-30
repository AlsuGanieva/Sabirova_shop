from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import (Alignment, Font, Border, Side)

book = load_workbook('sweets.xlsx')
sheet = book.active  # берем 1 лист активный
result_wb = Workbook()  # создаем рабочую книгу result_wb
result_ws = result_wb.active  # берем 1 лист
result_ws.title = 'Кондитерка'

# зададим размер и ориентацию страницы
result_ws.page_setup.orientation = 'portrait'
result_ws.page_setup.paperSize = result_ws.PAPERSIZE_A4

result_ws['A2'] = 'Наименование'
result_ws['B2'] = 'Кол-во'
result_ws['C2'] = 'Ед. изм'
result_ws['D2'] = 'Цена'
result_ws['E2'] = 'Сумма'

for row in range(2, sheet.max_row + 1):
    result_ws.append([sheet['D' + str(row)].value, sheet['E' + str(row)].value, sheet['F' + str(row)].value,
                      sheet['G' + str(row)].value, sheet['H' + str(row)].value])

for row in range(3, result_ws.max_row):
    result_ws[row][3].value = round(1.27 * result_ws[row][3].value)  # процент
    result_ws[row][3].font = Font(bold=True, size=14)  # жирный шрифт, размер шрифта 14

result_ws.merge_cells('A1:E1')
result_ws['A1'] = 'НАКЛАДНАЯ НА КОНДИТЕРКУ "НАДЕЖДА"'

# выравниваем текст по центру
result_ws.column_dimensions['A'].width = 50  # изменяю ширину 1 колонки

for row in range(3, result_ws.max_row):  # перенос текста, если не помещается
    if len(str(result_ws[row][0].value)) > 50:
        result_ws.row_dimensions[row].height = 40
        result_ws[row][0].alignment = Alignment(wrapText=True, vertical="center")

for cell in range(5):
    result_ws[2][cell].alignment = Alignment(horizontal="center", vertical="center")

result_ws['A1'].alignment = Alignment(horizontal="center", vertical="center")  # выравниваю название накл по центру
result_ws['A1'].font = Font(bold=True)  # жирный шрифт

for cell in range(1, 5):
    for row in range(3, result_ws.max_row):
        result_ws[row][cell].alignment = Alignment(horizontal="center", vertical="center")

thins = Side(border_style="thick", color="FF000000")  #жирный шрифт
double = Side(border_style="thin", color="FF000000")  #средний шрифт

for cell in range(5):  # делаю сетку
    for row in range(2, result_ws.max_row):
        result_ws[row][cell].border = Border(top=double, bottom=double, left=double, right=double)


result_wb.save("my_book.xlsx")
book.close()
