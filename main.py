import openpyxl as op

filename = 'Заказ клиента.xlsx'
subcategories_dict = {}

wb = op.load_workbook(filename, data_only=True)
sheet = wb.active
workbook = op.Workbook()
max_rows = sheet.max_row
sheet2 = workbook.active
row_ = 1
# list[]=[]
for i in range(13, max_rows + 1):
    sku = sheet.cell(row=i, column=6).value  # артикул товара
    name = sheet.cell(row=i, column=12).value  # наименование товара
    number = sheet.cell(row=i, column=38).value  # кол-во товара
    price = sheet.cell(row=i, column=48).value  # стоимость товара

    if not sku:
        continue

    print(sku, name, number, price)
    sheet2.cell(row=row_, column=1, value=sku)
    sheet2.cell(row=row_, column=2, value=name)
    sheet2.cell(row=row_, column=3, value=number)
    sheet2.cell(row=row_, column=4, value=price)
    row_ += 1

workbook.save('aldo_1c.xlsx')