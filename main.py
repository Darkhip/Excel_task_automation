# Automation of converting information from text file into Excel table
# Автоматизация переноса данных из текстового файла в таблицу Excel

from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

# Opening the text file for reading
# Открываем текстовый файл для чтения
text_file = open("textdata.txt")

records = []
text_file.seek(0)

for record in text_file.readlines():
    records.append(record.rstrip("/n").split(";"))

# Saving new Excel workbook with the path to location point
# Сохраняем новую книгу Excel
workbook = Workbook()
file_path = "exceldata.xlsx"
workbook.save(file_path)

# Renaming the default sheet to "Data"
# Переименовываем начальный лист в "Data"
sheet = workbook['Sheet']
sheet.title = 'Data'
sheet = workbook['Data']

for row in records:
    sheet.append(row)

# Creating the table in the sheet and defining a style for the table
# Создаем таблицу в листе Excel и определяем ее стиль
table = Table(displayName="Table", ref="A1:G11")
style = TableStyleInfo(name="TableStyleMedium11", showRowStripes=True, showColumnStripes=True)
table.tableStyleInfo = style
sheet.add_table(table)

workbook.save(file_path)

text_file.close()
workbook.close()
