from scripts import excel_handler
from scripts import pdf_tools
import time


pdf_path = 'C:/Users/MSI/Documents/PythonProjects/PPT/sample_files/cennik_siegenia.pdf'
excel_path = 'C:/Users/MSI/Documents/PythonProjects/PPT/sample_files/testowy_cennik.xlsm'
macro_name = 'Makro1'

pdf_data = pdf_tools.read_file(pdf_path)
discount_list = pdf_tools.read_standard_discounts(pdf_data, '1. Warunki rabatowe na grupy produktowe', '2. Ceny specjalne')
print(len(discount_list))
for discount in discount_list:
    print(discount)
my_app, workbook = excel_handler.open_workbook(excel_path)
time.sleep(5)
readed_data, error = excel_handler.xlsx_read(excel_path)
excel_handler.fill_discount_table_2(readed_data, discount_list, workbook, 'Warunki Handlowe')