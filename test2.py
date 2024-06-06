from scripts import excel_handler
from scripts import pdf_tools
import time


pdf_path = 'C:/Users/MSI/Documents/PythonProjects/PPT/sample_files/cennik_siegenia.pdf'
excel_path = 'C:/Users/MSI/Documents/PythonProjects/PPT/sample_files/testowy_cennik.xlsm'
macro_name = 'Makro1'

pdf_data = pdf_tools.read_file(pdf_path)
pdf_tools.read_non_standard_prices(pdf_data, '2. Ceny specjalne', '3. Warunki dostawy')
