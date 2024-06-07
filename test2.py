from scripts import pdf_tools

pdf_path = 'C:/Users/MSI/Documents/PythonProjects/PPT/sample_files/cennik_siegenia.pdf'
excel_path = 'C:/Users/MSI/Documents/PythonProjects/PPT/sample_files/testowy_cennik.xlsm'
macro_name = 'Makro1'
pdf_data = pdf_tools.read_file(pdf_path)
special_prices_data = pdf_tools.read_non_standard_prices(pdf_data[0], '2. Ceny specjalne',
                                                         '3. Warunki dostawy')
for price in special_prices_data[0]:
    print(price)
