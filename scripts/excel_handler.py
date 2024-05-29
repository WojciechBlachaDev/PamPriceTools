import openpyxl

def read_excel_data(path, sheet_name, starting_row):
    try:
        workbook = openpyxl.load_workbook(path)
        sheet = workbook[sheet_name]
        data = []
        for row in sheet.iter_rows(values_only=True):
            data.append(row)
        headers = data[starting_row]
        data = data[starting_row+1:]  # Pomijamy pierwszy wiersz z nagłówkami
        print(f"Dane zostały pomyślnie wczytane z arkusza '{sheet_name}' w pliku {path}.")
        return headers, data
    except Exception as e:
        raise Exception(f"Wystąpił błąd podczas wczytywania danych z arkusza '{sheet_name}' w pliku Excel: {str(e)}")
