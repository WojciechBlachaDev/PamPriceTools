import time
import pandas as pd
import xlwings as xw
import openpyxl


def xlsx_read(path):
    try:
        data = pd.read_excel(path)
        return data, None
    except Exception as e:
        return None, e


def get_column_data(data, column):
    try:
        result = data.iloc[:, column - 1]
        return result, None
    except Exception as e:
        return None, e


def get_columns_count(data):
    try:
        return data.shape[1], None
    except Exception as e:
        return None, e


def get_rows_count(data):
    try:
        length = len(data)
        return length, None
    except Exception as e:
        return None, e


def get_row_data(data, row_number):
    try:
        row_data = data.iloc[row_number - 2]
        return row_data, None
    except Exception as e:
        return None, e


def open_workbook(file_path):
    try:
        app = xw.apps.active
        if not app:
            app = xw.App(visible=True, add_book=False)
        wb = app.books.open(file_path)
        return app, wb
    except Exception as e:
        print(e)
        return None, None


def fill_discount_table_2(data, pdf_data, wb, sheet_name):
    ws = wb.sheets[sheet_name]
    not_found = []
    for item in pdf_data:
        for i in range(len(data)):
            if str(data.iloc[i, 1]) == item[0]:
                cell = ws.cells(i + 2, 3)
                new_value = item[1].split('%')
                print(new_value)

                cell.value = float(new_value[0].replace(',', '.')) / 100
                break
            if '0' + str(data.iloc[i, 1]) == item[0]:
                cell = ws.cells(i + 2, 3)
                new_value = item[1].split('%')
                print(new_value)

                cell.value = float(new_value[0].replace(',', '.')) / 100
                break
            not_found.append(f'{str(data.iloc[i, 1])} != {item[0]}')
    for r in not_found:
        print(r)



def fill_discount_table(data, wb, sheet_name):
    try:
        ws = wb.sheets[sheet_name]
        counter = 1
        for value_to_find, value_to_fill in data:
            print(f'Rabat {value_to_find}: {value_to_fill}')
            cell = ws.api.UsedRange.Find(value_to_find)
            try:
                cell2 = ws.api.UsedRange.Find(float(value_to_find + '.0'))
            except Exception as e:
                print(e)
                pass
            if cell and cell.Column == 2:
                new_value = value_to_fill.split('%')
                adjacent_cell = ws.cells(cell.Row, cell.Column + 1)
                adjacent_cell.value = float(new_value[0].replace(',', '.')) / 100
                counter += 1
            if cell2 and cell2.Column == 2:
                new_value = value_to_fill.split('%')
                adjacent_cell = ws.cells(cell.Row, cell.Column + 1)
                adjacent_cell.value = float(new_value[0].replace(',', '.')) / 100
                counter += 1
        print(counter)
    except Exception as e:
        print(e)


def start_macro(wb, macro_tag):
    try:
        wb.macro(macro_tag).run()
    except Exception as e:
        print("Wystąpił błąd:", e)
