import os.path
import pandas as pd


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
