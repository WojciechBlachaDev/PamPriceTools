import csv
import os


def read_csv(path):
    try:
        with open(path, 'r', newline='', encoding='windows-1250') as csv_file:
            reader = csv.reader(csv_file, delimiter=';')
            headers = next(reader)
            data = [row for row in reader]
        return None, headers, data
    except Exception as e:
        return e, None, None


def save_csv(path, headers, data):
    try:
        with open(path, 'w', newline='', encoding='windows-1250') as csv_file:
            writer = csv.writer(csv_file, delimiter=';')
            writer.writerow(headers)
            for row in data:
                writer.writerow(row)
        return True, None
    except Exception as e:
        return False, e


def multiple_data_check(data, column):
    try:
        id_count = {}
        for row in data:
            id_value = row[column]
            if id_value in id_count:
                id_count[id_value] += 1
            else:
                id_count[id_value] = 1
        duplicates = {k: v for k, v in id_count.items() if v > 1}
        return duplicates, None
    except Exception as e:
        return None, e


def remove_entry(data, column, value):
    try:
        updated_data = [row for row in data if row[column] != value]
        if len(data) == len(updated_data):
            return None, None
        else:
            return updated_data, None
    except Exception as e:
        return None, e


def update_entry(data, csv_index, csv_column, excel_value):
    row = data[csv_index - 2]
    print(row)
    row[csv_column - 1] = str(excel_value)
    print(row)
