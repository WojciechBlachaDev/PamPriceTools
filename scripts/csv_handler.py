import csv

def read_csv(file_path):
    try:
        with open(file_path, 'r', newline='', encoding='windows-1250') as file:
            reader = csv.reader(file, delimiter=';')
            headers = next(reader)
            data = [row for row in reader]
        print(f"Dane zostały pomyślnie wczytane z pliku {file_path}.")
        return headers, data
    except Exception as e:
        raise Exception(f"Wystąpił błąd podczas wczytywania pliku CSV: {str(e)}")

def save_csv(file_name, headers, data):
    try:
        with open(file_name, 'w', newline='', encoding='windows-1250') as file:
            writer = csv.writer(file, delimiter=';')
            writer.writerow(headers)
            for row in data:
                writer.writerow(row)
        print(f"Dane zostały pomyślnie zapisane do pliku {file_name}.")
    except Exception as e:
        raise Exception(f"Wystąpił błąd podczas zapisywania danych: {str(e)}")

def multiple_data_check(data, selected_column):
    try:
        id_count = {}
        for row in data:
            id_value = row[selected_column]
            if id_value in id_count:
                id_count[id_value] += 1
            else:
                id_count[id_value] = 1
        duplicates = {k: v for k, v in id_count.items() if v > 1}
        return duplicates
    except Exception as e:
        raise Exception(f"Wystąpił błąd podczas sprawdzania danych: {str(e)}")

def remove_entry(data, selected_column, value):
    updated_data = [row for row in data if row[selected_column] != value]
    if len(data) == len(updated_data):
        print(f"Nie znaleziono wpisu z wartością '{value}' w kolumnie {selected_column}.")
    else:
        print(f"Wpis z wartością '{value}' w kolumnie {selected_column} został usunięty.")
    return updated_data
