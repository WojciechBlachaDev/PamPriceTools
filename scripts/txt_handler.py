def save_txt(path, data):
    with open(path, 'w', encoding='utf-8') as file:
        for item in data:
            file.write(str(item) + ',' + '\n')