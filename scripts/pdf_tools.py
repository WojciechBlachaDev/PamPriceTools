import pdfplumber


def read_file(path):
    try:
        with pdfplumber.open(path) as pdf:
            data = ''
            for page in pdf.pages:
                data += page.extract_text() + '\n'
            return data, None
    except Exception as e:
        print(e)
        return None, e


def read_standard_discounts(data, start_header, end_header):
    standard_discounts = []
    try:
        start_index = data.find(start_header)
        end_index = data.find(end_header)
        if start_index == -1 or end_index == -1:
            return None, None
        standard_discounts_page = data[start_index:end_index]
        lines = standard_discounts_page.split('\n')
        for line in lines:
            if '%' in line:
                parts = line.split()
                group = parts[0]
                discount = parts[-2] + parts[-1]
                standard_discounts.append((group, discount))
        return standard_discounts, None
    except Exception as e:
        return None, e


def read_non_standard_prices(data, start_header, end_header):
    try:
        start_index = data.find(start_header)
        end_index = data.find(end_header)
        if start_index == -1 or end_index == -1:
            raise Exception('xD')
        standard_discounts_page = data[start_index:end_index]
        lines = standard_discounts_page.split('\n')
        index_list = []
        for line in lines:
            elements = line.split(' ')  # Podzielenie linii na poszczególne elementy
            index_filter = [
                'Materiał',
                '',
                'Ilość',
                'Potwierdzenie',
                'Siedziba:',
                'Telefon',
                'PORTAL'
            ]
            if len(elements) > 5 and len(elements[0]) > 4 and elements[0] not in index_filter:
                price_data = elements[-5:]
                price_converted = price_data[0].replace('.', '')
                price = float(price_converted.replace(',', '.')) / float(price_data[3])
                index_list.append([elements[0], price])
        for index in index_list:
            index[1] = str(index[1]).replace('.', ',')
        return index_list, None
    except Exception as e:
        return None, e
