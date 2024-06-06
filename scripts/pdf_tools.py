import os
import pdfplumber


def read_file(path):
    try:
        with pdfplumber.open(path) as pdf:
            data = ''
            for page in pdf.pages:
                data += page.extract_text() + '\n'
            return data
    except Exception as e:
        print(e)
        return None


def read_standard_discounts(data, start_header, end_header):
    standard_discounts = []
    try:
        start_index = data.find(start_header)
        end_index = data.find(end_header)
        if start_index == -1 or end_index == -1:
            return None
        standard_discounts_page = data[start_index:end_index]
        lines = standard_discounts_page.split('\n')
        for line in lines:
            if '%' in line:
                parts = line.split()
                group = parts[0]
                discount = parts[-2] + parts[-1]
                standard_discounts.append((group, discount))
        return standard_discounts
    except Exception as e:
        return None