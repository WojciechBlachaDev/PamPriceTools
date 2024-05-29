from scripts import csv_handler
from scripts import excel_handler
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog


class App:
    def __init__(self, main_root):
        self.data_exchange_frame = None
        self.csv_data = None
        self.headers = None
        self.excel_data = None
        self.files_frame = None
        self.notebook = None
        self.root = None
        self.csv_raw_path = tk.StringVar()
        self.excel_raw_path = tk.StringVar()
        self.excel_columns_count = 0
        self.csv_columns_count = 0
        self.excel_search_column = None
        self.csv_search_column = None
        self.excel_discount_value_column = None
        self.excel_discount_group_column = None
        self.excel_base_price_column = None
        self.excel_catalogue_price_column = None
        self.csv_discount_value_column = None
        self.csv_discount_group_column = None
        self.csv_base_price_column = None
        self.csv_catalogue_price_column = None
        self.set_main_window(main_root)
        self.set_notebook()
        self.set_file_page()
        self.set_data_exchange_page()

    def set_main_window(self, main_root):
        try:
            self.root = main_root
            self.root.title('Pam Price Tools')
        except Exception as e:
            raise Exception(f'Błąd tworzenia okna głównego aplikacji: {e}')

    def set_notebook(self):
        try:
            self.notebook = ttk.Notebook(self.root)
            self.notebook.grid(row=0, column=0, sticky='nsew')
        except Exception as e:
            raise Exception(f'Błąd tworzenia zeszytu dla okna głównego: {e}')

    def set_file_page(self):
        try:
            self.files_frame = ttk.Frame(self.notebook)
            self.files_frame.grid(row=0, column=0, sticky='nsew')
            self.notebook.add(self.files_frame, text='Pliki')
            csv_label = ttk.Label(self.files_frame, text='Wybierz plik CSV')
            csv_label.grid(row=0, column=0, sticky='w')
            csv_entry = ttk.Entry(self.files_frame, textvariable=self.csv_raw_path, width=100)
            csv_entry.grid(row=0, column=1, padx=5, pady=5)
            csv_button = ttk.Button(self.files_frame, text='Przeglądaj pliki', command=self.get_csv_path)
            csv_button.grid(row=0, column=2, padx=5, pady=5)
            excel_label = ttk.Label(self.files_frame, text='Wybierz plik Excel')
            excel_label.grid(row=1, column=0, sticky='w')
            excel_entry = ttk.Entry(self.files_frame, textvariable=self.excel_raw_path, width=100)
            excel_entry.grid(row=1, column=1, padx=5, pady=5)
            excel_button = ttk.Button(self.files_frame, text='Przeglądaj pliki', command=self.get_excel_path)
            excel_button.grid(row=1, column=2, padx=5, pady=5)
        except Exception as e:
            raise Exception(f'Błąd tworzenia strony obsługi plików: {e}')

    def set_data_exchange_page(self):
        excel_columns = []
        csv_columns = []
        for i in range(self.excel_columns_count):
            excel_columns.append(str(i + 1))
        for i in range(self.csv_columns_count):
            csv_columns.append(str(i + 1))
        self.data_exchange_frame = ttk.Frame(self.notebook)
        self.data_exchange_frame.grid(row=0, column=0, sticky='nsew')
        self.notebook.add(self.data_exchange_frame, text='Ustaw schemat wymiany danych')

        self.set_descriptions('Wybierz kolumnę do indeksowania wyszukiwania', 0)
        self.excel_search_column = ttk.Combobox(self.data_exchange_frame, values=excel_columns)
        self.excel_search_column.grid(row=1, column=2, padx=5, pady=5)
        self.csv_search_column = ttk.Combobox(self.data_exchange_frame, values=csv_columns)
        self.csv_search_column.grid(row=1, column=0, padx=5, pady=5)

        self.set_descriptions('Wybierz kolumnę z grupą rabatową', 2)
        self.excel_discount_group_column = ttk.Combobox(self.data_exchange_frame, values=excel_columns)
        self.excel_discount_group_column.grid(row=3, column=2, padx=5, pady=5)
        self.csv_discount_group_column = ttk.Combobox(self.data_exchange_frame, values=csv_columns)
        self.csv_discount_group_column.grid(row=3, column=0, padx=5, pady=5)

        self.set_descriptions('Wybierz kolumnę z wartością rabatu', 4)
        self.excel_discount_value_column = ttk.Combobox(self.data_exchange_frame, values=excel_columns)
        self.excel_discount_value_column.grid(row=5, column=2, padx=5, pady=5)
        self.csv_discount_value_column = ttk.Combobox(self.data_exchange_frame, values=csv_columns)
        self.csv_discount_value_column.grid(row=5, column=0, padx=5, pady=5)

        self.set_descriptions('Wybierz kolumnę z ceną bazową', 6)
        self.excel_base_price_column = ttk.Combobox(self.data_exchange_frame, values=excel_columns)
        self.excel_base_price_column.grid(row=7, column=2, padx=5, pady=5)
        self.csv_base_price_column = ttk.Combobox(self.data_exchange_frame, values=csv_columns)
        self.csv_base_price_column.grid(row=7, column=0, padx=5, pady=5)

        self.set_descriptions('Wybierz kolumnę z ceną katologową', 8)
        self.excel_catalogue_price_column = ttk.Combobox(self.data_exchange_frame, values=excel_columns)
        self.excel_catalogue_price_column.grid(row=9, column=2, padx=5, pady=5)
        self.csv_catalogue_price_column = ttk.Combobox(self.data_exchange_frame, values=csv_columns)
        self.csv_catalogue_price_column.grid(row=9, column=0, padx=5, pady=5)



    def update_data_exchange_frame(self):
        for widget in self.data_exchange_frame.winfo_children():
            widget.destroy()
        excel_columns = []
        csv_columns = []
        for i in range(self.excel_columns_count):
            excel_columns.append(str(i + 1))
        for i in range(self.csv_columns_count):
            csv_columns.append(str(i + 1))
        self.set_descriptions('Wybierz kolumnę do indeksowania wyszukiwania', 0)
        self.excel_search_column = ttk.Combobox(self.data_exchange_frame, values=excel_columns)
        self.excel_search_column.grid(row=1, column=2, padx=5, pady=5)
        self.csv_search_column = ttk.Combobox(self.data_exchange_frame, values=csv_columns)
        self.csv_search_column.grid(row=1, column=0, padx=5, pady=5)

        self.set_descriptions('Wybierz kolumnę z grupą rabatową', 2)
        self.excel_discount_group_column = ttk.Combobox(self.data_exchange_frame, values=excel_columns)
        self.excel_discount_group_column.grid(row=3, column=2, padx=5, pady=5)
        self.csv_discount_group_column = ttk.Combobox(self.data_exchange_frame, values=csv_columns)
        self.csv_discount_group_column.grid(row=3, column=0, padx=5, pady=5)

        self.set_descriptions('Wybierz kolumnę z wartością rabatu', 4)
        self.excel_discount_value_column = ttk.Combobox(self.data_exchange_frame, values=excel_columns)
        self.excel_discount_value_column.grid(row=5, column=2, padx=5, pady=5)
        self.csv_discount_value_column = ttk.Combobox(self.data_exchange_frame, values=csv_columns)
        self.csv_discount_value_column.grid(row=5, column=0, padx=5, pady=5)

        self.set_descriptions('Wybierz kolumnę z ceną bazową', 6)
        self.excel_base_price_column = ttk.Combobox(self.data_exchange_frame, values=excel_columns)
        self.excel_base_price_column.grid(row=7, column=2, padx=5, pady=5)
        self.csv_base_price_column = ttk.Combobox(self.data_exchange_frame, values=csv_columns)
        self.csv_base_price_column.grid(row=7, column=0, padx=5, pady=5)

        self.set_descriptions('Wybierz kolumnę z ceną katologową', 8)
        self.excel_catalogue_price_column = ttk.Combobox(self.data_exchange_frame, values=excel_columns)
        self.excel_catalogue_price_column.grid(row=9, column=2, padx=5, pady=5)
        self.csv_catalogue_price_column = ttk.Combobox(self.data_exchange_frame, values=csv_columns)
        self.csv_catalogue_price_column.grid(row=9, column=0, padx=5, pady=5)

    def set_descriptions(self, text, row):
        label_search_columns = ttk.Label(self.data_exchange_frame, text=text)
        label_search_columns.grid(row=row, column=1, padx=5, pady=5)
        excel_desc = ttk.Label(self.data_exchange_frame, text='Kolumna CSV')
        excel_desc.grid(row=row, column=0, padx=5, pady=5)
        csv_desc = ttk.Label(self.data_exchange_frame, text='Kolumna Excel')
        csv_desc.grid(row=row, column=2, padx=5, pady=5)

    def get_csv_path(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if path is not None:
            self.csv_raw_path.set(path)
            _, self.headers, self.csv_data = csv_handler.read_csv(self.csv_raw_path.get())
            self.csv_columns_count = len(self.headers)
            self.update_data_exchange_frame()

    def get_excel_path(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path is not None:
            self.excel_raw_path.set(path)
            self.excel_data, _ = excel_handler.xlsx_read(self.excel_raw_path.get())
            self.excel_columns_count, _ = excel_handler.get_columns_count(self.excel_data)
            self.update_data_exchange_frame()

if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()
