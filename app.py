import os
import ctypes
from scripts import csv_handler
from scripts import excel_handler
from scripts import json_handler
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog


class App:
    def __init__(self, main_root):
        self.start_button = None
        self.delete_option_check_box = None
        self.progress_bar_excel_value = None
        self.progress_bar_csv_value = None
        self.progress_bar_excel = None
        self.progress_bar_csv = None
        self.price_update_frame = None
        self.data_exchange_frame = None
        self.csv_data = None
        self.headers = None
        self.excel_data = None
        self.files_frame = None
        self.notebook = None
        self.excel_starting_row = tk.StringVar()
        self.csv_starting_row = tk.StringVar()
        self.root = None
        self.price_update_delete_option = tk.BooleanVar()
        self.settings_name = ''
        self.settings = json_handler.load_settings(os.path.join(os.getcwd(), 'exchange_settings.json'))
        if self.settings == {}:
            json_handler.save_settings(self.settings, os.path.join(os.getcwd(), 'exchange_settings.json'))
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
        self.current_settings = None
        self.set_main_window(main_root)
        self.set_notebook()
        self.set_file_page()
        self.set_data_exchange_page()
        self.csv_progress_counter = tk.StringVar()
        self.excel_progress_counter = tk.StringVar()
        self.set_price_update_page()

    def set_main_window(self, main_root):
        try:
            self.root = main_root
            self.root.title('Pam Price Tools')
            self.root.iconbitmap(os.path.join(os.getcwd(), 'app_icon.ico'))
            myappid = "BRK_Windows.PamPriceTools.PamPriceTools.version_0_0_1"
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
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
        settings_names = []
        for key, value in self.settings.items():
            settings_names.append(key)
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

        self.set_descriptions('Wprowadź wiersz startowy danych w plikach', 10)
        excel_starting_row = ttk.Entry(self.data_exchange_frame, textvariable=self.excel_starting_row)
        excel_starting_row.grid(row=11, column=2, padx=5, pady=5)
        csv_starting_row = ttk.Entry(self.data_exchange_frame, textvariable=self.csv_starting_row)
        csv_starting_row.grid(row=11, column=0, padx=5, pady=5)

        settings_label = ttk.Label(self.data_exchange_frame, text='Zapisane ustawienia')
        settings_label.grid(row=0, column=4, padx=10, pady=10)
        self.current_settings = ttk.Combobox(self.data_exchange_frame, values=settings_names)
        self.current_settings.grid(row=1, column=4, padx=10, pady=10)
        save_button = ttk.Button(self.data_exchange_frame, text='Zapisz', command=self.save_data_exchange_profile)
        save_button.grid(row=2, column=3, padx=5, pady=5)
        load_button = ttk.Button(self.data_exchange_frame, text='Ładuj', command=self.load_data_exchange_settings)
        load_button.grid(row=2, column=4, padx=5, pady=5)
        delete_button = ttk.Button(self.data_exchange_frame, text='Usuń', command=self.delete_data_exchange_settings)
        delete_button.grid(row=2, column=5, padx=5, pady=5)

    def set_price_update_page(self):
        self.csv_progress_counter.set('0')
        self.excel_progress_counter.set('0')
        self.price_update_frame = ttk.Frame(self.notebook)
        self.price_update_frame.grid(row=0, column=0, sticky='nsew')
        self.notebook.add(self.price_update_frame, text='Aktualizacja cen')
        self.progress_bar_csv = ttk.Progressbar(self.price_update_frame, orient='horizontal', length=300,
                                                mode="determinate")
        self.progress_bar_csv.grid(row=1, column=1, padx=5, pady=5)
        self.progress_bar_csv_value = ttk.Label(self.price_update_frame, textvariable=self.csv_progress_counter)
        self.progress_bar_csv_value.grid(row=1, column=2, padx=5, pady=5)
        self.delete_option_check_box = ttk.Checkbutton(self.price_update_frame,
                                                       variable=self.price_update_delete_option)
        self.delete_option_check_box.grid(row=3, column=1, padx=5, pady=5)
        delete_option_label = ttk.Label(self.price_update_frame, text='Usuń nieznalezione w cenniku wpisy')
        delete_option_label.grid(row=3, column=0, padx=5, pady=5)
        progress_label_csv = ttk.Label(self.price_update_frame, text='Postęp pliku CSV')
        progress_label_csv.grid(row=1, column=0, padx=5, pady=5)
        self.start_button = ttk.Button(self.price_update_frame,
                                       text='Aktualizuj bazę', command=self.update_prices)
        self.start_button.grid(row=4, column=1, padx=10, pady=10)

    def save_data_exchange_profile(self):
        new_settings = {
            'excel_discount_g': self.excel_discount_group_column.get(),
            'csv_discount_g': self.csv_discount_group_column.get(),
            'excel_discount_v': self.excel_discount_value_column.get(),
            'csv_discount_v': self.csv_discount_value_column.get(),
            'excel_base_price': self.excel_base_price_column.get(),
            'csv_base_price': self.csv_base_price_column.get(),
            'excel_cat_price': self.excel_catalogue_price_column.get(),
            'csv_cat_price': self.csv_catalogue_price_column.get(),
            'excel_search': self.excel_search_column.get(),
            'csv_search': self.csv_search_column.get(),
            'excel_start': self.excel_starting_row.get(),
            'csv_start': self.csv_starting_row.get()
        }
        self.settings[self.current_settings.get()] = new_settings
        json_handler.save_settings(self.settings, os.path.join(os.getcwd(), 'exchange_settings.json'))
        self.update_data_exchange_frame_settings()

    def load_data_exchange_settings(self):
        self.settings = json_handler.load_settings(os.path.join(os.getcwd(), 'exchange_settings.json'))
        for key, value in self.settings.items():
            if key == self.current_settings.get():
                self.excel_discount_group_column.set(self.settings[key]['excel_discount_g'])
                self.csv_discount_group_column.set(self.settings[key]['csv_discount_g'])
                self.excel_discount_value_column.set(self.settings[key]['excel_discount_v'])
                self.csv_discount_value_column.set(self.settings[key]['csv_discount_v'])
                self.excel_base_price_column.set(self.settings[key]['excel_base_price'])
                self.csv_base_price_column.set(self.settings[key]['csv_base_price'])
                self.excel_catalogue_price_column.set(self.settings[key]['excel_cat_price'])
                self.csv_catalogue_price_column.set(self.settings[key]['csv_cat_price'])
                self.excel_search_column.set(self.settings[key]['excel_search'])
                self.csv_search_column.set(self.settings[key]['csv_search'])
                self.excel_starting_row.set(self.settings[key]['excel_start'])
                self.csv_starting_row.set(self.settings[key]['csv_start'])

    def delete_data_exchange_settings(self):
        new_settings = {}
        for key, value in self.settings.items():
            if key != self.current_settings.get():
                new_settings[key] = value
        json_handler.save_settings(new_settings, os.path.join(os.getcwd(), 'exchange_settings.json'))
        self.settings = new_settings
        self.update_data_exchange_frame()

    def update_data_exchange_frame_settings(self):
        settings_names = []
        for key, value in self.settings.items():
            settings_names.append(key)
        selected_setting = self.current_settings.get()
        self.current_settings.destroy()
        self.current_settings = ttk.Combobox(self.data_exchange_frame, values=settings_names)
        self.current_settings.grid(row=1, column=4, padx=10, pady=10)
        self.current_settings.set(selected_setting)

    def update_data_exchange_frame(self):
        for widget in self.data_exchange_frame.winfo_children():
            widget.destroy()
        excel_columns = []
        csv_columns = []
        settings_names = []
        for key, value in self.settings.items():
            settings_names.append(key)
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

        self.set_descriptions('Wprowadź wiersz startowy danych w plikach', 10)
        excel_starting_row = ttk.Entry(self.data_exchange_frame, textvariable=self.excel_starting_row)
        excel_starting_row.grid(row=11, column=2, padx=5, pady=5)
        csv_starting_row = ttk.Entry(self.data_exchange_frame, textvariable=self.csv_starting_row)
        csv_starting_row.grid(row=11, column=0, padx=5, pady=5)

        settings_label = ttk.Label(self.data_exchange_frame, text='Zapisane ustawienia')
        settings_label.grid(row=0, column=4, padx=10, pady=10)
        self.current_settings = ttk.Combobox(self.data_exchange_frame, values=settings_names)
        self.current_settings.grid(row=1, column=4, padx=10, pady=10)
        save_button = ttk.Button(self.data_exchange_frame, text='Zapisz', command=self.save_data_exchange_profile)
        save_button.grid(row=2, column=3, padx=5, pady=5)
        load_button = ttk.Button(self.data_exchange_frame, text='Ładuj', command=self.load_data_exchange_settings)
        load_button.grid(row=2, column=4, padx=5, pady=5)
        delete_button = ttk.Button(self.data_exchange_frame, text='Usuń', command=self.delete_data_exchange_settings)
        delete_button.grid(row=2, column=5, padx=5, pady=5)

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
            self.csv_columns_count = len(self.csv_data[0])
            self.update_data_exchange_frame()

    def get_excel_path(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        self.root.config(cursor='wait')
        self.root.update()
        if path is not None:
            self.excel_raw_path.set(path)
            self.excel_data, _ = excel_handler.xlsx_read(self.excel_raw_path.get())
            self.excel_columns_count, _ = excel_handler.get_columns_count(self.excel_data)
            self.update_data_exchange_frame()
        self.root.config(cursor='')
        self.root.update()

    def update_prices(self):
        position_found = False
        positions_not_found = []
        self.start_button.config(state='disabled')
        search_values, _ = excel_handler.get_column_data(self.excel_data, int(self.excel_search_column.get()))
        for i in range(int(self.csv_starting_row.get()) - 1, len(self.csv_data)):
            self.csv_progress_counter.set(f"{i} / {len(self.csv_data)}")
            self.progress_bar_csv['value'] = (i / len(self.csv_data)) * 100
            self.price_update_frame.update()
            search = self.csv_data[i][int(self.csv_search_column.get()) - 1]
            for j in range(int(self.excel_starting_row.get()), len(self.excel_data)):
                if search == search_values[j]:
                    position_found = True
                    excel_row = excel_handler.get_row_data(self.excel_data, j + 2)
                    if excel_row is not None:
                        try:
                            csv_column = int(self.csv_discount_value_column.get()) - 1
                            excel_column = int(self.excel_discount_value_column.get()) - 1
                            if not pd.isna(excel_row[0].iloc[excel_column]):
                                self.csv_data[i][csv_column] = str(excel_row[0].iloc[excel_column]
                                                                   * 100).replace('.', ',')
                            else:
                                self.csv_data[i][csv_column] = '0.0'
                        except Exception as e:
                            print(f'Discount group: {e}')
                            pass
                        try:
                            csv_column = int(self.csv_discount_group_column.get()) - 1
                            excel_column = int(self.excel_discount_group_column.get()) - 1
                            if not pd.isna(excel_row[0].iloc[excel_column]):
                                self.csv_data[i][csv_column] = str(excel_row[0].iloc[excel_column])
                        except Exception as e:
                            print(f'Discount group: {e}')
                            pass
                        try:
                            csv_column = int(self.csv_base_price_column.get()) - 1
                            excel_column = int(self.excel_base_price_column.get()) - 1
                            if not pd.isna(excel_row[0].iloc[excel_column]):
                                self.csv_data[i][csv_column] = str(excel_row[0].iloc[excel_column]).replace('.', ',')
                            else:
                                self.csv_data[i][csv_column] = '0.0'
                        except Exception as e:
                            print(f'Discount group: {e}')
                            pass
                        try:
                            csv_column = int(self.csv_catalogue_price_column.get()) - 1
                            excel_column = int(self.excel_catalogue_price_column.get()) - 1
                            if not pd.isna(excel_row[0].iloc[excel_column]) and excel_row[0].iloc[excel_column] != 0.0:
                                self.csv_data[i][csv_column] = str(excel_row[0].iloc[excel_column]).replace('.', ',')
                            else:
                                self.csv_data[i][csv_column] = self.csv_data[i][int(self.csv_base_price_column.get()) - 1]
                        except Exception as e:
                            print(f'Discount group: {e}')
                            pass
                    break
            if not position_found:
                positions_not_found.append(search)
            if position_found:
                position_found = False
        if self.delete_option_check_box and len(positions_not_found) > 0:
            for i in range(len(positions_not_found)):
                updated_data, _ = csv_handler.remove_entry(self.csv_data, int(self.csv_search_column.get()) - 1,
                                                           positions_not_found[i])
                if updated_data is not None:
                    self.csv_data = updated_data
        new_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        print(new_path)
        if new_path is not None and new_path != '' and new_path != '.csv':
            csv_handler.save_csv(new_path, self.headers, self.csv_data)
        self.csv_progress_counter.set(f'Wielkość po operacji: {len(self.csv_data) + 1}')
        self.progress_bar_csv['value'] = 100
        self.start_button.config(state='enabled')
        self.price_update_frame.update()


if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()
