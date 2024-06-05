import ctypes
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import pandas as pd
from scripts import csv_handler
from scripts import excel_handler
from scripts import json_handler
from scripts import txt_handler


class App:
    def __init__(self, main_root):
        self.start_button_verify = None
        self.save_not_found_check_box = None
        self.dialog = None
        self.save_button = None
        self.empty_start_button = None
        self.multiple_input_listbox = None
        self.csv_descriptions = []
        self.multiple_inputs_selected = None
        self.delete_bool = False
        self.skip_all = False
        self.verify_start_button = None
        self.verify_start = None
        self.verify_progressbar = None
        self.verify_frame = None
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
        self.csv_data_verify = None
        self.headers_verify = None
        self.excel_data = None
        self.files_frame = None
        self.notebook = None
        self.excel_starting_row = tk.StringVar()
        self.csv_starting_row = tk.StringVar()
        self.save_not_found_index = tk.BooleanVar()
        self.root = None
        self.price_update_delete_option = tk.BooleanVar()
        self.settings_name = ''
        self.settings = json_handler.load_settings(os.path.join(os.getcwd(), 'exchange_settings.json'))
        if self.settings == {}:
            json_handler.save_settings(self.settings, os.path.join(os.getcwd(), 'exchange_settings.json'))
        self.csv_raw_path = tk.StringVar()
        self.excel_raw_path = tk.StringVar()
        self.csv_verify_path = tk.StringVar()
        self.excel_columns_count = 0
        self.csv_columns_count = 0
        self.csv_columns_count_verify = 0
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
        self.set_verify_page()
        self.items_to_delete = []

    def set_main_window(self, main_root):
        try:
            self.root = main_root
            self.root.title('Pam Price Tools')
            self.root.iconbitmap(os.path.join(os.getcwd(), 'app_icon.ico'))
            myappid = "BRK_Windows.PamPriceTools.PamPriceTools.version_0_0_1"
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception as e:
            messagebox.showerror('Pam Price Tools - ERROR:', f'Błąd podczas tworzenia okna głównego aplikacji: {e}')

    def set_notebook(self):
        try:
            self.notebook = ttk.Notebook(self.root)
            self.notebook.grid(row=0, column=0, sticky='nsew')
        except Exception as e:
            messagebox.showerror('Pam Price Tools - ERROR:', f'Błąd podczas tworzenia zakładek okna aplikacji: {e}')

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
            messagebox.showerror('Pam Price Tools - ERROR:', f'Błąd podczas tworzenia strony obsługi plików: {e}')

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
        self.notebook.add(self.data_exchange_frame, text='Ustaw schemat wymiany danych bazy okuć')

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

        self.set_descriptions_rows('Wprowadź wiersz startowy danych w plikach', 10)
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

    def set_verify_page(self):
        self.verify_frame = ttk.Frame(self.notebook)
        self.verify_frame.grid(row=0, column=0, sticky='nsew')
        self.notebook.add(self.verify_frame, text='Weryfikuj plik CSV')

        csv_label = ttk.Label(self.verify_frame, text='Wybierz plik CSV')
        csv_label.grid(row=0, column=0, sticky='w')
        csv_entry = ttk.Entry(self.verify_frame, textvariable=self.csv_verify_path, width=100)
        csv_entry.grid(row=0, column=1, padx=5, pady=5)
        csv_button = ttk.Button(self.verify_frame, text='Przeglądaj pliki', command=self.get_csv_verify_path)
        csv_button.grid(row=0, column=2, padx=5, pady=5)

        self.verify_start_button = ttk.Button(self.verify_frame, text='Szukaj zwielokrotnień',
                                              command=self.csv_data_verify_method)
        self.verify_start_button.grid(row=3, column=1, padx=50, pady=5, sticky='w')
        self.verify_start_button.config(state='disabled')

        self.empty_start_button = ttk.Button(self.verify_frame, text='Szukaj pustych wpisów',
                                             command=self.csv_data_empty_verify_method)
        self.empty_start_button.grid(row=3, column=1, padx=50, pady=5, sticky='e')
        self.empty_start_button.config(state='disabled')

        multiple_inputs_label = ttk.Label(self.verify_frame, text='Wybierz kolumnę dokumentu')
        multiple_inputs_label.grid(row=1, column=1, padx=5, pady=5)
        self.multiple_input_listbox = tk.Listbox(self.verify_frame, selectmode='single')
        self.multiple_input_listbox.grid(row=2, column=1)

        self.save_button = ttk.Button(self.verify_frame, text='Zapisz dane', command=self.save_verified_file)
        self.save_button.grid(row=4, column=1, padx=5, pady=5)
        self.save_button.config(state='disabled')

    def set_price_update_page(self):
        self.csv_progress_counter.set('0')
        self.excel_progress_counter.set('0')
        self.price_update_frame = ttk.Frame(self.notebook)
        self.price_update_frame.grid(row=0, column=0, sticky='nsew')
        self.notebook.add(self.price_update_frame, text='Aktualizacja cen okuć')
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
        self.save_not_found_check_box = ttk.Checkbutton(self.price_update_frame,
                                                        variable=self.save_not_found_index)
        self.save_not_found_check_box.grid(row=4, column=1, padx=5, pady=5)
        not_found_label = ttk.Label(self.price_update_frame, text='Zapisz nieznalezione indexy')
        not_found_label.grid(row=4, column=0, padx=5, pady=5)
        progress_label_csv = ttk.Label(self.price_update_frame, text='Postęp pliku CSV')
        progress_label_csv.grid(row=1, column=0, padx=5, pady=5)
        self.start_button = ttk.Button(self.price_update_frame,
                                       text='Aktualizuj bazę', command=self.update_prices)
        self.start_button.grid(row=5, column=1, padx=10, pady=10)
        self.start_button_verify = ttk.Button(self.price_update_frame,
                                              text='Porównaj ceny z plików', command=self.verify_prices)
        self.start_button_verify.grid(row=6, column=1, padx=10, pady=10)

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
        result = json_handler.save_settings(self.settings, os.path.join(os.getcwd(), 'exchange_settings.json'))
        if not result[0]:
            messagebox.showerror('Pam Price Tools - BŁĄD:', f'Wykryto bład podczas zapisu pliku ustawień: {result[1]}')
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
        result = json_handler.save_settings(new_settings, os.path.join(os.getcwd(), 'exchange_settings.json'))
        if not result[0]:
            messagebox.showerror('Pam Price Tools - BŁĄD:', f'Wykryto błąd podczas usuwania ustawień {result[1]}')
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

        self.set_descriptions_rows('Wprowadź wiersz startowy danych w plikach', 10)
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

    def set_descriptions_rows(self, text, row):
        label_search_columns = ttk.Label(self.data_exchange_frame, text=text)
        label_search_columns.grid(row=row, column=1, padx=5, pady=5)
        excel_desc = ttk.Label(self.data_exchange_frame, text='Wiersz CSV')
        excel_desc.grid(row=row, column=0, padx=5, pady=5)
        csv_desc = ttk.Label(self.data_exchange_frame, text='Wiersz Excel')
        csv_desc.grid(row=row, column=2, padx=5, pady=5)

    def get_csv_path(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if path is not None and path != '' and path != '.csv':
            self.csv_raw_path.set(path)
            error, self.headers, self.csv_data = csv_handler.read_csv(self.csv_raw_path.get())
            if error is not None:
                messagebox.showerror('Pam Price Tools - BŁĄD:', f'Wystapił bład odczytu danych z pliku CSV - {error}')
            self.csv_columns_count = len(self.csv_data[0])

    def get_csv_verify_path(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if path is not None and path != '' and path != '.csv':
            self.csv_verify_path.set(path)
            error, self.headers_verify, self.csv_data_verify = csv_handler.read_csv(self.csv_verify_path.get())
            if error is not None:
                messagebox.showerror('Pam Price Tools - BŁĄD:', f'Wystapił bład odczytu danych z pliku CSV - {error}')
            self.csv_columns_count_verify = len(self.csv_data_verify[0])
            for item in self.csv_data_verify[0]:
                self.multiple_input_listbox.insert(tk.END, item)
            self.verify_start_button.config(state='enabled')
            self.empty_start_button.config(state='enabled')
            self.start_button_verify.config(state='enabled')
            self.save_button.config(state='enabled')
            self.verify_frame.update()

    def get_excel_path(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path is not None and path != '' and path != '.xlsx':
            self.excel_raw_path.set(path)
            self.excel_data, error = excel_handler.xlsx_read(self.excel_raw_path.get())
            if error is not None:
                messagebox.showerror('Pam Price Tools - BŁĄD:',
                                     f'Wykryto błąd podczas odczytu danych pliku excel - {error}')
                return
            self.excel_columns_count, error = excel_handler.get_columns_count(self.excel_data)
            if error is not None:
                messagebox.showerror('Pam Price Tools - BŁĄD:',
                                     f'Wykryto błąd podczas obliczania ilości kolumn pliku excel - {error}')
                return
            self.update_data_exchange_frame()

    def update_prices(self):
        if (self.csv_search_column.get() is None or self.excel_search_column.get() is None or
                self.csv_starting_row.get() is None or self.excel_starting_row.get() is None):
            messagebox.showwarning('Pam Price Tools - OSTRZEŻENIE:', f'Ustaw wymianę danych pomiędzy plikami')
            return
        try:
            if int(self.csv_search_column.get()) <= 0 or int(self.excel_search_column.get()) <= 0 or int(
                    self.csv_starting_row.get()) <= 0 or int(self.excel_starting_row.get()) <= 0:
                messagebox.showwarning('Pam Price Tools - OSTRZEŻENIE:', f'Ustaw wymianę danych pomiędzy plikami')
                return
        except Exception:
            messagebox.showwarning('Pam Price Tools - OSTRZEŻENIE:', f'Ustaw wymianę danych pomiędzy plikami')
            return
        if self.excel_raw_path.get() is None or self.csv_raw_path.get() is None:
            messagebox.showwarning('Pam Price Tools - OSTRZEŻENIE:', f'Wybierz pliki CSV i Excell')
            return
        if (self.excel_raw_path.get() == '' or self.excel_raw_path.get() == '.xlsx'
                or self.csv_raw_path.get() == '' or self.csv_raw_path.get() == '.csv'):
            messagebox.showwarning('Pam Price Tools - OSTRZEŻENIE:', f'Wybierz pliki CSV i Excell')
            return
        error, self.headers, self.csv_data = csv_handler.read_csv(self.csv_raw_path.get())
        if error is not None:
            messagebox.showerror('Pam Price Tools - BŁĄD:', f'Wykryto błąd podczas ładowania pliku CSV - {error}')
            return
        self.csv_columns_count = len(self.csv_data[0])
        self.excel_data, error = excel_handler.xlsx_read(self.excel_raw_path.get())
        if error is not None:
            messagebox.showerror('Pam Price Tools - BŁĄD:',
                                 f'Wykryto błąd podczas odczytu danych pliku excel - {error}')
            return
        self.excel_columns_count, error = excel_handler.get_columns_count(self.excel_data)
        if error is not None:
            messagebox.showerror('Pam Price Tools - BŁĄD:',
                                 f'Wykryto błąd podczas obliczania ilości kolumn pliku excel - {error}')
            return
        position_found = False
        positions_not_found = []
        self.start_button.config(state='disabled')
        self.start_button_verify.config(state='disabled')
        was_error = False
        search_values, error = excel_handler.get_column_data(self.excel_data, int(self.excel_search_column.get()))
        if error is not None:
            messagebox.showerror('Pam Price Tools - BŁĄD:',
                                 f'Wykryto błąd podczas odczytu danych kolumny pliku excel - {error}')
            return
        for i in range(int(self.csv_starting_row.get()) - 1, len(self.csv_data)):
            self.csv_progress_counter.set(f"{i} / {len(self.csv_data)}")
            self.progress_bar_csv['value'] = (i / len(self.csv_data)) * 100
            self.price_update_frame.update()
            search = self.csv_data[i][int(self.csv_search_column.get()) - 1]
            for j in range(int(self.excel_starting_row.get()), len(self.excel_data)):
                if search == search_values[j]:
                    position_found = True
                    excel_row = excel_handler.get_row_data(self.excel_data, j + 2)
                    if excel_row[1] is not None:
                        messagebox.showerror('Pam Price Tools - BŁĄD:',
                                             f'Wykryto błąd podczas odczytu danych wiersza '
                                             f'pliku excel - {excel_row[1]}')
                        return
                    if excel_row is not None:
                        if self.csv_discount_value_column.get() != '' and self.excel_discount_value_column.get() != '':
                            try:
                                csv_column = int(self.csv_discount_value_column.get()) - 1
                                excel_column = int(self.excel_discount_value_column.get()) - 1
                                if not pd.isna(excel_row[0].iloc[excel_column]):
                                    self.csv_data[i][csv_column] = str(excel_row[0].iloc[excel_column]
                                                                       * 100).replace('.', ',')
                                else:
                                    self.csv_data[i][csv_column] = '0.0'
                            except Exception as e:
                                messagebox.showwarning('Pam Price Tools - OSTRZEŻENIE:',
                                                       f'Błąd wymiany danych (wartosć rabatu): {e}')
                                was_error = True
                                pass
                        if str(self.csv_discount_group_column.get()) != '' and str(
                                self.excel_discount_group_column.get()) != '':
                            try:
                                csv_column = int(self.csv_discount_group_column.get()) - 1
                                excel_column = int(self.excel_discount_group_column.get()) - 1
                                if not pd.isna(excel_row[0].iloc[excel_column]):
                                    self.csv_data[i][csv_column] = str(excel_row[0].iloc[excel_column])
                            except Exception as e:
                                messagebox.showwarning('Pam Price Tools - OSTRZEŻENIE:',
                                                       f'Błąd wymiany danych (Grupa rabatu): {e}')
                                was_error = True
                                pass
                        if self.csv_base_price_column.get() != '' and self.excel_base_price_column.get() != '':
                            try:
                                csv_column = int(self.csv_base_price_column.get()) - 1
                                excel_column = int(self.excel_base_price_column.get()) - 1
                                if not pd.isna(excel_row[0].iloc[excel_column]):
                                    self.csv_data[i][csv_column] = str(excel_row[0].iloc[excel_column]).replace('.',
                                                                                                                ',')
                                else:
                                    self.csv_data[i][csv_column] = '0.0'
                            except Exception as e:
                                messagebox.showwarning('Pam Price Tools - OSTRZEŻENIE:',
                                                       f'Błąd wymiany danych (Cena bazowa): {e}')
                                was_error = True
                                pass
                        if (self.csv_catalogue_price_column.get() != '' and
                                self.excel_catalogue_price_column.get() != ''):
                            try:
                                csv_column = int(self.csv_catalogue_price_column.get()) - 1
                                excel_column = int(self.excel_catalogue_price_column.get()) - 1
                                if not pd.isna(excel_row[0].iloc[excel_column]) and excel_row[0].iloc[
                                        excel_column] != 0.0:
                                    self.csv_data[i][csv_column] = str(excel_row[0].iloc[excel_column]).replace('.',
                                                                                                                ',')
                                else:
                                    self.csv_data[i][csv_column] = self.csv_data[i][
                                        int(self.csv_base_price_column.get()) - 1]
                            except Exception as e:
                                messagebox.showwarning('Pam Price Tools - OSTRZEŻENIE:',
                                                       f'Błąd wymiany danych (Cena katalogowa): {e}')
                                was_error = True
                                pass
                    break
            if was_error:
                if messagebox.askyesno('Pam Price Tools', 'Czy anulować wymianę danych?'):
                    break
                was_error = False

            if not position_found:
                positions_not_found.append(search)
            if position_found:
                position_found = False

        if self.price_update_delete_option.get() and len(positions_not_found) > 0:
            for i in range(len(positions_not_found)):
                updated_data, error = csv_handler.remove_entry(self.csv_data, int(self.csv_search_column.get()) - 1,
                                                               positions_not_found[i])
                if error is not None:
                    messagebox.showerror('Pam Price Tools - BŁĄD:',
                                         f'Wykryto błąd podczas usuwania wiersza pliku CSV - {error}')
                if updated_data is not None:
                    self.csv_data = updated_data
        if self.save_not_found_index.get() and len(positions_not_found) > 0:
            messagebox.showinfo('PPT - File save', 'Podaj lokalizację zapisu indexów nie znalezionych')
            not_found_path = filedialog.asksaveasfilename(defaultextension='.txt', filetypes=[('Text', '*.txt')])
            if not_found_path is not None and not_found_path != '' and not_found_path != '.txt':
                txt_handler.save_txt(not_found_path, positions_not_found)
        messagebox.showinfo('PPT - File save', 'Podaj lokalizację zapisu pliku CSV')
        new_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if new_path is not None and new_path != '' and new_path != '.csv':
            result = csv_handler.save_csv(new_path, self.headers, self.csv_data)
            if not result[0]:
                messagebox.showerror('Pam Price Tools', f'Wykryto błąd zapisu danych pliku CSV: {result[1]}')
        self.csv_progress_counter.set(f'Wielkość po operacji: {len(self.csv_data) + 1}')
        self.progress_bar_csv['value'] = 100
        self.start_button.config(state='enabled')
        self.start_button_verify.config(state='enabled')
        self.price_update_frame.update()

    def csv_data_verify_method(self):
        multiple_length = self.multiple_input_listbox.curselection()
        self.multiple_inputs_selected = [self.multiple_input_listbox.get(i) for i in multiple_length]
        print(len(self.csv_data_verify))
        for selected_item in self.multiple_inputs_selected:
            duplicates, error = csv_handler.multiple_data_check(self.csv_data_verify,
                                                                self.csv_data_verify[0].index(selected_item))
            if error is not None:
                messagebox.showerror('Pam Price Tools - BŁĄD',
                                     f'Wystąpił błąd sprawdzania duplikatów dla wybranej kolumny - {error}')
            for duplicate in duplicates:
                data_list = []
                for i in range(len(self.csv_data_verify)):
                    if self.csv_data_verify[i][self.csv_data_verify[0].index(selected_item)] == duplicate:
                        data_list.append(self.csv_data_verify[i])
                if len(data_list) > 0:
                    self.show_data_dialog(data_list, skip_all=True)
                    self.root.wait_window(self.dialog)
                    if self.skip_all:
                        self.skip_all = False
                        break
            else:
                messagebox.showinfo('Pam Price Tools', 'Nie znaleziono zwielokrotnionych wpisów')
            if self.skip_all:
                self.skip_all = False
                break

    def save_verified_file(self):
        new_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if new_path is not None and new_path != '' and new_path != '.csv':
            csv_handler.save_csv(new_path, self.headers_verify, self.csv_data_verify)

    def csv_data_empty_verify_method(self):
        multiple_length = self.multiple_input_listbox.curselection()
        self.multiple_inputs_selected = [self.multiple_input_listbox.get(i) for i in multiple_length]
        print(len(self.csv_data_verify))
        data_list = []
        for selected_item in self.multiple_inputs_selected:
            for i in range(len(self.csv_data_verify)):
                if self.csv_data_verify[i][self.csv_data_verify[0].index(selected_item)] == '' or \
                        self.csv_data_verify[i][self.csv_data_verify[0].index(selected_item)] is None:
                    data_list.append(self.csv_data_verify[i])
            if len(data_list) > 0:
                self.show_data_dialog(data_list)
                self.root.wait_window(self.dialog)
            else:
                messagebox.showinfo('Pam Price Tools', 'Nie znaleziono pustych wpisów')

    def show_data_dialog(self, data_list, skip_all=False):
        self.dialog = tk.Toplevel(self.root)
        self.dialog.title("Wykryte dane")
        self.dialog.geometry('800x600')
        self.dialog.iconbitmap(os.path.join(os.getcwd(), 'app_icon.ico'))
        self.dialog.grid_columnconfigure(0, weight=1)
        self.dialog.grid_rowconfigure(1, weight=1)

        label = ttk.Label(self.dialog, text="Zduplikowane dane:")
        label.grid(row=0, column=0, padx=10, pady=10, sticky='w')

        frame = ttk.Frame(self.dialog)
        frame.grid(row=1, column=0, padx=10, pady=10, sticky='nsew')
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)
        data_listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE)
        data_listbox.grid(row=0, column=0, sticky='nsew')
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=data_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        data_listbox.config(yscrollcommand=scrollbar.set)
        for item in data_list:
            data_listbox.insert(tk.END, item)
        button_frame = ttk.Frame(self.dialog)
        button_frame.grid(row=2, column=0, padx=10, pady=10, sticky='ew')
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)

        delete_button = ttk.Button(button_frame, text="Usuń",
                                   command=lambda: self.delete_selected(data_listbox, self.dialog))
        delete_button.grid(row=0, column=0, padx=10, pady=10, sticky='ew')

        skip_button = ttk.Button(button_frame, text="Pomiń", command=self.dialog.destroy)
        skip_button.grid(row=0, column=1, padx=10, pady=10, sticky='ew')

        if skip_all:
            skip_all_button = ttk.Button(button_frame, text="Pomiń wszystkie", command=self.skip_all_func)
            skip_all_button.grid(row=0, column=2, padx=10, pady=10, sticky='ew')

        if len(data_list) < 1:
            delete_button.config(state='disabled')
        self.center_dialog(self.dialog)

    def show_verification_dialog(self, data_list):
        self.dialog = tk.Toplevel(self.root)
        self.dialog.title("Wykryte dane")
        self.dialog.geometry('800x600')
        self.dialog.iconbitmap(os.path.join(os.getcwd(), 'app_icon.ico'))
        self.dialog.grid_columnconfigure(0, weight=1)
        self.dialog.grid_rowconfigure(1, weight=1)

        label = ttk.Label(self.dialog, text="Błędne:")
        label.grid(row=0, column=0, padx=10, pady=10, sticky='w')

        frame = ttk.Frame(self.dialog)
        frame.grid(row=1, column=0, padx=10, pady=10, sticky='nsew')
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)
        data_listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE)
        data_listbox.grid(row=0, column=0, sticky='nsew')
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=data_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        data_listbox.config(yscrollcommand=scrollbar.set)
        for item in data_list:
            data_listbox.insert(tk.END, item)
        button_frame = ttk.Frame(self.dialog)
        button_frame.grid(row=2, column=0, padx=10, pady=10, sticky='ew')
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        skip_button = ttk.Button(button_frame, text="OK", command=self.dialog.destroy)
        skip_button.grid(row=0, column=1, padx=10, pady=10, sticky='nsew')
        self.center_dialog(self.dialog)

    def delete_selected(self, listbox, dialog):
        selected_indices = listbox.curselection()
        selected_items = [listbox.get(i) for i in selected_indices]
        for item in selected_items:
            for row in self.csv_data_verify:
                if row[0] == item[0]:
                    self.csv_data_verify.remove(row)
                    print(len(self.csv_data_verify))
        dialog.destroy()

    def verify_prices(self):
        if (self.csv_search_column.get() is None or self.excel_search_column.get() is None or
                self.csv_starting_row.get() is None or self.excel_starting_row.get() is None):
            messagebox.showwarning('Pam Price Tools - OSTRZEŻENIE:', f'Ustaw wymianę danych pomiędzy plikami')
            return
        try:
            if int(self.csv_search_column.get()) <= 0 or int(self.excel_search_column.get()) <= 0 or int(
                    self.csv_starting_row.get()) <= 0 or int(self.excel_starting_row.get()) <= 0:
                messagebox.showwarning('Pam Price Tools - OSTRZEŻENIE:', f'Ustaw wymianę danych pomiędzy plikami')
                return
        except Exception:
            messagebox.showwarning('Pam Price Tools - OSTRZEŻENIE:', f'Ustaw wymianę danych pomiędzy plikami')
            return
        if self.excel_raw_path.get() is None or self.csv_raw_path.get() is None:
            messagebox.showwarning('Pam Price Tools - OSTRZEŻENIE:', f'Wybierz pliki CSV i Excell')
            return
        if (self.excel_raw_path.get() == '' or self.excel_raw_path.get() == '.xlsx'
                or self.csv_raw_path.get() == '' or self.csv_raw_path.get() == '.csv'):
            messagebox.showwarning('Pam Price Tools - OSTRZEŻENIE:', f'Wybierz pliki CSV i Excell')
            return
        error, self.headers, self.csv_data = csv_handler.read_csv(self.csv_raw_path.get())
        if error is not None:
            messagebox.showerror('Pam Price Tools - BŁĄD:', f'Wykryto błąd podczas ładowania pliku CSV - {error}')
            return
        self.csv_columns_count = len(self.csv_data[0])
        self.excel_data, error = excel_handler.xlsx_read(self.excel_raw_path.get())
        if error is not None:
            messagebox.showerror('Pam Price Tools - BŁĄD:',
                                 f'Wykryto błąd podczas odczytu danych pliku excel - {error}')
            return
        self.excel_columns_count, error = excel_handler.get_columns_count(self.excel_data)
        if error is not None:
            messagebox.showerror('Pam Price Tools - BŁĄD:',
                                 f'Wykryto błąd podczas obliczania ilości kolumn pliku excel - {error}')
            return
        position_found = False
        positions_not_found = []
        differences = []
        self.start_button.config(state='disabled')
        self.start_button_verify.config(state='disabled')
        was_error = False
        search_values, error = excel_handler.get_column_data(self.excel_data, int(self.excel_search_column.get()))
        if error is not None:
            messagebox.showerror('Pam Price Tools - BŁĄD:',
                                 f'Wykryto błąd podczas odczytu danych kolumny pliku excel - {error}')
            return
        for i in range(int(self.csv_starting_row.get()) - 1, len(self.csv_data)):
            self.csv_progress_counter.set(f"{i} / {len(self.csv_data)}")
            self.progress_bar_csv['value'] = (i / len(self.csv_data)) * 100
            self.price_update_frame.update()
            search = self.csv_data[i][int(self.csv_search_column.get()) - 1]
            for j in range(int(self.excel_starting_row.get()), len(self.excel_data)):
                if search == search_values[j]:
                    difference = False
                    position_found = True
                    excel_row = excel_handler.get_row_data(self.excel_data, j + 2)
                    if excel_row[1] is not None:
                        messagebox.showerror('Pam Price Tools - BŁĄD:', f'Wykryto błąd podczas odczytu '
                                                                        f'danych wiersza pliku excel - {excel_row[1]}')
                        return
                    if excel_row is not None:
                        if str(self.csv_discount_value_column.get()) != '' and str(
                                self.excel_discount_value_column.get()) != '':
                            try:
                                csv_column = int(self.csv_discount_value_column.get()) - 1
                                excel_column = int(self.excel_discount_value_column.get()) - 1
                                if self.csv_data[i][csv_column] != str(excel_row[0].iloc[excel_column]
                                                                       * 100).replace('.', ','):
                                    if (not pd.isna(excel_row[0].iloc[excel_column]) and
                                            self.csv_data[i][csv_column] != '0.0'):
                                        difference = True
                                        print(f'{self.csv_data[i][csv_column]} != {str(excel_row[0].iloc[excel_column]
                                                                                       * 100).replace('.', ',')}')
                            except Exception as e:
                                messagebox.showwarning('Pam Price Tools - OSTRZEŻENIE:',
                                                       f'Błąd weryfikacji danych (wartosć rabatu): {e}')

                                was_error = True
                                pass
                        if str(self.csv_discount_group_column.get()) != '' and str(
                                self.excel_discount_group_column.get()) != '':
                            try:
                                csv_column = int(self.csv_discount_group_column.get()) - 1
                                excel_column = int(self.excel_discount_group_column.get()) - 1
                                if self.csv_data[i][csv_column] != str(excel_row[0].iloc[excel_column]):
                                    difference = True
                                    print(f'{self.csv_data[i][csv_column]} != {str(excel_row[0].iloc[excel_column])}')
                            except Exception as e:
                                messagebox.showwarning('Pam Price Tools - OSTRZEŻENIE:',
                                                       f'Błąd weryfikacji danych (Grupa rabatu): {e}')
                                was_error = True
                                pass
                        if str(self.csv_base_price_column.get()) != '' and str(
                                self.excel_base_price_column.get()) != '':
                            try:
                                csv_column = int(self.csv_base_price_column.get()) - 1
                                excel_column = int(self.excel_base_price_column.get()) - 1
                                if self.csv_data[i][csv_column] != str(excel_row[0].iloc[excel_column]).replace('.',
                                                                                                                ','):
                                    difference = True
                                    print(f'{self.csv_data[i][csv_column]} != {str(excel_row[0].iloc[excel_column]
                                                                                   ).replace('.', ',')}')
                            except Exception as e:
                                messagebox.showwarning('Pam Price Tools - OSTRZEŻENIE:',
                                                       f'Błąd weryfikacji danych (Cena bazowa): {e}')
                                was_error = True
                                pass
                        if str(self.csv_catalogue_price_column.get()) != '' and str(
                                self.excel_catalogue_price_column.get()) != '':
                            try:
                                csv_column = int(self.csv_catalogue_price_column.get()) - 1
                                excel_column = int(self.excel_catalogue_price_column.get()) - 1
                                if not pd.isna(excel_row[0].iloc[excel_column]) and excel_row[0].iloc[
                                        excel_column] != 0.0:
                                    if (self.csv_data[i][csv_column] !=
                                            str(excel_row[0].iloc[excel_column]).replace('.', ',')):
                                        print(f'{self.csv_data[i][csv_column]} != {str(excel_row[0].iloc[excel_column]
                                                                                       ).replace('.', ',')}')
                                        difference = True
                                else:
                                    if (self.csv_data[i][csv_column] !=
                                            self.csv_data[i][int(self.csv_base_price_column.get()) - 1]):
                                        print(f'{self.csv_data[i][csv_column]} != '
                                              f'{self.csv_data[i][int(self.csv_base_price_column.get()) - 1]}')
                                        difference = True
                            except Exception as e:
                                messagebox.showwarning('Pam Price Tools - OSTRZEŻENIE:',
                                                       f'Błąd weryfikacji danych (Cena katalogowa): {e}')
                                was_error = True
                                pass
                    if difference:
                        differences.append(self.csv_data[i])
                    break
            if was_error:
                if messagebox.askyesno('Pam Price Tools', 'Czy anulować weryfikację danych?'):
                    break
            if not position_found:
                positions_not_found.append(search)
            if position_found:
                position_found = False
        if not was_error:
            if self.save_not_found_index.get() and len(positions_not_found) > 0:
                messagebox.showinfo('PPT - File save', 'Podaj lokalizację zapisu indexów nie znalezionych')
                not_found_path = filedialog.asksaveasfilename(defaultextension='.txt', filetypes=[('Text', '*.txt')])
                if not_found_path is not None and not_found_path != '' and not_found_path != '.txt':
                    txt_handler.save_txt(not_found_path, positions_not_found)

            self.csv_progress_counter.set(f'Wielkość po operacji: {len(self.csv_data) + 1}')
            self.progress_bar_csv['value'] = 100

            if len(differences) > 0:
                self.show_verification_dialog(differences)
            else:
                messagebox.showinfo('Pam Price Tools', 'Nie znaleziono róznic w kolumnach')
        self.start_button.config(state='enabled')
        self.start_button_verify.config(state='enabled')
        self.price_update_frame.update()

    @staticmethod
    def center_dialog(dialog):
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        dialog.geometry(f'{width}x{height}+{x}+{y}')

    def skip_all_func(self):
        self.skip_all = True
        self.dialog.destroy()


if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()
