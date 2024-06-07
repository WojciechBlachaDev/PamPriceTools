import ctypes
import os
from tkinterdnd2 import DND_FILES
import tkinterdnd2
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import pandas as pd
from scripts import csv_handler
from scripts import excel_handler
from scripts import json_handler
from scripts import txt_handler
from scripts import pdf_tools


class App:
    def __init__(self, main_root):
        # Gui elements
        # Main
        self.special_price_checkbox = None
        self.root = None
        self.dialog = None
        self.notebook = None
        # Pages
        self.verify_frame = None
        self.files_frame = None
        self.price_update_frame = None
        self.data_exchange_frame = None
        # CheckBox
        self.calculate_checkbox = None
        self.save_not_found_check_box = None
        self.multiple_inputs_selected = None
        self.delete_option_check_box = None
        # Buttons
        self.start_button_verify = None
        self.save_button = None
        self.empty_start_button = None
        self.start_button = None
        self.verify_start = None
        self.verify_start_button = None
        # ListBox
        self.multiple_input_listbox = None
        self.excel_no_discount_column = None
        # ProgressBar
        self.verify_progressbar = None
        self.progress_bar_csv = None
        self.progress_bar_excel = None
        # Labels
        self.progress_bar_excel_value = None
        self.progress_bar_csv_value = None

        # VARIABLES
        # CSV
        self.csv_starting_row = tk.StringVar()
        self.csv_raw_path = tk.StringVar()
        self.csv_descriptions = []
        self.csv_data = None
        self.csv_data_verify = None
        self.headers = None
        self.headers_verify = None
        self.csv_verify_path = tk.StringVar()
        self.csv_columns_count = 0
        self.csv_columns_count_verify = 0
        self.csv_search_column = None
        self.csv_discount_value_column = None
        self.csv_discount_group_column = None
        self.csv_base_price_column = None
        self.csv_catalogue_price_column = None
        # EXCEL
        self.excel_data = None
        self.excel_starting_row = tk.StringVar()
        self.excel_raw_path = tk.StringVar()
        self.excel_search_column = None
        self.excel_discount_value_column = None
        self.excel_discount_group_column = None
        self.excel_base_price_column = None
        self.excel_catalogue_price_column = None
        # PDF
        self.pdf_raw_path = tk.StringVar()
        self.pdf_data = None
        self.pdf_standard_discounts = None
        self.pdf_special_prices = None
        self.pdf_search_phrase_1 = tk.StringVar()
        self.pdf_search_phrase_2 = tk.StringVar()
        self.pdf_search_phrase_3 = tk.StringVar()
        # OTHERS
        self.save_not_found_index = tk.BooleanVar()
        self.calculate_catalogue_price = tk.BooleanVar()
        self.include_special_prices = tk.BooleanVar()
        self.price_update_delete_option = tk.BooleanVar()
        self.settings_name = ''
        self.current_settings = None
        self.items_to_delete = []
        self.startup = True
        self.delete_bool = False
        self.skip_all = False
        # SETTINGS WITH LOAD FUNCTION
        self.settings = json_handler.load_settings(os.path.join(os.getcwd(), 'exchange_settings.json'))
        if self.settings == {}:
            json_handler.save_settings(self.settings, os.path.join(os.getcwd(), 'exchange_settings.json'))
        self.excel_columns_count = 0
        # APP STARTUP
        self.set_main_window(main_root)
        self.set_notebook()
        self.set_file_page()
        self.set_data_exchange_page()
        self.csv_progress_counter = tk.StringVar()
        self.excel_progress_counter = tk.StringVar()
        self.set_price_update_page()
        self.set_verify_page()

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
            excel_label = ttk.Label(self.files_frame, text='Wybierz plik Excel')
            excel_label.grid(row=1, column=0, sticky='w')
            pdf_label = ttk.Label(self.files_frame, text='Wybierz plik cennika PDF (Opcjonalnie)')
            pdf_label.grid(row=2, column=0, sticky='w')
            pdf_label1 = ttk.Label(self.files_frame, text='Podaj nagłówek rozpoczęcia rabatów standard')
            pdf_label1.grid(row=3, column=0, sticky='w')
            pdf_label2 = ttk.Label(self.files_frame, text='Podaj nagłówek zakończenia rabatów standard')
            pdf_label2.grid(row=4, column=0, sticky='w')
            pdf_label3 = ttk.Label(self.files_frame, text='Podaj nagłówek zakończenia cen specjalnych')
            pdf_label3.grid(row=5, column=0, sticky='w')

            csv_entry = ttk.Entry(self.files_frame, textvariable=self.csv_raw_path, width=100)
            csv_entry.grid(row=0, column=1, padx=5, pady=5)
            excel_entry = ttk.Entry(self.files_frame, textvariable=self.excel_raw_path, width=100)
            excel_entry.grid(row=1, column=1, padx=5, pady=5)
            pdf_entry = ttk.Entry(self.files_frame, textvariable=self.pdf_raw_path, width=100)
            pdf_entry.grid(row=2, column=1, padx=5, pady=5)
            pdf_search_phrase_1 = ttk.Entry(self.files_frame, textvariable=self.pdf_search_phrase_1, width=100)
            pdf_search_phrase_1.grid(row=3, column=1, padx=5, pady=5)
            pdf_search_phrase_2 = ttk.Entry(self.files_frame, textvariable=self.pdf_search_phrase_2, width=100)
            pdf_search_phrase_2.grid(row=4, column=1, padx=5, pady=5)
            pdf_search_phrase_3 = ttk.Entry(self.files_frame, textvariable=self.pdf_search_phrase_3, width=100)
            pdf_search_phrase_3.grid(row=5, column=1, padx=5, pady=5)

            csv_button = ttk.Button(self.files_frame, text='Przeglądaj pliki', command=self.get_csv_path)
            csv_button.grid(row=0, column=2, padx=5, pady=5)
            excel_button = ttk.Button(self.files_frame, text='Przeglądaj pliki', command=self.get_excel_path)
            excel_button.grid(row=1, column=2, padx=5, pady=5)
            pdf_button = ttk.Button(self.files_frame, text='Przeglądaj pliki', command=self.get_pdf_path)
            pdf_button.grid(row=2, column=2, padx=5, pady=5)
            load_pdf_button = ttk.Button(self.files_frame, text='Ładuj dane cennika PDF', command=self.load_pdf)
            load_pdf_button.grid(row=6, column=1, padx=5, pady=5)

            self.files_frame.drop_target_register(DND_FILES)
            self.files_frame.dnd_bind('<<Drop>>', self.on_drop)
        except Exception as e:
            print(e)
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

        label = ttk.Label(self.data_exchange_frame, text='Wybierz kolumnę z opcją grupy nierabatowej w pliku excel')
        label.grid(row=10, column=1, padx=5, pady=5)
        self.excel_no_discount_column = ttk.Combobox(self.data_exchange_frame, values=excel_columns)
        self.excel_no_discount_column.grid(row=11, column=1, padx=5, pady=5)

        self.set_descriptions_rows('Wprowadź wiersz startowy danych w plikach', 12)
        excel_starting_row = ttk.Entry(self.data_exchange_frame, textvariable=self.excel_starting_row)
        excel_starting_row.grid(row=13, column=2, padx=5, pady=5)
        csv_starting_row = ttk.Entry(self.data_exchange_frame, textvariable=self.csv_starting_row)
        csv_starting_row.grid(row=13, column=0, padx=5, pady=5)

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
        self.verify_frame.grid(row=0, column=0, sticky='nsew', )
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
        self.verify_frame.drop_target_register(DND_FILES)
        self.verify_frame.dnd_bind('<<Drop>>', self.on_drop2)

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
        calculate_label = ttk.Label(self.price_update_frame, text='Kalkuluj wartości ceny katologowej')
        calculate_label.grid(row=5, column=0, padx=5, pady=5)
        special_price_label = ttk.Label(self.price_update_frame, text='Uwzględniaj ceny specjalne cennika PDF')
        special_price_label.grid(row=6, column=0, padx=5, pady=5)
        self.calculate_checkbox = ttk.Checkbutton(self.price_update_frame, variable=self.calculate_catalogue_price)
        self.calculate_checkbox.grid(row=5, column=1, padx=5, pady=5)
        self.special_price_checkbox = ttk.Checkbutton(self.price_update_frame, variable=self.include_special_prices)
        self.special_price_checkbox.grid(row=6, column=1, padx=5, pady=5)
        self.start_button = ttk.Button(self.price_update_frame,
                                       text='Aktualizuj bazę', command=self.update_prices)
        self.start_button.grid(row=7, column=1, padx=10, pady=10)
        self.start_button_verify = ttk.Button(self.price_update_frame,
                                              text='Porównaj ceny z plików', command=self.verify_prices)
        self.start_button_verify.grid(row=8, column=1, padx=10, pady=10)

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
            'csv_start': self.csv_starting_row.get(),
            'no_discount_column': self.excel_no_discount_column.get()
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
                self.excel_no_discount_column.set(self.settings[key]['no_discount_column'])

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

        label = ttk.Label(self.data_exchange_frame, text='Wybierz kolumnę z opcją grupy nierabatowej w pliku excel')
        label.grid(row=10, column=1, padx=5, pady=5)
        self.excel_no_discount_column = ttk.Combobox(self.data_exchange_frame, values=excel_columns)
        self.excel_no_discount_column.grid(row=11, column=1, padx=5, pady=5)

        self.set_descriptions_rows('Wprowadź wiersz startowy danych w plikach', 12)
        excel_starting_row = ttk.Entry(self.data_exchange_frame, textvariable=self.excel_starting_row)
        excel_starting_row.grid(row=13, column=2, padx=5, pady=5)
        csv_starting_row = ttk.Entry(self.data_exchange_frame, textvariable=self.csv_starting_row)
        csv_starting_row.grid(row=13, column=0, padx=5, pady=5)

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
            self.update_data_exchange_frame()

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

    def get_pdf_path(self):
        path = filedialog.askopenfilename(filetypes=[("Pdf Files", "*.pdf")])
        if path is not None and path != '' and path != '.pdf':
            self.pdf_raw_path.set(path)
        else:
            messagebox.showerror('Pam Price Tools - BŁĄD', 'Błąd wczytywania ścieżki do pliku PDF!')

    def load_pdf(self):
        if self.pdf_raw_path.get() == '' or self.pdf_raw_path.get() is None:
            messagebox.showerror('Pam Price Tools', 'Sciezka do pliku PDF jest pusta!')
            return
        if self.pdf_search_phrase_1.get() == '' or self.pdf_search_phrase_1.get() is None:
            messagebox.showerror('Pam Price Tools', 'Podaj dane wyszukiwania w pliku PDF')
            return
        if self.pdf_search_phrase_2.get() == '' or self.pdf_search_phrase_2.get() is None:
            messagebox.showerror('Pam Price Tools', 'Podaj dane wyszukiwania w pliku PDF')
            return
        if self.pdf_search_phrase_3.get() == '' or self.pdf_search_phrase_3.get() is None:
            messagebox.showerror('Pam Price Tools', 'Podaj dane wyszukiwania w pliku PDF')
            return
        self.pdf_data, error = pdf_tools.read_file(self.pdf_raw_path.get())
        if error is not None:
            messagebox.showerror('Pam Price Tools - BŁĄD',
                                 f'Wykryto błąd podczas ładowania danych pliku PDF - {error}')
        self.pdf_standard_discounts, error = pdf_tools.read_standard_discounts(self.pdf_data,
                                                                               self.pdf_search_phrase_1.get(),
                                                                               self.pdf_search_phrase_2.get())
        if error is not None:
            messagebox.showerror('Pam Price Tools - BŁĄD',
                                 f'Wykryto błąd podczas szukania rabatów standardowych - {error}')
        self.pdf_special_prices, error = pdf_tools.read_non_standard_prices(self.pdf_data,
                                                                            self.pdf_search_phrase_2.get(),
                                                                            self.pdf_search_phrase_3.get())
        if messagebox.askyesno('Pam Price Tools', 'Czy wygenerować nowy cennik w programie excel?'):
            excel_path = os.path.join(os.getcwd(), 'sample_excel.xlsm')
            macro_name = 'Makro1'
            my_app, workbook = excel_handler.open_workbook(excel_path)
            readed_data, error = excel_handler.xlsx_read(excel_path)
            excel_handler.fill_discount_table_2(readed_data, self.pdf_standard_discounts, workbook, 'Warunki Handlowe')
            excel_handler.fill_empty_cells_in_column_c(workbook, 'Warunki Handlowe')
            excel_handler.start_macro(workbook, macro_name)
            workbook.close()

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
        if self.include_special_prices.get():
            if len(self.pdf_data) <= 0 or self.pdf_data is None:
                messagebox.showerror('Pam Price Tools - BŁĄD:',
                                     f'Najpierw załaduj plik PDF')
                return
            if len(self.pdf_special_prices) <= 0 or self.pdf_special_prices is None:
                messagebox.showerror('Pam Price Tools - BŁĄD:',
                                     f'Najpierw załaduj plik PDF')
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
        for i in range(int(self.csv_starting_row.get()) - 2, len(self.csv_data)):
            self.csv_progress_counter.set(f"{i + 1} / {len(self.csv_data)}")
            self.progress_bar_csv['value'] = (i / len(self.csv_data)) * 100 + 1
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
                                no_discount = pd.isna(excel_row[0].iloc[int(self.excel_no_discount_column.get()) - 1])
                                if not pd.isna(excel_row[0].iloc[excel_column]) and no_discount:
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
                                if not pd.isna(excel_row[0].iloc[excel_column]):
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
                        else:
                            if self.excel_no_discount_column.get() is None or self.excel_no_discount_column.get() == '':
                                messagebox.showerror('Pam Price Tools - BŁĄD: Nie wybrano kolumny z'
                                                     ' wartością nierabatowalną')
                                pass
                            if (pd.isna(excel_row[0].iloc[int(self.excel_no_discount_column.get()) - 1]) and
                                    self.calculate_catalogue_price.get()):
                                csv_column = int(self.csv_catalogue_price_column.get()) - 1
                                excel_column = int(self.excel_base_price_column.get()) - 1
                                base_price = float(excel_row[0].iloc[excel_column])
                                discount_value = float(excel_row[0].iloc
                                                       [int(self.excel_discount_value_column.get()) - 1])
                                price = float(base_price - (base_price * discount_value))
                                self.csv_data[i][csv_column] = str(price).replace('.', ',')
                                print(self.csv_data[i][csv_column])
                        if self.include_special_prices.get():
                            if len(self.pdf_special_prices) > 0:
                                for data in self.pdf_special_prices:
                                    if data[0] == search:
                                        self.csv_data[i][int(self.csv_base_price_column.get()) - 1] = data[1]
                                        self.csv_data[i][int(self.csv_catalogue_price_column.get()) - 1] = data[1]
                                        print(f'{search}: {data[1]}')
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
            else:
                messagebox.showwarning('Pam price Tools - OSTRZEŻENIE:', 'Nie zapisano listy '
                                                                         'indexów, które nie zostały odnalezione.')
        messagebox.showinfo('PPT - File save', 'Podaj lokalizację zapisu pliku CSV')
        new_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if new_path is not None and new_path != '' and new_path != '.csv':
            result = csv_handler.save_csv(new_path, self.headers, self.csv_data)
            if not result[0]:
                messagebox.showerror('Pam Price Tools', f'Wykryto błąd zapisu danych pliku CSV: {result[1]}')
        else:
            messagebox.showwarning('Pam Price Tools - OSTRZEŻENIE:', 'Nie zapisano wygenerowanego pliku')
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
        data_listbox.insert(tk.END, self.csv_data[0])
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
        data_listbox.insert(tk.END, self.csv_data[0])
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
        if self.include_special_prices.get():
            if len(self.pdf_data) <= 0 or self.pdf_data is None:
                messagebox.showerror('Pam Price Tools - BŁĄD:',
                                     f'Najpierw załaduj plik PDF')
                return
            if len(self.pdf_special_prices) <= 0 or self.pdf_special_prices is None:
                messagebox.showerror('Pam Price Tools - BŁĄD:',
                                     f'Najpierw załaduj plik PDF')
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
        for i in range(int(self.csv_starting_row.get()) - 2, len(self.csv_data)):
            self.csv_progress_counter.set(f"{i} / {len(self.csv_data)}")
            self.progress_bar_csv['value'] = (i / len(self.csv_data)) * 100 + 1
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
                                no_discount = pd.isna(excel_row[0].iloc[int(self.excel_no_discount_column.get()) - 1])
                                if self.csv_data[i][csv_column] != str(excel_row[0].iloc[excel_column]
                                                                       * 100).replace('.', ','):
                                    if ((not pd.isna(excel_row[0].iloc[excel_column])
                                         and self.csv_data[i][csv_column] != '0.0')
                                            and not no_discount):
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
                                if not self.include_special_prices.get():
                                    csv_column = int(self.csv_base_price_column.get()) - 1
                                    excel_column = int(self.excel_base_price_column.get()) - 1
                                    if (self.csv_data[i][csv_column] != str(excel_row[0].iloc[excel_column]).replace
                                        ('.', ',')):
                                        difference = True
                                        print(f'{self.csv_data[i][csv_column]} != {str(excel_row[0].iloc[excel_column]
                                                                                       ).replace('.', ',')}')
                                else:
                                    special_price_found = False
                                    for data in self.pdf_special_prices:
                                        if search == data[0]:
                                            if self.csv_data[i][int(self.csv_base_price_column.get()) - 1] != data[1]:
                                                difference = True
                                            special_price_found = True
                                    if not special_price_found:
                                        csv_column = int(self.csv_base_price_column.get()) - 1
                                        excel_column = int(self.excel_base_price_column.get()) - 1
                                        if self.csv_data[i][csv_column] != str(excel_row[0].iloc[excel_column]).replace(
                                                '.',
                                                ','):
                                            difference = True
                                            print(
                                                f'{self.csv_data[i][csv_column]} != {str(excel_row[0].iloc[excel_column]
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
                                if not self.include_special_prices.get():
                                    if (not pd.isna(excel_row[0].iloc[excel_column]) and
                                            excel_row[0].iloc[excel_column] != 0.0):
                                        if (self.csv_data[i][csv_column] !=
                                                str(excel_row[0].iloc[excel_column]).replace('.', ',')):
                                            difference = True
                                    else:
                                        if (self.csv_data[i][csv_column] !=
                                                self.csv_data[i][int(self.csv_base_price_column.get()) - 1]):
                                            difference = True
                                else:
                                    special_price_found = False
                                    for data in self.pdf_special_prices:
                                        if search == data[0]:
                                            if self.csv_data[i][csv_column] != data[1]:
                                                difference = True
                                            special_price_found = True
                                    if not special_price_found:
                                        if (not pd.isna(excel_row[0].iloc[excel_column]) and
                                                excel_row[0].iloc[excel_column] != 0.0):
                                            if (self.csv_data[i][csv_column] !=
                                                    str(excel_row[0].iloc[excel_column]).replace('.', ',')):
                                                difference = True
                                        else:
                                            if (self.csv_data[i][csv_column] !=
                                                    self.csv_data[i][int(self.csv_base_price_column.get()) - 1]):
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
                else:
                    messagebox.showwarning('Pam price Tools - OSTRZEŻENIE:', 'Nie zapisano listy '
                                                                             'indexów, które nie zostały odnalezione.')
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

    def on_drop(self, event):
        files = event.data.strip().split('\n')
        for dropped_file in files:
            if not dropped_file:
                continue
            if not os.path.isfile(dropped_file):
                messagebox.showinfo("Information", "Przeciągnięty obiekt nie jest plikiem.")
                return
        for dropped_file in files:
            if dropped_file.endswith('.csv'):
                self.csv_raw_path.set(dropped_file)
                error, self.headers, self.csv_data = csv_handler.read_csv(self.csv_raw_path.get())
                if error is not None:
                    messagebox.showerror('Pam Price Tools - BŁĄD:',
                                         f'Wystapił bład odczytu danych z pliku CSV - {error}')
                self.csv_columns_count = len(self.csv_data[0])
                self.update_data_exchange_frame()
            elif dropped_file.endswith('.xlsx'):
                self.excel_raw_path.set(dropped_file)
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
            elif dropped_file.endswith('.pdf'):
                self.pdf_raw_path.set(dropped_file)
            else:
                messagebox.showinfo("Information", "Przeciągnięty plik nie jest plikiem CSV ani Excel.")

    def on_drop2(self, event):
        files = event.data.strip().split('\n')
        for dropped_file in files:
            if not dropped_file:
                continue
            if not os.path.isfile(dropped_file):
                messagebox.showinfo("Information", "Przeciągnięty obiekt nie jest plikiem.")
                return
        for dropped_file in files:
            if dropped_file.endswith('.csv'):
                self.csv_verify_path.set(dropped_file)
                error, self.headers_verify, self.csv_data_verify = csv_handler.read_csv(self.csv_verify_path.get())
                if error is not None:
                    messagebox.showerror('Pam Price Tools - BŁĄD:',
                                         f'Wystapił bład odczytu danych z pliku CSV - {error}')
                self.csv_columns_count_verify = len(self.csv_data_verify[0])
                for item in self.csv_data_verify[0]:
                    self.multiple_input_listbox.insert(tk.END, item)
                self.verify_start_button.config(state='enabled')
                self.empty_start_button.config(state='enabled')
                self.start_button_verify.config(state='enabled')
                self.save_button.config(state='enabled')
                self.verify_frame.update()
            else:
                messagebox.showinfo("Information", "Przeciągnięty plik nie jest plikiem CSV ani Excel.")


if __name__ == '__main__':
    root = tkinterdnd2.Tk()
    app = App(root)
    root.mainloop()
