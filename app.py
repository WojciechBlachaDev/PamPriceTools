from scripts import csv_handler
from scripts import excel_handler
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog


class App:
    def __init__(self, main_root):
        self.files_frame = None
        self.notebook = None
        self.root = None
        self.csv_raw_path = tk.StringVar()
        self.excel_raw_path = tk.StringVar()
        self.set_main_window(main_root)
        self.set_notebook()
        self.set_file_page()

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

    def get_csv_path(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if path is not None:
            self.csv_raw_path.set(path)

    def get_excel_path(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path is not None:
            self.excel_raw_path.set(path)



if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()
