import os
import ctypes
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import  messagebox
from tkinter import ttk
from scripts import csv_handler as csv
from scripts import excel_handler as excel
from scripts import json_handler as json
from scripts import txt_handler as txt


class App:
    def __init__(self, root):
        self.root = None
        self.notebook = None
        self.app_id = "BRK_Windows.PamPriceTools.PamPriceTools.version_0_0_1"
        self.window_title = 'Pam Price Tools'
        self.path_icon = os.path.join(os.getcwd(), 'app_icon.ico')

        self.csv_path = tk.StringVar()
        self.excel_path = tk.StringVar()

        self.csv_button_browse = None
        self.excel_button_browse = None

    def main_window(self, root):
        try:
            self.root = root
            self.root.title(self.window_title)
            self.root.iconbitmap(self.path_icon)
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(self.app_id)
        except Exception as e:
            tk.messagebox.showerror(f'Błąd tworzenia okna głównego aplikacji: {e}')

    def page_viewer(self):
        try:
            self.notebook = ttk.Notebook(self.root)
            self.notebook.grid(row=0, column=0, sticky='nsew')
        except Exception as e:
            tk.messagebox.showerror(f'Błąd tworzenia przeglądarki zakładek: {e}')

    def page_data_exchange(self):
        try:
            self.page_exchange_options = ttk.Frame(self.notebook)
            self.page_exchange_options.grid(row=0, column=0, sticky='nsew')
            self.notebook.add(self.page_exchange_options, text='Opcje aktualizacji cen okuć')

            csv_browse_label = ttk.Label(self.page_exchange_options, text='Wybierz plik CSV')
            csv_browse_label.grid(row=0, column=0, sticky='w')
            excel_browse_label = ttk.Label(self.page_exchange_options, text='Wybierz plik Excel')
            excel_browse_label.grid(row=1, column=0, sticky='w')

            csv_entry = ttk.Entry(self.page_exchange_options, textvariable=self.csv_path, width=100)
            csv_entry.grid(row=0, column=0, padx=5, pady=5)
            excel_entry = ttk.Entry(self.page_exchange_options, textvariable=self.excel_path, width=100)
            excel_entry.grid(row=1, column=1, padx=5, pady=5)

            self.csv_button_browse = ttk.Button(self.page_exchange_options, text='Przeglądaj pliki', command=self.browse_csv)
            self.csv_button_browse.grid(row=)



