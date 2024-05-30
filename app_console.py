import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import csv_handler
import xlsx_handler
import pandas as pd

class PriceUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PamPriceUpdate")

        # Notebook
        self.notebook = tk.ttk.Notebook(self.root)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        # CSV Tab
        self.frame_csv = tk.Frame(self.notebook, padx=10, pady=10)
        self.frame_csv.grid(row=0, column=0, sticky="nsew")

        self.frame_csv_options = tk.Frame(self.frame_csv, padx=10, pady=10)
        self.frame_csv_options.grid(row=1, column=0, sticky="nsew")

        self.notebook.add(self.frame_csv, text="CSV")

        # Excel Tab
        self.frame_excel = tk.Frame(self.notebook, padx=10, pady=10)
        self.frame_excel.grid(row=0, column=0, sticky="nsew")

        self.frame_excel_options = tk.Frame(self.frame_excel, padx=10, pady=10)
        self.frame_excel_options.grid(row=1, column=0, sticky="nsew")

        self.notebook.add(self.frame_excel, text="Excel")

        # Price Tab
        self.frame_price = tk.Frame(self.notebook, padx=10, pady=10)
        self.frame_price.grid(row=0, column=0, sticky="nsew")

        self.frame_price_options = tk.Frame(self.frame_price, padx=10, pady=10)
        self.frame_price_options.grid(row=1, column=0, sticky="nsew")

        self.notebook.add(self.frame_price, text="Aktualizacja cen")

        # CSV File Selection
        self.csv_path = tk.StringVar()
        tk.Label(self.frame_csv, text="Wybierz plik CSV:").grid(row=0, column=0, sticky="w")
        self.csv_entry = tk.Entry(self.frame_csv, textvariable=self.csv_path, width=50)
        self.csv_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(self.frame_csv, text="Przeglądaj", command=self.browse_csv).grid(row=0, column=2, padx=5, pady=5)

        # CSV Options
        self.csv_start_row = tk.StringVar()
        tk.Label(self.frame_csv_options, text="CSV Start Row:").grid(row=0, column=0, sticky="w")
        self.csv_start_row_entry = tk.Entry(self.frame_csv_options, textvariable=self.csv_start_row, width=10)
        self.csv_start_row_entry.grid(row=0, column=1, padx=5, pady=5)

        # Excel File Selection
        self.excel_path = tk.StringVar()
        tk.Label(self.frame_excel, text="Select Excel File:").grid(row=0, column=0, sticky="w")
        self.excel_entry = tk.Entry(self.frame_excel, textvariable=self.excel_path, width=50)
        self.excel_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(self.frame_excel, text="Browse", command=self.browse_excel).grid(row=0, column=2, padx=5, pady=5)

        # Excel Options
        self.excel_start_row = tk.StringVar()
        tk.Label(self.frame_excel_options, text="Excel Start Row:").grid(row=0, column=0, sticky="w")
        self.excel_start_row_entry = tk.Entry(self.frame_excel_options, textvariable=self.excel_start_row, width=10)
        self.excel_start_row_entry.grid(row=0, column=1, padx=5, pady=5)

        self.csv_data = None
        self.excel_data = None
        self.csv_header = None

        #Price options
        self.price_csv_search_index = tk.StringVar()
        self.price_xlsx_search_index = tk.StringVar()
        self.price_csv_discount_value = tk.StringVar()
        self.price_xlsx_discount_value = tk.StringVar()
        self.price_csv_discount_group = tk.StringVar()
        self.price_xlsx_discount_group = tk.StringVar()
        self.price_csv_base_price = tk.StringVar()
        self.price_xlsx_base_price = tk.StringVar()
        self.price_csv_cat_price = tk.StringVar()
        self.price_xlsx_cat_price = tk.StringVar()

        tk.Label(self.frame_price_options, text="Podaj numery kolumn wyszukiwania").grid(row=0, column=1, sticky="w")
        tk.Label(self.frame_price_options, text="      CSV").grid(row=0, column=0, sticky="w")
        tk.Label(self.frame_price_options, text="      EXCEL").grid(row=0, column=2, sticky="w")
        self.price_csv_search_index_entry = tk.Entry(self.frame_price_options, textvariable=self.price_csv_search_index, width=10)
        self.price_csv_search_index_entry.grid(row=1, column=0, padx=5, pady=5)
        self.price_excel_search_index_entry = tk.Entry(self.frame_price_options, textvariable=self.price_xlsx_search_index,
                                                     width=10)
        self.price_excel_search_index_entry.grid(row=1, column=2, padx=5, pady=5)

        tk.Label(self.frame_price_options, text="Podaj numery kolumn wartości rabatu").grid(row=2, column=1, sticky="w")
        tk.Label(self.frame_price_options, text="      CSV").grid(row=2, column=0, sticky="w")
        tk.Label(self.frame_price_options, text="      EXCEL").grid(row=2, column=2, sticky="w")
        self.price_csv_discount_value_entry = tk.Entry(self.frame_price_options, textvariable=self.price_csv_discount_value,
                                                     width=10)
        self.price_csv_discount_value_entry.grid(row=3, column=0, padx=5, pady=5)
        self.price_xlsx_discount_value_entry = tk.Entry(self.frame_price_options,
                                                       textvariable=self.price_xlsx_discount_value,
                                                       width=10)
        self.price_xlsx_discount_value_entry.grid(row=3, column=2, padx=5, pady=5)

        tk.Label(self.frame_price_options, text="Podaj numery kolumn grupy rabatowej").grid(row=4, column=1, sticky="w")
        tk.Label(self.frame_price_options, text="      CSV").grid(row=4, column=0, sticky="w")
        tk.Label(self.frame_price_options, text="      EXCEL").grid(row=4, column=2, sticky="w")
        self.price_csv_discount_group_entry = tk.Entry(self.frame_price_options, textvariable=self.price_csv_discount_group,
                                                     width=10)
        self.price_csv_discount_group_entry.grid(row=5, column=0, padx=5, pady=5)
        self.price_xlsx_discount_group_entry = tk.Entry(self.frame_price_options,
                                                       textvariable=self.price_xlsx_discount_group,
                                                       width=10)
        self.price_xlsx_discount_group_entry.grid(row=5, column=2, padx=5, pady=5)

        tk.Label(self.frame_price_options, text="Podaj numery kolumn ceny bazowej").grid(row=6, column=1, sticky="w")
        tk.Label(self.frame_price_options, text="      CSV").grid(row=6, column=0, sticky="w")
        tk.Label(self.frame_price_options, text="      EXCEL").grid(row=6, column=2, sticky="w")
        self.price_csv_base_price_entry = tk.Entry(self.frame_price_options, textvariable=self.price_csv_base_price,
                                                     width=10)
        self.price_csv_base_price_entry.grid(row=7, column=0, padx=5, pady=5)
        self.price_xlsx_base_price_entry = tk.Entry(self.frame_price_options,
                                                       textvariable=self.price_xlsx_base_price,
                                                       width=10)
        self.price_xlsx_base_price_entry.grid(row=7, column=2, padx=5, pady=5)

        tk.Label(self.frame_price_options, text="Podaj numery kolumn ceny katalogowej").grid(row=8, column=1, sticky="w")
        tk.Label(self.frame_price_options, text="      CSV").grid(row=8, column=0, sticky="w")
        tk.Label(self.frame_price_options, text="      EXCEL").grid(row=8, column=2, sticky="w")
        self.price_csv_cat_price_entry = tk.Entry(self.frame_price_options, textvariable=self.price_csv_cat_price,
                                                     width=10)
        self.price_csv_cat_price_entry.grid(row=9, column=0, padx=5, pady=5)
        self.price_xlsx_cat_price_entry = tk.Entry(self.frame_price_options,
                                                       textvariable=self.price_xlsx_cat_price,
                                                       width=10)
        self.price_xlsx_cat_price_entry.grid(row=9, column=2, padx=5, pady=5)
        self.remove_not_found_items = tk.BooleanVar()
        tk.Checkbutton(self.frame_price_options, text='Usuń pozycje spoza cennika', command=self.update_remove).grid(row=10, column=1, padx=5, pady=5)
        tk.Button(self.frame_price_options, text="Aktualizuj ceny w bazie", command=self.price_update).grid(row=11, column=1, padx=5, pady=5)

        #progressbars
        progress_bar_csv = ttk.Progressbar(self.frame_price_options, orient='horizontal', length=300,
                                           mode="determinate")
        progress_bar_csv.grid(row=12, column=1, padx=5, pady=5)
        tk.Label(self.frame_price_options, text='Postęp pliku CSV').grid(row=12, column=0, sticky="w")
        tk.Label(self.frame_price_options, text='Postęp pliku XLSX').grid(row=13, column=0, sticky="w")
        progress_bar_xlsx = ttk.Progressbar(self.frame_price_options, orient='horizontal', length=300,
                                            mode="determinate")
        progress_bar_xlsx.grid(row=13, column=1, padx=5, pady=5)
        self.remove_option = False
        self.position_found = False
        self.not_found_list = []

    def update_remove(self):
        self.remove_option = True

    def browse_csv(self):
        csv_file = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if csv_file is not None:
            self.csv_path.set(csv_file)
            _, self.csv_header, self.csv_data = csv_handler.read_csv(csv_file)

    def browse_excel(self):
        excel_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if excel_file is not None:
            self.excel_path.set(excel_file)
            self.excel_data, _ = xlsx_handler.xlsx_read(excel_file)

    def price_update(self):
        starting_excel_row = int(self.excel_start_row.get())
        starting_csv_row = int(self.csv_start_row.get())
        current_csv_index = starting_csv_row
        csv_search_indes = int(self.price_csv_search_index.get()) - 1
        position_found = None
        search_values, s = xlsx_handler.get_column_data(self.excel_data, int(self.price_xlsx_search_index.get()))
        progress_bar_csv = ttk.Progressbar(self.frame_price_options, orient='horizontal', length=300,
                                           mode="determinate")
        progress_bar_csv.grid(row=12, column=1, padx=5, pady=5)
        tk.Label(self.frame_price_options, text='Postęp pliku CSV').grid(row=12, column=0, sticky="w")
        tk.Label(self.frame_price_options, text='Postęp pliku XLSX').grid(row=13, column=0, sticky="w")
        progress_bar_xlsx = ttk.Progressbar(self.frame_price_options, orient='horizontal', length=300,
                                            mode="determinate")
        progress_bar_xlsx.grid(row=13, column=1, padx=5, pady=5)
        for i in range(starting_csv_row, len(self.csv_data)):
            progress_bar_csv['value'] = (i / len(self.csv_data)) * 100
            self.frame_price_options.update()
            search = self.csv_data[i][csv_search_indes]
            for j in range(starting_excel_row, len(self.excel_data)):
                progress_bar_xlsx['value'] = (j / len(self.excel_data)) * 100
                self.frame_price_options.update()
                if search_values[j] == search:
                    print(search_values[j])
                    print(search)
                    self.position_found = True
                    position_found = xlsx_handler.get_row_data(self.excel_data, j + 2)
                    if position_found is not None:
                        print(f'Dane z excela: {position_found}')
                        try:
                            csv_column = int(self.price_csv_discount_value.get()) - 1
                            xlsx_column = int(self.price_xlsx_discount_value.get()) - 1
                            if not pd.isna(position_found[0].iloc[xlsx_column]):
                                self.csv_data[i][csv_column] = str(position_found[0].iloc[xlsx_column] * 100).replace('.',
                                                                                                                  ',')
                            else:
                                self.csv_data[i][csv_column] = '0'
                        except Exception as e:
                            print(f'Discount group: {e}')
                            pass
                        try:
                            csv_column = int(self.price_csv_discount_group.get()) - 1
                            xlsx_column = int(self.price_xlsx_discount_group.get()) - 1
                            self.csv_data[i][csv_column] = str(position_found[0].iloc[xlsx_column])
                        except Exception as e:
                            print(f'Discount value: {e}')
                            pass
                        try:
                            csv_column = int(self.price_csv_base_price.get()) - 1
                            xlsx_column = int(self.price_xlsx_base_price.get()) - 1
                            self.csv_data[i][csv_column] = str(position_found[0].iloc[xlsx_column]).replace('.', ',')
                        except Exception as e:
                            print(f'Base price: {e}')
                            pass
                        try:
                            csv_column = int(self.price_csv_cat_price.get()) - 1
                            xlsx_column = int(self.price_xlsx_cat_price.get()) - 1
                            if not pd.isna(position_found[0].iloc[xlsx_column]):
                                self.csv_data[i][csv_column] = str(position_found[0].iloc[xlsx_column]).replace('.',
                                                                                                            ',')
                            else:
                                csv_column = int(self.price_csv_base_price.get()) - 1
                                self.csv_data[i][csv_column] = self.csv_data[i][csv_column]
                        except Exception as e:
                            print(f'Cat price: {e}')
                            pass
                    print(f'Dane po modyfiukacji: {self.csv_data[i]}')
                    break
            if not self.position_found:
                self.not_found_list.append(search)
            if self.position_found:
                self.position_found = False
        if self.remove_option and len(self.not_found_list) > 0:
            print(len(self.csv_data))
            for item in self.not_found_list:
                updated_data, _ = csv_handler.remove_entry(self.csv_data, csv_search_indes, item)
                if updated_data is not None:
                    self.csv_data = updated_data
                else:
                    print('Remove NONE')
            print(len(self.csv_data))
        csv_file_raw = filedialog.asksaveasfile(filetypes=[("CSV Files", "*.csv")])
        if csv_file_raw is not None:
            csv_file = str(csv_file_raw.name) + '.csv'
            if csv_file is not None and csv_file != '.csv':
                csv_handler.save_csv(csv_file, self.csv_header, self.csv_data)





if __name__ == "__main__":
    root = tk.Tk()
    app = PriceUpdaterApp(root)
    root.mainloop()
