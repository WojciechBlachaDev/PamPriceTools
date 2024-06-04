import ctypes
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox


def create_main_window(icon_path, window_tittle, app_id):
    try:
        root = tk.Tk()
        root.title(window_tittle)
        root.iconbitmap(icon_path)
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
        return True, root
    except Exception as e:
        return False, e


def create_main_notebook(root, row=0, column=0, sticky='nsew'):
    try:
        notebook = ttk.Notebook(root)
        notebook.grid(row=row, column=column, sticky=sticky)
        return True, notebook
    except Exception as e:
        return False, e


def create_file_frame(notebook, row, column, sticky, frame_description, fields):
    try:
        frame = ttk.Frame(notebook)
        frame.grid(row=row, column=column, sticky=sticky)
        notebook.add(frame, text=frame_description)
        for i in range(len(fields)):
            row_label = ttk.Label(frame, text=fields[0])
            row_label.grid(row=i, column=0, sticky='w')
            row_entry = ttk.Entry(frame, textvariable=fields[1], width=100)
            row_entry.grid(row=i, column=1, padx=5, pady=5)
            row_browse_button = ttk.Button(frame, text='Przegladaj pliki', command=fields[2])
            row_browse_button.grid(row=i, column=2, padx=5, pady=5)
        return True, frame
    except Exception as e:
        return False, e

def get_path(filetypes):
    try:
        return True, filedialog.askopenfilename(filetypes=filetypes)
    except Exception as e:
        return False, e
