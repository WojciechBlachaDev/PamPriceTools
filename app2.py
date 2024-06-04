import gui


class App:
    def __init__(self, root):
        self.path = None
        _, notebook = gui.create_main_notebook(root)
        fields = [('Plik CSV', self.path, self.set_path)]
        file_page = gui.create_file_frame(notebook, 0, 0, 'nsew', 'Pliki', fields)

    def set_path(self):
        result, self.path = gui.get_path([("CSV Files", "*.csv")])
        print(self.path)

if __name__ == '__main__':
    _, root = gui.create_main_window('app_icon.ico', 'test', 'test')
    app = App(root)
    root.mainloop()