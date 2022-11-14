import tkinter as tk
import tkinter.filedialog as fd
from tkinter.messagebox import showinfo


def select_file():
    filetypes = (
        ('excel old files', '*.xls'),
        ('excel files', '*.xlsx')
    )

    filename = fd.askopenfilename(
        title='Выберите файл',
        initialdir='/',
        filetypes=filetypes)
    return filename


def select_directory():
    directory = fd.askdirectory(
        title='Куда сохранить накладные?',
        mustexist=True
    )
    return directory


def on_save():
    if len(file_names[0]) == 0 or len(file_names[1]) == 0:
        showinfo(
            title='Ошибка',
            message="Необходимо выбрать оба файла"
        )
        return
    directory = select_directory()
    f.process_fruits_facade(input_file_names=file_names, output_directory=directory)
    showinfo(
        title='Успех',
        message="Накладные сгенерированы"
    )


class FileChooser:
    def __init__(self, title, button_names, generate_and_save_fun):
        self.root = tk.Tk()
        self.root.geometry("750x350")
        self.root.title(title)
        self.button_names = button_names
        self.file_names = ["" for _ in button_names]
        self.labels = []
        for index, file_name in enumerate(self.file_names):
            label = tk.Label(self.root, text=file_name)
            label.grid(row=index, column=1)
            self.labels.append(label)
        for index, button_name in enumerate(self.button_names):
            button = tk.Button(self.root, text=button_name, command=lambda inx=index: self.__choose_file(inx))
            button.grid(row=index)
        save_button = tk.Button(self.root, text="Сгенерировать накладные",
                                command=lambda: self.__on_save(generate_and_save_fun))
        save_button.grid(row=len(self.button_names))

    def __choose_file(self, button_index):
        filename = select_file()
        self.file_names[button_index] = filename
        self.labels[button_index].configure(text=filename)

    def __on_save(self, generate_and_save_fun):
        if not all(self.file_names):
            showinfo(
                title='Ошибка',
                message="Необходимо выбрать все файлы"
            )
            return
        directory = select_directory()
        generate_and_save_fun(self.file_names, directory)
        showinfo(
            title='Успех',
            message="Накладные сгенерированы"
        )

    def show(self):
        self.root.mainloop()


if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("750x350")
    root.title("Фрукты и Сухофрукты")

    fruits_label = tk.Label(root, text="")
    fruits_label.grid(row=0, column=1)
    dried_fruits_label = tk.Label(root, text="")
    dried_fruits_label.grid(row=1, column=1)

    tk.Button(root, text="Фрукты", command=select_fruits).grid(row=0)
    tk.Button(root, text="Сухофрукты", command=select_dried_fruits).grid(row=1)
    tk.Button(root, text="Сгенерировать накладные", command=on_save).grid(row=2)

    root.mainloop()
