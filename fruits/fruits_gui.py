import tkinter as tk
import tkinter.filedialog as fd
from tkinter.messagebox import showinfo

import fruits.fruits as f

file_names = ["", ""]


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


def show_message():
    showinfo(
        title='Выбран файл',
        message="filename"
    )


def select_fruits():
    filename = select_file()
    file_names[0] = filename
    fruits_label.configure(text=filename)


def select_dried_fruits():
    filename = select_file()
    file_names[1] = filename
    dried_fruits_label.configure(text=filename)


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
