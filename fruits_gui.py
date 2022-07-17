import tkinter as tk
import tkinter.filedialog as fd
from tkinter.messagebox import showinfo

import fruits

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
    for shop in fruits.get_shops():
        fruits_for_shop = []
        name_for_shop = ""
        for file_name in file_names:
            input_worksheet, name = fruits.load_input_file(file_name)
            name_for_shop += name + "-"
            input_fruits = fruits.read_data(input_worksheet, shop.column_number)
            fruits_for_shop += input_fruits
            if not shop.should_join:
                output_name = fruits.generate_filename(directory, name, shop.name)
                fruits.save_fruits(input_fruits, shop, output_name)
        if shop.should_join:
            output_name = fruits.generate_filename(directory, name_for_shop[:-1], shop.name)
            fruits.save_fruits(fruits_for_shop, shop, output_name)


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Фрукты и Сухофрукты")

    fruits_label = tk.Label(root, text="")
    fruits_label.grid(row=0, column=1)
    dried_fruits_label = tk.Label(root, text="")
    dried_fruits_label.grid(row=1, column=1)

    tk.Button(root, text="Фрукты", command=select_fruits).grid(row=0)
    tk.Button(root, text="Сухофрукты", command=select_dried_fruits).grid(row=1)
    tk.Button(root, text="Сохранить", command=on_save).grid(row=2)

    root.mainloop()
