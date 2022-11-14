from ui.common_file_choose_ui import FileChooser
import candy.candy as c

if __name__ == '__main__':
    FileChooser(title="Накладные на кондитерку", button_names=["Выгрузка из 1С", "Накладная \"кондитерка\""],
                generate_and_save_fun=lambda file_names, directory:
                c.process_candy_facade(input_file_names=file_names, output_directory=directory)).show()
