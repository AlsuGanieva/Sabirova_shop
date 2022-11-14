from ui.common_file_choose_ui import FileChooser

import fruits.fruits as f


if __name__ == "__main__":
    FileChooser(title="Фрукты и Сухофрукты", button_names=["Фрукты", "Сухофрукты"],
                generate_and_save_fun=lambda file_names, directory:
                f.process_fruits_facade(input_file_names=file_names,
                                        output_directory=directory)).show()
