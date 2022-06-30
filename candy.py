from openpyxl import Workbook
from openpyxl import load_workbook
from argparse import ArgumentParser
from argparse import FileType
from datetime import date
from fuzzywuzzy import process
from fuzzywuzzy import fuzz
from hungarian_algorithm import algorithm
import os


def get_date():
    today = date.today()
    return today.strftime("%d-%m-%Y")


def generate_filename(output, name):
    return "{output}/{name}-{date}.xlsx".format(output=output, name=name, date=get_date())


def dir_path(string):
    if os.path.isdir(string):
        return string
    else:
        raise NotADirectoryError(string)


def init_args():
    parser = ArgumentParser(description="Накладные на кондитерку")
    parser.add_argument("-i1", "--input-1c", required=True, help="Выгрузка из 1С", type=FileType('r'))
    parser.add_argument("-iс", "--input-candy", required=True, help="Накладная \"кондитерка\"", type=FileType('r'))
    parser.add_argument("-o",
                        "--output",
                        required=False,
                        help="Папка, в которую сохранятся накладные",
                        type=dir_path,
                        default=".")
    return parser.parse_args()


def load_input_file(input_path):
    input_workbook = load_workbook(input_path)
    input_worksheet = input_workbook.active
    return input_worksheet


def read_1c_worksheet(input_worksheet):
    rows = []
    for row_number in range(2, input_worksheet.max_row):
        name = input_worksheet.cell(row_number, 6).value
        if name:
            rows.append((input_worksheet.cell(row_number, 3).value, name))
    return rows


def read_candy_worksheet(input_worksheet):
    rows = []
    for row_number in range(2, input_worksheet.max_row):
        name = input_worksheet.cell(row_number, 4).value
        if name:
            rows.append((input_worksheet.cell(row_number, 2).value,
                         name,
                         input_worksheet.cell(row_number, 5).value,
                         input_worksheet.cell(row_number, 6).value,
                         input_worksheet.cell(row_number, 7).value,
                         input_worksheet.cell(row_number, 8).value))
    return rows


def one_c_row_tuple_processor(s, force_ascii=False):
    return process.default_processor(s[1], force_ascii)


def generate_new_sheet(rows, title, file_name):
    output_workbook = Workbook()
    output_worksheet = output_workbook.active

    output_worksheet.title = title
    output_worksheet["A1"] = "арт"
    output_worksheet["B1"] = "штрих"
    output_worksheet["C1"] = "наименование"
    output_worksheet["D1"] = "Кол-во"
    output_worksheet["E1"] = "Ед.изм."
    output_worksheet["F1"] = "цена"
    output_worksheet["G1"] = "сумма"

    for row_index, row in enumerate(rows):
        output_worksheet.cell(row_index + 2, 1, value=row[0])
        output_worksheet.cell(row_index + 2, 3, value=row[1])
        output_worksheet.cell(row_index + 2, 4, value=row[2])
        output_worksheet.cell(row_index + 2, 5, value=row[3])
        output_worksheet.cell(row_index + 2, 6, value=row[4])
        output_worksheet.cell(row_index + 2, 7, value=row[5])
        if len(row) > 6:
            output_worksheet.cell(row_index + 2, 2, value=row[6]) #temp

    output_workbook.save(file_name)


def calculate_result(one_c_rows, candy_rows):
    identicals = 0
    likes = 0
    news = 0
    rows = []
    for candy_row in candy_rows:
        result = process.extractOne(candy_row, one_c_rows, processor=one_c_row_tuple_processor,
                                    scorer=fuzz.token_sort_ratio, score_cutoff=85)
        if result and result[0][0] == candy_row[0]:
            print("{} -> {} | {} IDENTICAL".format(result[0][1], candy_row[1], result[1]))
            rows.append([candy_row[0], candy_row[1], candy_row[2], candy_row[3], candy_row[4], candy_row[5], "X{}".format(result[1])])
            identicals += 1
        elif result:
            rows.append(candy_row)
            print("{} -> {} | {} LIKE".format(result[0][1], candy_row[1], result[1]))
            likes += 1
        else:
            rows.append(candy_row)
            # print("NOT EXIST -> {} NOT EXIST".format(candy_row[1]))
            news += 1
    print("IDENTICALS:{} | LIKES:{} | NEW ONES:{}".format(identicals, likes, news))
    return rows


if __name__ == '__main__':
    args = init_args()

    input_1c_worksheet = load_input_file(args.input_1c.name)
    input_candy_worksheet = load_input_file(args.input_candy.name)

    one_c_rows = read_1c_worksheet(input_1c_worksheet)
    candy_rows = read_candy_worksheet(input_candy_worksheet)

    rows = calculate_result(one_c_rows, candy_rows)
    generate_new_sheet(rows, "Кондитерка", generate_filename(args.output, "Кондитерка"))
