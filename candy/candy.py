import os
from argparse import ArgumentParser
from argparse import FileType
from datetime import date
from typing import List

from utils import text_utils, workbook_utils
from candy.candy_name_text_processor import Model


class Candy:
    def __init__(self, art, name, unit, count=None, cost=None, summary=None):
        self.art = art
        self.name = name
        self.count = count
        self.unit = unit
        self.cost = cost
        self.summary = summary


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
    parser.add_argument("-ic", "--input-candy", required=True, help="Накладная \"кондитерка\"", type=FileType('r'))
    parser.add_argument("-o",
                        "--output",
                        required=False,
                        help="Папка, в которую сохранятся накладные",
                        type=dir_path,
                        default=".")
    return parser.parse_args()


def read_1c_worksheet(input_worksheet):
    candies = []
    for row_number in range(2, input_worksheet.max_row):
        name = input_worksheet.cell(row_number, 6).value
        if name:
            candies.append(Candy(art=input_worksheet.cell(row_number, 3).value,
                                 name=name,
                                 unit=input_worksheet.cell(row_number, 4).value))
    return candies


def read_candy_worksheet(input_worksheet):
    candies = []
    for row_number in range(2, input_worksheet.max_row):
        name = input_worksheet.cell(row_number, 4).value
        if name:
            candies.append(Candy(art=input_worksheet.cell(row_number, 2).value,
                                 name=name,
                                 count=input_worksheet.cell(row_number, 5).value,
                                 unit=input_worksheet.cell(row_number, 6).value,
                                 cost=input_worksheet.cell(row_number, 7).value,
                                 summary=input_worksheet.cell(row_number, 8).value))
    return candies


def calculate_result(one_c_candies, candy_candies) -> List[workbook_utils.OneCRow]:
    identicals = 0
    likes = 0
    news = 0
    one_c_rows = []
    model = Model()
    model.fit(one_c_candies)
    for candy in candy_candies:
        if not text_utils.candy_name_split_regex.match(candy.name):
            raise ValueError("{} не по формату".format(candy.name))

        prediction, similarity, position, vector = model.predict(candy)
        similarity = round(similarity * 100)
        min_similarity = 95

        if prediction.art == candy.art:
            one_c_rows.append(map_candy_to_one_c_row(candy, "Одинаковы {}%".format(similarity)))
            identicals += 1
            # print("{} -> {} | {} IDENTICAL".format(prediction.name, candy.name, similarity))
        elif similarity >= min_similarity:
            one_c_rows.append(map_candy_to_one_c_row(candy, "Похож {}% на {}".format(similarity, prediction.art)))
            likes += 1
            # print("{} -> {} | {} LIKE".format(prediction.name, candy.name, similarity).replace("\n", ""))
            # array_1c = model.result.toarray()
            # array_v = vector.toarray()[0]
            # for feature_name, point_1c, point in zip(model.pipeline.get_feature_names_out(), array_1c[position],
            #                                          array_v):
            #     if (point_1c > 0 or point > 0) and (point_1c != point):
            #         print(feature_name, point_1c, point)
        else:
            one_c_rows.append(map_candy_to_one_c_row(candy, ""))
            news += 1
            # print("NOT EXIST -> {} NOT EXIST".format(candy.name))
    print("IDENTICALS:{} | LIKES:{} | NEW ONES:{}".format(identicals, likes, news))
    return one_c_rows


def map_candy_to_one_c_row(candy: Candy, similarity: str) -> workbook_utils.OneCRow:
    return workbook_utils.OneCRow(
        art=candy.art,
        code=similarity,
        name=candy.name,
        count=candy.count,
        cost=candy.cost,
        summary=candy.summary
    )


if __name__ == '__main__':
    args = init_args()

    input_1c_worksheet = workbook_utils.load_input_worksheet(args.input_1c.name)
    input_candy_worksheet = workbook_utils.load_input_worksheet(args.input_candy.name)

    one_c_rows = read_1c_worksheet(input_1c_worksheet)
    candy_rows = read_candy_worksheet(input_candy_worksheet)

    rows = calculate_result(one_c_rows, candy_rows)
    workbook = workbook_utils.generate_1c_sheet(rows, "Кондитерка")
    workbook.save(generate_filename(args.output, "Кондитерка"))
