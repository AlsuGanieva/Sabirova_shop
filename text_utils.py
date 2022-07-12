import re

candy_name_split_regex = re.compile(r'^([а-я.\-\s]+)\s(.+)\s/(.+)/$', re.DOTALL)
candy_name_count_regex = re.compile(r'\(\s*\d+\s*\*\s*\d+\s*\)')
match_gr_regex = re.compile(r'(0[.,]\d+|\d+)\s*(г|гр|гра|грам|грамм)(\.|\s)', re.DOTALL)
punctuation_regex = re.compile(r'[\W\d]')


def preprocess_candy_string(s):
    """
    Removes "(1*4)"
    :param s: input string
    :return: preprocessed string
    """
    return candy_name_count_regex.sub(" ", s)


def split_weight(s):
    """
    Splits 400гр, 400г, 400 г etc
    :param s: input string
    :return: string without weight and weight
    """
    match = match_gr_regex.search(s)
    if match:
        weight = match.group(1)
        return match_gr_regex.sub(" ", s), weight
    else:
        return s, None


def split_candy_name_by_type_and_firm(s):
    """
    Splits candy name to type, name and firm
    :param s: input candy name E.x: "кф. Москва Нива (1*5) 5 кг. /ОАО к/к Бабаевский/"
    :return: type (кф), name (Москва Нива (1*5) 5 кг.), firm (ОАО к/к Бабаевский)
    """
    match = candy_name_split_regex.match(s)
    if match:
        return match.group(1), match.group(2), match.group(3)
    else:
        return None, None, None


def remove_punctuation_and_digits(s, placeholder=" "):
    return punctuation_regex.sub(placeholder, s)


def create_ngrams(s, n):
    return [s[i:i + n] for i in range(len(s) - n + 1)]
