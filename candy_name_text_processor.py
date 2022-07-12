import nltk
import numpy as np
from nltk.corpus import stopwords
from nltk.stem import SnowballStemmer
from nltk.tokenize import word_tokenize
from pymorphy2 import MorphAnalyzer
from sklearn.compose import ColumnTransformer
from sklearn.feature_extraction import DictVectorizer
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

import text_utils

nltk.download('stopwords')
nltk.download('punkt')

stop_words = stopwords.words("russian")
stemmer = SnowballStemmer(language="russian")
morph = MorphAnalyzer()


def __remove_stop_words(tokens):
    filtered_tokens = []
    for token in tokens:
        if token not in ["шт", "№", "г"]:  # and token not in stop_words:
            filtered_tokens.append(token)
    return filtered_tokens


def analyze_name(name, use_ngram: bool = True):
    """

    :param use_ngram: split to ngrams by 3 characters
    :param names: list of names
    :return: Prepared corpus. List of lists of tokens
    """
    name = text_utils.preprocess_candy_string(name)
    name = text_utils.remove_punctuation_and_digits(name)
    name = name.lower()
    tokens = word_tokenize(name, language='russian')
    tokens = __remove_stop_words(tokens)
    tokens = [stemmer.stem(token) for token in tokens]
    if use_ngram:
        tokens = [ngram for token in tokens for ngram in text_utils.create_ngrams(token, n=3)]
    else:
        tokens = [morph.parse(token)[0].normal_form for token in tokens]
    return tokens


def __analyze_type_name(type_name):
    return text_utils.remove_punctuation_and_digits(type_name, '')


def split_to_features(candies):
    names_features = []
    for candy in candies:
        type_name, name, firm = text_utils.split_candy_name_by_type_and_firm(candy.name)
        name, weight = text_utils.split_weight(name + " " + firm)
        unit = __analyze_type_name(candy.unit)
        if not weight:  # or unit == "шт":
            weight = "None"
        art = ""
        if candy.art:
            art = candy.art
        feature_dict = {'type': __analyze_type_name(type_name),
                        'weight': weight,
                        'unit': unit,
                        'art': art}
        names_features.append((name, feature_dict))
    return names_features


class Model:
    def __init__(self):
        self.candies = None
        self.result = None
        self.pipeline = ColumnTransformer([
            ('tf-idf', TfidfVectorizer(analyzer=analyze_name), 0),
            ('dict', DictVectorizer(), 1)
        ])

    def fit(self, candies):
        filtered_candies = []
        for candy in candies:
            if text_utils.candy_name_split_regex.match(candy.name):
                filtered_candies.append(candy)
        self.candies = filtered_candies
        names_features = split_to_features(filtered_candies)
        self.result = self.pipeline.fit_transform(names_features)

    def print(self):
        feature_names = self.pipeline.get_feature_names_out()
        print(feature_names)

    def predict(self, candy):
        name_feature = split_to_features([candy])
        prediction_result = self.pipeline.transform(name_feature)
        result = cosine_similarity(self.result, prediction_result)
        position = np.argmax(result)
        return self.candies[position], result.flatten()[position], position, prediction_result
