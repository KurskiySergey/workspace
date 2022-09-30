import os
FASTTEXT_MODEL_FOLDER = "fasttext_model"
JAMSPELL_MODEL_FOLDER = "jamspell_model"

FASTTEXT_MODEL = "araneum_none_fasttextskipgram_300_5_2018.model"
JAMSPELL_MODEL = "my_model_2.bin"
ALPHABET = "alphabet_ru.txt"
TRAIN_DATA_FILENAME = "train.txt"

EXCEL_ORIGIN_FILENAME = "D:\PycharmProjects\WorkSpace\Книга2.xlsx"
EXCEL_RESULT_FILENAME = "D:\PycharmProjects\WorkSpace\Книга2.xlsx"

CHECK_DATA_INFO = ("F", "Лист1 (2)", 3, 346)  # column id, start row, end row
REAL_DATA_INFO = ("A", "Лист1 (2)", 3, 316)  # column id, start row, end row
OUTPUT_DATA_INFO = ("B", "Лист1 (2)", 3, 316)  # column id, start row, end row


ERROR_COLOR = "FFFF00"
WARN_COLOR = "FFA500"
ORIGIN_COLOR = "FFFFFF"

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
FASTTEXT_DIR = os.path.join(BASE_DIR, FASTTEXT_MODEL_FOLDER)
JAMSPELL_DIR = os.path.join(BASE_DIR, JAMSPELL_MODEL_FOLDER)

FASTTEXT_MODEL_FULL = os.path.join(FASTTEXT_DIR, FASTTEXT_MODEL)
JAMSPELL_MODEL_FULL = os.path.join(JAMSPELL_DIR, JAMSPELL_MODEL)
JAMSPELL_MODEL_ALPHABET = os.path.join(JAMSPELL_DIR, ALPHABET)
JAMSPELL_MODEL_TRAIN = os.path.join(JAMSPELL_DIR, TRAIN_DATA_FILENAME)

SIMILARITY_VALUE = 0.6
OUTPUT_SIMILARITY = 0.5


def set_qt_params(params: dict):
    global EXCEL_RESULT_FILENAME, EXCEL_ORIGIN_FILENAME, OUTPUT_SIMILARITY, CHECK_DATA_INFO, \
        REAL_DATA_INFO, OUTPUT_DATA_INFO

    EXCEL_RESULT_FILENAME = params.get("EXCEL_RESULT_FILENAME")
    EXCEL_ORIGIN_FILENAME = params.get("EXCEL_ORIGIN_FILENAME")
    OUTPUT_SIMILARITY = params.get("OUTPUT_SIMILARITY")
    CHECK_DATA_INFO = params.get("CHECK_DATA_INFO")
    REAL_DATA_INFO = params.get("REAL_DATA_INFO")
    OUTPUT_DATA_INFO = params.get("OUTPUT_DATA_INFO")

