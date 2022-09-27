import os
FASTTEXT_MODEL_FOLDER = "fasttext_model"
JAMSPELL_MODEL_FOLDER = "jamspell_model"

FASTTEXT_MODEL = "araneum_none_fasttextskipgram_300_5_2018.model"
JAMSPELL_MODEL = "my_model_2.bin"

EXCEL_ORIGIN_FILENAME = "/home/kurskiysv/PycharmProjects/WorkSpace/Книга2.xlsx"
EXCEL_RESULT_FILENAME = "/home/kurskiysv/PycharmProjects/WorkSpace/Книга2.xlsx"

CHECK_DATA_INFO = ("F", "Лист1 (2)", 3, 346)  # column id, start row, end row
REAL_DATA_INFO = ("A", "Лист1 (2)", 3, 316)  # column id, start row, end row
OUTPUT_DATA_INFO = ("D", "Лист1 (2)", 3, 316)  # column id, start row, end row

ERROR_COLOR = "FFFF00"
WARN_COLOR = "FFA500"
ORIGIN_COLOR = "FFFFFF"

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
FASTTEXT_DIR = os.path.join(BASE_DIR, FASTTEXT_MODEL_FOLDER)
JAMSPELL_DIR = os.path.join(BASE_DIR, JAMSPELL_MODEL_FOLDER)

FASTTEXT_MODEL_FULL = os.path.join(FASTTEXT_DIR, FASTTEXT_MODEL)
JAMSPELL_MODEL_FULL = os.path.join(JAMSPELL_DIR, JAMSPELL_MODEL)

SIMILARITY_VALUE = 0.6
OUTPUT_SIMILARITY = 0.45
