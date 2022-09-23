import os
FASTTEXT_MODEL_FOLDER = "fasttext_model"
JAMSPELL_MODEL_FOLDER = "jamspell_model"

FASTTEXT_MODEL = "araneum_none_fasttextskipgram_300_5_2018.model"
JAMSPELL_MODEL = "my_model.bin"

EXCEL_ORIGIN_FILENAME = ""
EXCEL_RESULT_FILENAME = ""

CHECK_DATA_INFO = ("F", 3, 346)  # column id, start row, end row
REAL_DATA_INFO = ("A", 3, 315)  # column id, start row, end row

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
FASTTEXT_DIR = os.path.join(BASE_DIR, FASTTEXT_MODEL_FOLDER)
JAMSPELL_DIR = os.path.join(BASE_DIR, JAMSPELL_MODEL_FOLDER)

FASTTEXT_MODEL_FULL = os.path.join(FASTTEXT_DIR, FASTTEXT_MODEL)
JAMSPELL_MODEL_FULL = os.path.join(JAMSPELL_DIR, JAMSPELL_MODEL)

