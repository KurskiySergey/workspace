from config import EXCEL_ORIGIN_FILENAME, EXCEL_RESULT_FILENAME
from excel_reader import ExcelReader
from jamspell_correction import JamSpell
from fasttext_correction import FastTextCorrector
from config import CHECK_DATA_INFO, REAL_DATA_INFO


class Main:
    def __init__(self):
        self.excel_reader = ExcelReader(EXCEL_ORIGIN_FILENAME)
        self.jam_corrector = JamSpell()
        self.fasttext_corrector = FastTextCorrector()
        self.check_info = ()
        self.real_info = ()

    def convert_data_info(self):
        excel_en = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
