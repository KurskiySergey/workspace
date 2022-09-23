import jamspell
from config import JAMSPELL_MODEL_FULL


class JamSpell:
    def __init__(self):
        self.corrector = jamspell.TSpellCorrector()
        self.corrector.LoadLangModel(modelFile=JAMSPELL_MODEL_FULL)

    def fix_sentence(self, sentence):
        result = self.corrector.FixFragment(sentence)
        return result

    # def save_data_to_excel(self, data):
    #     self.excel_reader.save_excel_data(data=data, filename=f"{self.excel_reader.filename}_result.xlsx")
