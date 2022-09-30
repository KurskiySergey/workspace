import jamspell
from config import JAMSPELL_MODEL_FULL, JAMSPELL_MODEL_ALPHABET, JAMSPELL_MODEL_TRAIN


class JamSpell:
    def __init__(self):
        self.corrector = jamspell.TSpellCorrector()

    def fix_sentence(self, sentence):
        result = self.corrector.FixFragment(sentence)
        return result

    @staticmethod
    def get_n_gramms(clear_data, n_gramm):
        n_gramms = set()
        for data in clear_data:
            words = data.split(" ")
            for word in words:
                for i in range(len(word) - n_gramm):
                    n_gramms.add(word[i:i + n_gramm])
        return n_gramms

    @staticmethod
    def save_clear_txt_data(clear_data: list):

        splitted_data = []
        double_splitted = []
        for cl_data in clear_data:
            cl_data = cl_data.split(" ")
            double_splitted.append(" ".join(cl_data[:2]))
            double_splitted.append(" ".join(cl_data[1:]))
            for cl in cl_data:
                splitted_data.append(cl)
        n_gramms = []
        for i in range(4, 10):
            n_gramms.append("\n".join(JamSpell.get_n_gramms(clear_data, n_gramm=i)))

        clear_data = "\n".join(clear_data) + "\n" + "\n".join(n_gramms) + "\n" + "\n".join(double_splitted)
        # clear_data = "\n".join(clear_data)
        with open(JAMSPELL_MODEL_TRAIN, "w") as text:
            text.write(clear_data)

    def train_model(self, clear_data):
        self.save_clear_txt_data(clear_data)
        self.corrector.TrainLangModel(JAMSPELL_MODEL_TRAIN, alphabetFile=JAMSPELL_MODEL_ALPHABET,
                                      modelFile=JAMSPELL_MODEL_FULL)

    def load_model(self):
        self.corrector.LoadLangModel(modelFile=JAMSPELL_MODEL_FULL)

