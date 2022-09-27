from gensim.models import KeyedVectors
from config import FASTTEXT_MODEL_FULL


class FastTextCorrector:
    def __init__(self):
        self.model_name = FASTTEXT_MODEL_FULL
        self.model = KeyedVectors.load(FASTTEXT_MODEL_FULL)
        self.surnames = set()
        self.names = set()
        self.last_names = set()
        self.words = set()

    def find_similarity(self, origin, word):
        # similarity = self.model.similarity(origin, word)
        lower_similarity = self.model.similarity(origin.lower(), word)
        title_similarity = self.model.similarity(origin.title(), word)
        return max([lower_similarity, title_similarity])

    def find_most_similar(self, origin, words=None, get_similarity=False):
        if words is None:
            words = self.words
        word_similarity = {word: self.find_similarity(origin, word) for word in words}
        most_similar = max(word_similarity.items(), key=lambda x: x[-1])[0]
        if get_similarity:
            return word_similarity.get(most_similar)
        return most_similar

    def split_names(self, real_info):
        for info in real_info:
            name_info = info.rstrip().split(" ")
            next_index = 0
            if len(name_info[0]) <= 3:
                surname = " ".join(name_info[:1])
                next_index += 2
            else:
                surname = name_info[0]
                next_index += 1
            name = name_info[next_index]
            next_index += 1
            last_name = " ".join(name_info[next_index:])
            self.surnames.add(surname)
            self.names.add(name)
            self.last_names.add(last_name)

            words = self.surnames.copy()
            words.update(self.names)
            words.update(self.last_names)
            self.words = words


if __name__ == "__main__":
    a = {"1": 123, "2": 174}
    for el in a.items():
        print(el)
    max_v = max(a.items(), key=lambda x: x[-1])
    print(max_v)
