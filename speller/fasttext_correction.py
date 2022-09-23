from gensim.models import KeyedVectors
from config import FASTTEXT_MODEL_FULL


class FastTextCorrector:
    def __init__(self):
        self.model_name = FASTTEXT_MODEL_FULL
        self.model = KeyedVectors.load(FASTTEXT_MODEL_FULL)

    def find_similarity(self, origin, word):
        similarity = self.model.similarity(origin, word)
        return similarity

    def find_most_similar(self, origin, words):
        word_similarity = {word: self.find_similarity(origin, word) for word in words}
        most_similar = max(word_similarity.items(), key=lambda x: x[-1])[0]
        return most_similar


if __name__ == "__main__":
    a = {"1": 123, "2": 174}
    for el in a.items():
        print(el)
    max_v = max(a.items(), key=lambda x: x[-1])
    print(max_v)

