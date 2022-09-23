from gensim.models import KeyedVectors
if __name__ == "__main__":
    # with open("text.txt", "r") as text:
    #     sentences = text.readlines()

    model = KeyedVectors.load("araneum_none_fasttextskipgram_300_5_2018.model")

    print(model.similarity("Абдлатипова", "Абдуллатипова"))
    print(model.similarity("Абдлатипова", "Птмат"))
    print(model.similarity("Абдлатипова", "Абакумова"))
    print(model.similarity("Патимат", "Патимат"))
