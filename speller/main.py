from config import EXCEL_ORIGIN_FILENAME, EXCEL_RESULT_FILENAME, SIMILARITY_VALUE
from excel_reader import ExcelWorker
from jamspell_correction import JamSpell
from fasttext_correction import FastTextCorrector


class Main:
    def __init__(self):
        self.excel_reader = ExcelWorker(EXCEL_ORIGIN_FILENAME)
        self.jam_corrector = JamSpell()
        self.fasttext_corrector = FastTextCorrector()
        self.check_info = ()
        self.real_info = ()
        self.check_info = None
        self.real_info = None

    def convert_data_info(self):
        self.check_info, self.real_info = self.excel_reader.convert_data_info()

    def read_excel_data(self):
        self.excel_reader.read_data()
        self.excel_reader.read_list(list_name=self.check_info[1])
        self.check_info = self.excel_reader.get_data(column=self.check_info[0]-1, row_start=self.check_info[-2]-1,
                                                     row_end=self.check_info[-1])

        if self.check_info[1] != self.real_info[1]:
            self.excel_reader.read_list(list_name=self.real_info[1])

        self.real_info = self.excel_reader.get_data(column=self.real_info[0]-1, row_start=self.real_info[-2]-1,
                                                    row_end=self.real_info[-1])

    def fasttext_start(self, word_info):

        # word, word_type = self.find_fastcorrect_word(name="Абрамова")
        # print(self.find_candidates(word=word, word_type=word_type))
        # input()
        name_info = word_info.split(" ")
        # print(name_info)
        next_index = 0
        if len(name_info[0]) < 3:
            first_name = " ".join(name_info[:1])
            next_index += 2
        else:
            first_name = name_info[0]
            next_index += 1

        if len(first_name) <= 3:
            first_name = name_info[-1]

        word, word_type = self.find_fastcorrect_word(name=first_name)
        candidates = self.find_candidates(word=word, word_type=word_type)
        # print(f"candidates = {candidates}")
        if len(candidates) == 0:
            return []
        elif len(candidates) == 1:
            # only one candidate

            return [candidates[0][0]]
        else:
            result = False
            for candidate in candidates:
                if word_info == candidate[0]:
                    result = True
                    break
            if result:
                # no mistakes
                return [word_info]
            else:
                if first_name == name_info[-1]:
                    tail = name_info[:next_index]
                    tail.insert(next_index, first_name)
                else:
                    tail = name_info[next_index:]
                    tail.insert(0, first_name)
                result = self.jam_corrector.fix_sentence(sentence=" ".join(tail))
                find = False
                for candidate in candidates:
                    if result != word_info and candidate[0] == result:
                        find = True
                        break
                if find:
                    # find and correct
                    return [result]
                else:
                    candidates = self.name_analyzer(name_info, word, word_type, candidates=candidates)
                    return candidates
    @staticmethod
    def candidates_checker(letters, find_search_type, candidates):
        possible_candidates = []

        for candidate in candidates:
            clear_letters = letters.copy()
            checked_letters = []
            info = candidate.split(" ")
            # find_word = info[find_search_type]
            for data in info:
                # if self.fasttext_corrector.find_similarity(origin=data, word=find_word) < SIMILARITY_VALUE:
                if data.lower()[0] in clear_letters:
                    checked_letters.append(data.lower()[0])
                    clear_letters.remove(data.lower()[0])
            if len(checked_letters) == len(letters):
                ordered = True
                for i in range(len(checked_letters)):
                    if checked_letters[i] != letters[i]:
                        ordered = False
                if ordered:
                    possible_candidates.append(candidate)
        return possible_candidates

    def name_analyzer(self, name_info, search_word, search_word_type, candidates):

        if search_word_type == "surname":
            word_type = 0
        elif search_word_type == "name":
            word_type = 1
        else:
            word_type = 2
        candidates_types = {}
        for i, candidate in enumerate(candidates):
            candidate_names = candidate[0].split(" ")
            cand_type = candidate_names.index(candidate[1])
            candidate_names.pop(cand_type)
            candidate = " ".join(candidate_names)
            candidates_types[candidate] = [candidates[i].copy(), cand_type]
            candidates[i] = candidate

        if len(name_info) == 2:
            data = name_info[-1]
            if self.fasttext_corrector.find_similarity(origin=data, word=search_word) >= SIMILARITY_VALUE:
                data = name_info[0]

            if len(data) == 2:
                # has 2 optional letters
                first_letter = data[0].lower()
                second_letter = data[1].lower()
                checked_candidates = self.candidates_checker(letters=[first_letter, second_letter],
                                                             find_search_type=word_type, candidates=candidates)

            elif len(data) == 1:
                # has 1 optional letter
                letter = data.lower()
                checked_candidates = self.candidates_checker(letters=[letter],
                                                             find_search_type=word_type, candidates=candidates)
            else:
                checked_candidates = [self.fasttext_corrector.find_most_similar(origin=data, words=candidates)]

            # print(f"checked candidates: {checked_candidates}")

        elif len(name_info) > 2:
            if self.fasttext_corrector.find_similarity(origin=name_info[0], word=search_word) >= SIMILARITY_VALUE:
                data = name_info[1:]
            else:
                data = name_info[:-1]
            # print(candidates)
            # print(data)
            if len(data) == 2:
                if len(data[0]) == 1 and len(data[1]) == 1:
                    checked_candidates = Main.candidates_checker(letters=[data[0].lower(), data[1].lower()],
                                                                 find_search_type=word_type, candidates=candidates)
                else:
                    checked_candidates = [self.fasttext_corrector.find_most_similar(origin=" ".join(data),
                                                                                    words=candidates)]

            else:
                possible_candidates = []
                for name in data:
                    word, _ = self.find_fastcorrect_word(name=name)
                    for candidate in candidates:
                        if word in candidate:
                            possible_candidates.append(candidate)
                    if len(possible_candidates) != 0:
                        break
                if len(possible_candidates) != 0:
                    checked_candidates = [self.fasttext_corrector.find_most_similar(origin=" ".join(data),
                                                                                    words=possible_candidates)]
                else:
                    checked_candidates = []
            # print(checked_candidates)
            # input()

        else:
            checked_candidates = candidates

        for i, candidate in enumerate(checked_candidates):
            candidate_info = candidates_types.get(candidate)
            candidate = candidate.split(" ")
            candidate = candidate_info[0]
            # candidate = " ".join(candidate)
            checked_candidates[i] = candidate[0]

        return checked_candidates

    def find_candidates(self, word, word_type, no_name=False):
        candidates = []
        if word_type == "surname":
            word_type = 0
        elif word_type == "name":
            word_type = 1
        else:
            word_type = 2

        for info in self.check_info:
            name_info = info.rstrip().split(" ")
            # if info == "Алдонина Анна Александровна":
            #     input()
            if len(name_info) >= 3:
                for name in name_info:
                    if self.fasttext_corrector.find_similarity(origin=word, word=name) >= SIMILARITY_VALUE:
                        if no_name:
                            candidates.append(info)
                        else:
                            candidates.append([info, name])
                        break

        return list(candidates)

    def find_fastcorrect_word(self, name):
        most_similar_value = self.fasttext_corrector.find_most_similar(name)
        if most_similar_value in self.fasttext_corrector.surnames:
            return most_similar_value, "surname"
        elif most_similar_value in self.fasttext_corrector.names:
            return most_similar_value, "name"
        else:
            return most_similar_value, "last_name"

    def save_result(self, info):
        self.excel_reader.save_excel_data(data=info, filename=EXCEL_RESULT_FILENAME)

    def cross_similarity_check(self, result_info):
        for key, value in result_info.items():
            if len(value) == 1:
                data = value[0].split(" ")
                surname = data[0]
                surname_similarity = self.fasttext_corrector.find_most_similar(origin=surname, words=key.split(" "),
                                                                               get_similarity=True)
                similarity = round(self.fasttext_corrector.find_similarity(origin=key, word=value[0]), 2)
                avr_similarity = []
                for dt in data:
                    avr_similarity.append(self.fasttext_corrector.find_most_similar(origin=dt, words=key.split(" "),
                                                                                    get_similarity=True))
                avr_similarity = sum(avr_similarity) / len(avr_similarity)
                total_similarity = (avr_similarity) / 1
                value[0] += f"similarity{total_similarity}"
                result_info.update({key: [value, avr_similarity]})
            else:
                result_info.update({key: [value, 0]})
        return result_info

    def start(self):
        self.convert_data_info()
        self.read_excel_data()
        self.jam_corrector.train_model(clear_data=self.check_info)
        self.jam_corrector.load_model()
        self.fasttext_corrector.split_names(self.check_info)
        result_info = {}
        for info in self.real_info:
            info = info.replace(chr(235), "е")
            origin_info = info.split(" и ")
            for data in origin_info:
                total = []
                result = self.jam_corrector.fix_sentence(sentence=data)
                surname_info = result.split(" ")[0]
                if result != info and result in self.find_candidates(surname_info, word_type="surname", no_name=True):
                    # find and correct
                    total.append(", ".join([result]))
                else:
                    result = self.fasttext_start(word_info=data)
                    # print(result)
                    total.append(", ".join(result))
            result_info[info] = total
        result_info = self.cross_similarity_check(result_info)
        print(result_info)
        self.save_result(info=result_info)


if __name__ == "__main__":
    test = Main()
    test.start()
