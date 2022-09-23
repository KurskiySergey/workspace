import jamspell
from pyexcel_xlsx import get_data, save_data
from collections import OrderedDict


def jamsprell_test():
    # Васильев Анатолий
    # василев Анатлий
    # Абдуллаева Умуксу Умахан кызы
    # Абдуллатипова Патимат Нарсулаевна
    corrector = jamspell.TSpellCorrector()
    corrector.LoadLangModel("my_model.bin")
    # result = corrector.GetCandidates(['Абдлатипова', 'Птмат', 'Нарсулаева'], 0)
    result = corrector.FixFragment("Абдуллатипова Птмат Нарсулаевна")
    print(result)


def excel_test():
    data = get_data("test.xlsx")
    list_name = "Лист1"
    clear_data_rows = 0
    clear_data_start_point = 1
    clear_data_end_point = 1
    list_data = data.get(list_name)
    clear_data = list_data[clear_data_start_point:clear_data_end_point] if clear_data_end_point > clear_data_start_point else [list_data[clear_data_start_point]]
    clear_data = [info[clear_data_rows] for info in clear_data]

    check_data_rows = 1
    check_data_start_point = 1
    check_data_end_point = 1

    check_data = list_data[check_data_start_point:check_data_end_point] if check_data_end_point > check_data_start_point else [list_data[check_data_start_point]]
    check_data = [info[check_data_rows] for info in check_data]

    print(clear_data, check_data, sep="\n")


def get_check_info(filename, list_name):
    list_data = read_excel_list(filename, list_name)
    clear_data = get_excel_clear_data(list_data)
    check_data = get_excel_check_data(list_data)

    print(clear_data, check_data, sep="\n")
    return clear_data, check_data


def read_excel_list(filename, list_name):
    excel = get_data(filename)
    return excel.get(list_name)


def get_excel_data(list_data, row, start, end):
    data = list_data[start:end] if end > start else [list_data[start]]
    data = [info[row] for info in data]
    return data


def get_excel_clear_data(list_data):
    row = 5
    start = 2
    end = 345
    return get_excel_data(list_data, row=row, start=start, end=end)


def get_excel_check_data(list_data):
    row = 0
    start = 2
    end = 314
    return get_excel_data(list_data, row=row, start=start, end=end)


def save_excel(filename, checked_data):
    data = OrderedDict()
    checked_data = [[checked_data_name] for checked_data_name in checked_data]
    data.update({"Sheet1": checked_data})
    save_data(filename, data)


def find_candidate(raw_data, clear_data, corrector):
    result = corrector.GetCandidates([raw_data], 0)
    for cl_data in clear_data:
        if cl_data in result:
            return cl_data


def check_spelling(check_data, clear_data):
    # format Surname Name
    clear_surnames = [name.split(" ")[0].lower() for name in clear_data]
    clear_names = [name.split(" ")[1].lower() for name in clear_data]

    checked_data = []

    corrector = jamspell.TSpellCorrector()
    corrector.LoadLangModel("my_model.bin")

    for data in check_data:
        name_info = data.lower().split(" ")
        # print(surname, name)
        surname = find_candidate(name_info[0], clear_surnames, corrector)
        # name = find_candidate(name, clear_names, corrector)
        if surname is not None:
            checked_data.append(f"{surname.title()}")

    # print(checked_data)
    return checked_data


def main():
    filename = "Книга2.xlsx"
    list_name = "Лист1 (2)"
    clear_data, check_data = get_check_info(filename=filename, list_name=list_name)
    checked_data = check_spelling(check_data=check_data, clear_data=clear_data)
    save_excel(filename="result_test.xlsx", checked_data=checked_data)

def get_n_gramms(clear_data, n_gramm):
    n_gramms = set()
    for data in clear_data:
        words = data.split(" ")
        for word in words:
            for i in range(len(word)-n_gramm):
                n_gramms.add(word[i:i+n_gramm])
    return n_gramms


def save_clear_txt_data():
    filename = "Книга2.xlsx"
    list_name = "Лист1 (2)"

    clear_data, _ = get_check_info(filename, list_name)
    splitted_data = []
    for cl_data in clear_data:
        cl_data = cl_data.split(" ")
        for cl in cl_data:
            splitted_data.append(cl)
    n_gramms = []
    for i in range(2, 10):
        n_gramms.append("\n".join(get_n_gramms(clear_data, n_gramm=i)))

    clear_data = "\n".join(clear_data) + "\n" + "\n".join(n_gramms) + "\n" + "\n".join(splitted_data)
    # clear_data = "\n".join(clear_data)
    with open("text.txt", "w") as text:
        text.write(clear_data)
    print(clear_data)


def train_data():
    corrector = jamspell.TSpellCorrector()
    # corrector.LoadLangModel("ru_small.bin")
    corrector.TrainLangModel("text.txt", "alphabet_ru.txt", "my_model_2.bin")


if __name__ == "__main__":
    # train_data()
    # save_clear_txt_data()
    jamsprell_test()
    # main()


