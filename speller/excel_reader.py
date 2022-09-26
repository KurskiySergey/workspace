from pyexcel_xlsx import get_data, save_data
from collections import OrderedDict
from config import OUTPUT_DATA_INFO, REAL_DATA_INFO, CHECK_DATA_INFO


class ExcelWorker:
    def __init__(self, filename):
        self.filename = filename
        self.excel_en = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        self.excel_en_dict = {el: i + 1 for i, el in enumerate(self.excel_en)}
        self.data = None
        self.data_list = None

    def read_data(self):
        self.data = get_data(self.filename)

    def read_list(self, list_name):
        if self.data is not None:
            self.data_list = self.data.get(list_name)

    def get_data(self, column, row_start, row_end):
        rows = self.data_list[row_start:row_end] if row_end > row_start else [self.data_list[row_start]]
        data = [row[column] for row in rows]
        return data

    def convert_data_info(self):
        check_info = (self.__convert_to_number(CHECK_DATA_INFO[0]), CHECK_DATA_INFO[1],
                      CHECK_DATA_INFO[2], CHECK_DATA_INFO[3])
        real_info = (self.__convert_to_number(REAL_DATA_INFO[0]), REAL_DATA_INFO[1],
                     REAL_DATA_INFO[2], REAL_DATA_INFO[3])

        return check_info, real_info

    def __convert_to_number(self, column):
        column = sum([self.excel_en_dict[el]*pow(len(self.excel_en), len(column)-i-1) for i, el in enumerate(column)])
        return column

    def save_excel_data(self, data: dict, filename):
        excel_data = OrderedDict()
        column = self.__convert_to_number(column=OUTPUT_DATA_INFO[0])
        list_name = OUTPUT_DATA_INFO[1]
        row_start = OUTPUT_DATA_INFO[2]
        row_end = OUTPUT_DATA_INFO[3]
        if self.filename == filename:
            self.read_list(list_name=list_name)
            if self.data_list is None:
                self.data_list = [_ for _ in range(row_end)]
        items = list(data.items())
        counter = 0
        print(len(items))
        if len(self.data_list) != 0:
            for i, row in enumerate(self.data_list):
                if row_start <= i + 1 <= row_end:
                    if len(row) < column:
                        for _ in range(column - len(row)):
                            row.append()
                    row[column-1] = " ".join(items[counter][-1])
                    counter += 1
                    self.data_list[i] = row
            if counter + row_start < row_end:
                for i in range(row_start + counter, row_end+1):
                    row = [_ for _ in range(column)]
                    row[column-1] = " ".join(items[counter][-1])
                    counter += 1
                    self.data_list.append(row)
        excel_data.update({list_name: self.data_list})
        save_data(filename, excel_data)
