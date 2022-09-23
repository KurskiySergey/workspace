from pyexcel_xlsx import get_data, save_data
from collections import OrderedDict


class ExcelReader:
    def __init__(self, filename):
        self.filename = filename
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

    @staticmethod
    def save_excel_data(data, filename):
        excel_data = OrderedDict()
        data = [[dt] for dt in data]
        excel_data.update({"Sheet1": data})
        save_data(filename, excel_data)
