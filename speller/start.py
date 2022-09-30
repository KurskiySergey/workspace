import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QHBoxLayout, QVBoxLayout, QFileDialog, QPushButton, QLineEdit, QDialog, QDoubleSpinBox
from PyQt5.QtCore import Qt


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setGeometry(50, 50, 700, 200)
        self.setWindowTitle("Excel names transformer")
        self.left_init_name = QLineEdit()
        self.left_output_name = QLineEdit()
        self.right_check_list_name = QLineEdit()
        self.right_check_list_column = QLineEdit()
        self.right_check_list_row = QLineEdit()
        self.similarity_value = QDoubleSpinBox()

        self.right_real_list_name = QLineEdit()
        self.right_real_list_column = QLineEdit()
        self.right_real_row = QLineEdit()

        self.right_out_list_name = QLineEdit()
        self.right_out_list_column = QLineEdit()
        self.right_out_list_row = QLineEdit()

        start_layout = QVBoxLayout()
        layout = QHBoxLayout()
        self.set_left_side(layout)
        self.set_right_side(layout)

        start_layout.addLayout(layout)
        similarity_value_label = QLabel("Критерий схожести")
        similarity_value_label.setAlignment(Qt.AlignCenter)
        start_layout.addWidget(similarity_value_label)
        self.similarity_value.setRange(0, 1)
        self.similarity_value.setSingleStep(0.01)
        self.similarity_value.setValue(0.5)
        start_layout.addWidget(self.similarity_value)
        start_btn = QPushButton()
        start_btn.setText("Запуск")
        start_btn.clicked.connect(self.launch_speller)
        start_layout.addWidget(start_btn)
        self.setLayout(start_layout)

    def chooseInitFile(self):
        dialog = QFileDialog()
        filename = dialog.getOpenFileName(
            initialFilter="Excel File (*.xlsx, *.xls)",
            options=QFileDialog.DontUseNativeDialog
        )

        self.left_init_name.setText(filename[0])
        self.left_output_name.setText(filename[0])

    def chooseOutputFile(self):
        dialog = QFileDialog()
        filename = dialog.getOpenFileName(
            initialFilter="Excel File (*.xlsx, *.xls)",
            options=QFileDialog.DontUseNativeDialog
        )
        self.left_output_name.setText(filename[0])

    def set_left_side(self, start_layout):
        left_layout = QVBoxLayout()

        init_layout = QHBoxLayout()
        output_layout = QHBoxLayout()
        left_init_label = QLabel("Исходный файл")
        left_output_label = QLabel("Выходной файл")

        init_push_button = QPushButton("Выбрать")
        output_push_button = QPushButton("Выбрать")

        left_layout.addWidget(left_init_label)
        left_layout.addLayout(init_layout)
        left_layout.addWidget(left_output_label)
        left_layout.addLayout(output_layout)

        init_layout.addWidget(self.left_init_name)
        init_layout.addWidget(init_push_button)
        output_layout.addWidget(self.left_output_name)
        output_layout.addWidget(output_push_button)

        init_push_button.clicked.connect(self.chooseInitFile)
        output_push_button.clicked.connect(self.chooseOutputFile)

        start_layout.addLayout(left_layout)

    def set_right_side(self, start_layout):
        right_layout = QVBoxLayout()
        check_info_layout = QHBoxLayout()
        check_label = QLabel("Местоположение списка правильных имен")
        right_layout.addWidget(check_label)
        self.right_check_list_name.setPlaceholderText("List1")
        self.right_check_list_column.setPlaceholderText("A")
        self.right_check_list_row.setPlaceholderText("start:end")
        check_info_layout.addWidget(self.right_check_list_name)
        check_info_layout.addWidget(self.right_check_list_column)
        check_info_layout.addWidget(self.right_check_list_row)
        right_layout.addLayout(check_info_layout)

        real_info_layout = QHBoxLayout()
        real_label = QLabel("Местоположение списка проверяемых имен")
        right_layout.addWidget(real_label)
        self.right_real_list_name.setPlaceholderText("List2")
        self.right_real_list_column.setPlaceholderText("B")
        self.right_real_row.setPlaceholderText("start:end")
        real_info_layout.addWidget(self.right_real_list_name)
        real_info_layout.addWidget(self.right_real_list_column)
        real_info_layout.addWidget(self.right_real_row)
        right_layout.addLayout(real_info_layout)

        out_info_layout = QHBoxLayout()
        output_label = QLabel("Местоположение результата")
        right_layout.addWidget(output_label)
        self.right_out_list_name.setPlaceholderText("List3")
        self.right_out_list_column.setPlaceholderText("C")
        self.right_out_list_row.setPlaceholderText("start:end")
        out_info_layout.addWidget(self.right_out_list_name)
        out_info_layout.addWidget(self.right_out_list_column)
        out_info_layout.addWidget(self.right_out_list_row)
        right_layout.addLayout(out_info_layout)

        start_layout.addLayout(right_layout)

    def set_config(self):
        print("setting config")
        from config import set_qt_params
        check_rows = self.right_check_list_row.text().split(":")
        real_rows = self.right_real_row.text().split(":")
        out_rows = self.right_out_list_row.text().split(":")
        if self.left_init_name.text() != "":
            params = {
                "EXCEL_RESULT_FILENAME": self.left_output_name.text().encode("utf-8"),
                "EXCEL_ORIGIN_FILENAME": self.left_init_name.text().encode("utf-8"),
                "OUTPUT_SIMILARITY": float(self.similarity_value.value()),
                "CHECK_DATA_INFO": (self.right_check_list_column.text().encode("utf-8").upper(), self.right_check_list_name.text().encode("utf-8"), int(check_rows[0].rstrip()), int(check_rows[1].rstrip())),
                "REAL_DATA_INFO": (self.right_real_list_column.text().encode("utf-8").upper(), self.right_real_list_name.text().encode("utf-8"), int(real_rows[0].rstrip()), int(real_rows[1].rstrip())),
                "OUTPUT_DATA_INFO": (self.right_out_list_column.text().encode("utf-8").upper(), self.right_out_list_name.text().encode("utf-8"), int(out_rows[0].rstrip()), int(out_rows[1].rstrip()))
            }
            set_qt_params(params=params)
        print("done")

    def launch_speller(self):
        self.set_config()
        print("initializing object...")
        from main import Main
        speller = Main()
        print("done")
        print("starting...")
        speller.start()
        print("finished")
        result = QDialog(parent=self)
        result.setWindowTitle("Excel names transformer")
        result_btn = QPushButton("ОК")
        result_btn.clicked.connect(result.close)
        layout = QVBoxLayout()
        layout.addWidget(result_btn)
        result.setLayout(layout)
        result.show()


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
