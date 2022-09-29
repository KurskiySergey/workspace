import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QHBoxLayout
from PyQt5.QtCore import Qt


def main():
    app = QApplication(sys.argv)
    widget = QWidget()
    widget.setGeometry(50, 50, 320, 200)
    widget.setWindowTitle("Excel names transformer")

    layout = QHBoxLayout()
    left_label = QLabel("left")
    layout.addWidget(left_label,  alignment=Qt.AlignCenter)

    right_label = QLabel("right")
    layout.addWidget(right_label,  alignment=Qt.AlignCenter)

    widget.setLayout(layout)
    widget.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
