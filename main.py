import sys

from PyQt6.QtWidgets import QApplication

from companion_gui import TruancyWindow

if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = TruancyWindow()
    window.show()

    sys.exit(app.exec())

