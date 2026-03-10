from PyQt6.QtWidgets import QMainWindow, QWidget, QPushButton, QVBoxLayout, QTextEdit

from companion_btns.print_students import print_students
from companion_btns.open_excel import open_excel
from companion_btns.add_total_absences import add_total_absences

class TruancyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("TruancyRecorder")
        
        # Store loaded students and workbook
        self.students = []
        self.workbook = None

        # Associated with print_students
        pdf_button = QPushButton("Load PDF")
        pdf_button.clicked.connect(lambda: print_students(self))

        # Associated with open excel
        excel_button = QPushButton("Open Excel File")
        excel_button.clicked.connect(lambda: open_excel(self))
        
        # Button to add total absences to Excel
        add_absences_button = QPushButton("Add Total Absences to Excel")
        add_absences_button.clicked.connect(lambda: add_total_absences(self))

        # Text box to hold status messages for user
        self.status_box = QTextEdit()
        self.status_box.setReadOnly(True)

        center_layout = QVBoxLayout()
        center_layout.addWidget(pdf_button)
        center_layout.addWidget(excel_button)
        center_layout.addWidget(add_absences_button)
        center_layout.addWidget(self.status_box)
        
        center_widget = QWidget()
        center_widget.setLayout(center_layout)
        self.setCentralWidget(center_widget)