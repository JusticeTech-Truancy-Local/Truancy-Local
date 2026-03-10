from PyQt6.QtWidgets import QMainWindow, QWidget, QPushButton, QVBoxLayout, QTextEdit, QHBoxLayout, QLineEdit, QLabel
from PyQt6.QtCore import QSettings

from companion_btns.print_students import print_students
from companion_btns.open_excel import open_excel
from companion_btns.add_total_absences import add_total_absences

class TruancyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.settings = QSettings("TruancyApp", "TruancyRecorder")
        self.setWindowTitle("TruancyRecorder")
        
        # Store loaded students and workbook
        self.students = []
        self.workbook = None

        # Associated with print_students
        pdf_button = QPushButton("Load PDF")
        pdf_button.clicked.connect(lambda: print_students(self))

        # Bar displaying the path
        self.pdf_path_bar = QLineEdit()
        self.pdf_path_bar.setReadOnly(True)
        self.pdf_path_bar.setPlaceholderText("No PDF loaded")

        # Horizontal Ordering
        pdf_row = QHBoxLayout()
        pdf_row.addWidget(self.pdf_path_bar)
        pdf_row.addWidget(pdf_button)

        # Associated with open excel
        excel_button = QPushButton("Open Excel File")
        excel_button.clicked.connect(lambda: open_excel(self))

        # Excel Bar displaying paths
        self.excel_path_bar = QLineEdit()
        self.excel_path_bar.setPlaceholderText("No Excel file selected")

        # Horizontal Ordering
        excel_row = QHBoxLayout()
        excel_row.addWidget(self.excel_path_bar)
        excel_row.addWidget(excel_button)
        
        # Button to add total absences to Excel
        add_absences_button = QPushButton("Add Total Absences to Excel")
        add_absences_button.clicked.connect(lambda: add_total_absences(self))

        # Text box to hold status messages for user
        self.status_box = QTextEdit()
        self.status_box.setReadOnly(True)

        center_layout = QVBoxLayout()
        center_layout.addLayout(pdf_row)
        center_layout.addLayout(excel_row)
        center_layout.addWidget(add_absences_button)
        center_layout.addWidget(self.status_box)
        
        center_widget = QWidget()
        center_widget.setLayout(center_layout)
        self.setCentralWidget(center_widget)

        # Restores the saved paths for the labels
        saved_excel = self.settings.value("excel_path", "")
        if saved_excel:
            self.excel_path_bar.setText(saved_excel)