from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QMainWindow, QWidget, QPushButton, QGridLayout, QTextEdit, QCheckBox, QSizePolicy, QLabel, \
    QScrollArea, QLineEdit, QMessageBox, QComboBox
from PyQt6.QtCore import pyqtSlot, pyqtSignal, QSettings, Qt
import xlwings as xw

from companion_btns.open_pdf import select_pdf, open_pdf
from companion_btns.open_excel import open_excel
from companion_btns.add_report_to_sheet import add_report_to_sheet
import os


class TruancyWindow(QMainWindow):

    pdf_opened = pyqtSignal(str, list, str)
    excel_opened = pyqtSignal(xw.Book)

    def __init__(self):
        super().__init__()

        self.setWindowTitle("TruancyRecorder")
        self.settings = QSettings("TruancyApp", "TruancyRecorder")
        
        # Store loaded students and workbook
        self.pdf_path = ""
        self.students = []
        self.school_name = ""
        self.workbook = None

        # Associated with select_pdf and open_pdf
        select_pdf_button = QPushButton("Select Report PDF")
        select_pdf_button.clicked.connect(lambda: select_pdf(self))
        select_pdf_button.setIcon(QIcon(os.path.join(os.path.dirname(__file__), "assets/pdf.ico")))
        self.open_pdf_button = QPushButton()
        self.open_pdf_button.clicked.connect(lambda: open_pdf(self))
        self.open_pdf_button.setIcon(QIcon(os.path.join(os.path.dirname(__file__), "assets/open.png")))
        self.open_pdf_button.setFixedWidth(30)
        self.open_pdf_button.setFlat(True)
        self.pdf_opened.connect(self.update_students)
        self.pdf_check = QCheckBox()
        self.pdf_check.setEnabled(False)
        self.pdf_check.setSizePolicy(QSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Preferred))
        self.pdf_path_bar = QLineEdit()
        self.pdf_path_bar.setReadOnly(True)
        self.pdf_path_bar.setPlaceholderText("No PDF loaded")

        # Associated with open excel
        excel_button = QPushButton("Connect to Excel")
        excel_button.clicked.connect(lambda: open_excel(self))
        excel_button.setIcon(QIcon(os.path.join(os.path.dirname(__file__), "assets/excel.ico")))
        self.excel_opened.connect(self.update_workbook)
        self.excel_check = QCheckBox()
        self.excel_check.setEnabled(False)
        self.excel_check.setSizePolicy(QSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Preferred))
        self.excel_path_bar = QLineEdit()
        self.excel_path_bar.setReadOnly(True)
        self.excel_path_bar.setPlaceholderText("No Excel file selected")
        
        # Button to add report to sheet
        self.add_absences_button = QPushButton("Add Report to Sheet")
        self.add_absences_button.clicked.connect(lambda: add_report_to_sheet(self))
        self.sheets_combo = QComboBox()

        # Text box to hold status messages for user
        self.status_box = QTextEdit()
        self.status_box.setReadOnly(True)
        status_scroll = QScrollArea()
        status_scroll.setWidget(self.status_box)
        status_scroll.setWidgetResizable(True)

        center_layout = QGridLayout()
        center_layout.addWidget(self.pdf_check, 0, 0, 1, 1)
        center_layout.addWidget(select_pdf_button, 0, 1, 1, 1)
        center_layout.addWidget(self.pdf_path_bar, 0, 2, 1, 1)
        center_layout.addWidget(self.open_pdf_button, 0, 3, 1, 1)
        center_layout.addWidget(self.excel_check, 1, 0, 1, 1)
        center_layout.addWidget(excel_button, 1, 1, 1, 1)
        center_layout.addWidget(self.excel_path_bar, 1, 2, 1, 2)
        center_layout.addWidget(QLabel("⤷"), 2, 0, 1, 1)
        center_layout.addWidget(self.add_absences_button, 2, 1, 1, 1)
        center_layout.addWidget(self.sheets_combo, 2, 2, 1, 2)
        center_layout.addWidget(status_scroll, 3, 1, 1, 3)
        
        center_widget = QWidget()
        center_widget.setLayout(center_layout)
        self.setCentralWidget(center_widget)

        # Keep window above all other windows
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)

        self.check_files_ready()


    @pyqtSlot(str, list, str)
    def update_students(self, file_path, new_students, school_name):
        self.pdf_path = file_path
        self.students = new_students
        self.school_name = school_name
        self.check_files_ready()


    @pyqtSlot(xw.Book)
    def update_workbook(self, new_workbook):
        self.workbook = new_workbook
        # Update sheets in combo box
        self.sheets_combo.clear()
        if bool(self.workbook):
            self.sheets_combo.addItems(["[Create new]"] + [x.name for x in self.workbook.sheets])
        self.sheets_combo.setEnabled(bool(self.workbook))
        self.check_files_ready()

    def check_files_ready(self):
        has_students = bool(self.students)

        # Check if excel window currently exists; clear if the window has been closed
        if self.workbook and self.workbook.fullname not in [i.fullname for i in xw.books]:
           self.workbook = None

        has_workbook = bool(self.workbook)

        # Update checkboxes to show whether file is loaded
        self.pdf_check.setChecked(has_students)
        self.excel_check.setChecked(has_workbook)
        # Grey out the add report to sheet button unless all data has been loaded
        self.add_absences_button.setEnabled(has_students and has_workbook)
        # Grey out the open pdf in new window button unless a pdf has been selected
        self.open_pdf_button.setEnabled(bool(self.pdf_path))

        return has_students and has_workbook
