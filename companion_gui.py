from PyQt6.QtWidgets import QMainWindow, QWidget, QPushButton, QGridLayout, QTextEdit, QCheckBox, QSizePolicy, QLabel, \
    QScrollArea
from PyQt6.QtCore import pyqtSlot, pyqtSignal
import xlwings as xw

from companion_btns.open_pdf import open_pdf
from companion_btns.open_excel import open_excel
from companion_btns.add_report_to_sheet import add_report_to_sheet


class TruancyWindow(QMainWindow):

    pdf_opened = pyqtSignal(list)
    excel_opened = pyqtSignal(xw.Book)

    def __init__(self):
        super().__init__()

        self.setWindowTitle("TruancyRecorder")
        
        # Store loaded students and workbook
        self.students = []
        self.workbook = None

        # Associated with open_pdf
        pdf_button = QPushButton("Open Report PDF")
        pdf_button.clicked.connect(lambda: open_pdf(self))
        self.pdf_opened.connect(self.update_students)
        self.pdf_check = QCheckBox()
        self.pdf_check.setEnabled(False)
        self.pdf_check.setSizePolicy(QSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Preferred))

        # Associated with open excel
        excel_button = QPushButton("Open Excel Sheet")
        excel_button.clicked.connect(lambda: open_excel(self))
        self.excel_opened.connect(self.update_workbook)
        self.excel_check = QCheckBox()
        self.excel_check.setEnabled(False)
        self.excel_check.setSizePolicy(QSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Preferred))
        
        # Button to add report to sheet
        self.add_absences_button = QPushButton("Add Report to Sheet")
        self.add_absences_button.clicked.connect(lambda: add_report_to_sheet(self))

        # Text box to hold status messages for user
        self.status_box = QTextEdit()
        self.status_box.setReadOnly(True)
        status_scroll = QScrollArea()
        status_scroll.setWidget(self.status_box)
        status_scroll.setWidgetResizable(True)

        center_layout = QGridLayout()
        center_layout.addWidget(pdf_button, 0, 1, 1, 1)
        center_layout.addWidget(self.pdf_check, 0, 0, 1, 1)
        center_layout.addWidget(excel_button, 1, 1, 1, 1)
        center_layout.addWidget(self.excel_check, 1, 0, 1, 1)
        center_layout.addWidget(self.add_absences_button, 2, 1, 1, 1)
        center_layout.addWidget(QLabel("⤷"), 2, 0, 1, 1)
        center_layout.addWidget(status_scroll, 3, 0, 1, 2)
        
        center_widget = QWidget()
        center_widget.setLayout(center_layout)
        self.setCentralWidget(center_widget)

        self.check_files_ready()


    @pyqtSlot(list)
    def update_students(self, new_students):
        self.students = new_students
        self.check_files_ready()


    @pyqtSlot(xw.Book)
    def update_workbook(self, new_workbook):
        self.workbook = new_workbook
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

        return has_students and has_workbook
