from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QMainWindow, QWidget, QPushButton, QGridLayout, QTextEdit, QCheckBox, QSizePolicy, QLabel, \
    QScrollArea, QLineEdit, QMessageBox, QComboBox, QVBoxLayout, QGroupBox, QHBoxLayout
from PyQt6.QtCore import pyqtSlot, pyqtSignal, QSettings, Qt
import xlwings as xw
from docxtpl import DocxTemplate

from companion_btns.open_pdf import select_pdf, open_pdf
from companion_btns.open_excel import open_excel
from companion_btns.open_docx import open_docx
from companion_btns.add_report_to_sheet import add_report_to_sheet
from difflib import SequenceMatcher
import os

from companion_btns.status_box import StatusBox


class TruancyWindow(QMainWindow):

    pdf_opened = pyqtSignal(str, list, str)
    excel_opened = pyqtSignal(xw.Book)
    docx_opened = pyqtSignal(str, object)  # file path, DocxTemplate instance

    def __init__(self):
        super().__init__()

        self.setWindowTitle("TruancyRecorder")
        self.settings = QSettings("TruancyApp", "TruancyRecorder")
        
        # Store loaded students and workbook
        self.pdf_path = ""
        self.students = []
        self.school_name = ""
        self.workbook = None

        # Store loaded Word doc
        self.docx_path = ""
        self.docx_template = None

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
        self.pdf_path_bar = QLineEdit()
        self.pdf_path_bar.setReadOnly(True)
        self.pdf_path_bar.setPlaceholderText("No PDF loaded")

        # Associated with open excel
        excel_button = QPushButton("Connect to Excel")
        excel_button.clicked.connect(lambda: open_excel(self))
        excel_button.setIcon(QIcon(os.path.join(os.path.dirname(__file__), "assets/excel.ico")))
        self.excel_opened.connect(self.update_workbook)
        self.excel_path_bar = QLineEdit()
        self.excel_path_bar.setReadOnly(True)
        self.excel_path_bar.setPlaceholderText("No Excel file selected")

        # Associated with select_docx (Word document template)
        select_docx_button = QPushButton("Select Word Doc")
        select_docx_button.clicked.connect(lambda: open_docx(self))
        select_docx_button.setIcon(QIcon(os.path.join(os.path.dirname(__file__), "assets/word.ico")))
        self.docx_opened.connect(self.update_docx)
        self.docx_path_bar = QLineEdit()
        self.docx_path_bar.setReadOnly(True)
        self.docx_path_bar.setPlaceholderText("No Word document loaded")
        
        # Button to add report to sheet
        self.add_absences_button = QPushButton("Add Report to Sheet")
        self.add_absences_button.clicked.connect(lambda: add_report_to_sheet(self))
        self.sheets_combo = QComboBox()

        # Text box to hold status messages for user
        self.status_box = StatusBox()
        self.status_box.go_to_cell.connect(self.go_to_cell)
        status_scroll = QScrollArea()
        status_scroll.setWidget(self.status_box)
        status_scroll.setWidgetResizable(True)

        center_layout = QVBoxLayout()

        def contain_widgets(group_name, widgets):
            # Encapsulates widgets in a named box
            container = QGroupBox(group_name)
            hlayout = QHBoxLayout()
            for w in widgets:
                hlayout.addWidget(w)
            container.setLayout(hlayout)
            return container

        self.step_containers = [
            contain_widgets("1. ☐", [excel_button, self.excel_path_bar]),
            contain_widgets("2. ☐", [select_pdf_button, self.pdf_path_bar, self.open_pdf_button]),
            contain_widgets("3. ☐", [select_docx_button, self.docx_path_bar]),
            contain_widgets("4. ☐", [self.add_absences_button, self.sheets_combo]),
        ]
        for sc in self.step_containers:
            center_layout.addWidget(sc)
        center_layout.addWidget(status_scroll)
        
        center_widget = QWidget()
        center_widget.setLayout(center_layout)
        self.setCentralWidget(center_widget)

        # Keep window above all other windows
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)

        self.check_files_ready()

    @pyqtSlot(str, str)
    def go_to_cell(self, sheet, address):
        assert(self.workbook is not None)
        self.workbook.activate(steal_focus=True)
        self.workbook.sheets[sheet].select()
        self.workbook.sheets[sheet].range(address).select()

    @pyqtSlot(str, list, str)
    def update_students(self, file_path, new_students, school_name):
        self.pdf_path = file_path
        self.students = new_students
        self.school_name = school_name
        self.check_files_ready(did_update=True)

    @pyqtSlot(xw.Book)
    def update_workbook(self, new_workbook):
        self.workbook = new_workbook
        # Update sheets in combo box
        self.sheets_combo.clear()
        if bool(self.workbook):
            self.sheets_combo.addItems(["[Create new]"] + [x.name for x in self.workbook.sheets])
        self.sheets_combo.setEnabled(bool(self.workbook))
        self.check_files_ready(did_update=True)

    @pyqtSlot(str, object)
    def update_docx(self, file_path, template):
        self.docx_path = file_path
        self.docx_template = template
        self.check_files_ready(did_update=True)

    def check_files_ready(self, did_update=False):
        has_students = bool(self.students)
        has_docx = bool(self.docx_path)

        # Check if excel window currently exists; clear if the window has been closed
        if self.workbook and self.workbook.fullname not in [i.fullname for i in xw.books]:
           self.workbook = None

        has_workbook = bool(self.workbook)

        # Grey out the add report to sheet button unless all data has been loaded
        self.add_absences_button.setEnabled(has_students and has_workbook)
        # Grey out the open pdf in new window button unless a pdf has been selected
        self.open_pdf_button.setEnabled(bool(self.pdf_path))

        # Set checkboxes for each step
        self.step_containers[0].setTitle("1. " + ("☑" if has_workbook else "☐"))
        self.step_containers[1].setTitle("2. " + ("☑" if has_students else "☐"))
        self.step_containers[2].setTitle("3. " + ("☑" if has_docx else "☐"))

        # Set dropdown to sheet that best matches school name
        if did_update:
            self.step_containers[3].setTitle("4. ☐")
            if has_workbook and has_students:
                best_sheet = self.best_match(self.school_name, [x.name for x in self.workbook.sheets])
                self.sheets_combo.setCurrentIndex(best_sheet + 1)

        return has_students and has_workbook

    def best_match(self, name, options):
        ## Returns best matching string in options, -1 if none match well
        ratios = [SequenceMatcher(None, name.lower(), opt.lower()).ratio() for opt in options]
        maxr = max(ratios)
        if maxr > 0.5:
            return ratios.index(maxr)
        return -1