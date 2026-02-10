from PyQt6.QtWidgets import QMainWindow, QWidget, QPushButton, QVBoxLayout, QFileDialog
import xlwings as xw

from pdf_parser import extract_students_from_pdf

class TruancyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("TruancyRecorder")

        # Associated with print_students
        pdf_button = QPushButton("Load PDF")
        pdf_button.clicked.connect(self.print_students)

        # Associated with open excel
        excel_button = QPushButton("Open Excel File")
        excel_button.clicked.connect(self.open_excel)

        center_layout = QVBoxLayout()
        center_layout.addWidget(pdf_button)
        center_layout.addWidget(excel_button)
        
        center_widget = QWidget()
        center_widget.setLayout(center_layout)
        self.setCentralWidget(center_widget)


    def print_students(self):
        pdf_path = QFileDialog.getOpenFileName(self, "Open Truancy Report", "/home", "PDF (*.pdf)")[0]
        students = extract_students_from_pdf(pdf_path)
        if len(students) == 0:
            print("No students")
        else:
            students[0].printHeaders()
        for s in students:
            s.print()

    def open_excel(self):
        excel_path = QFileDialog.getOpenFileName(self, "Open Excel File", "/home", "Excel Files (*.xlsx *.xls)")[0]
        if excel_path:
            try:
                # Open the Excel file with xlwings
                wb = xw.Book(excel_path)
                print(f"Opened Excel file: {excel_path}")
                print(f"Workbook has {len(wb.sheets)} sheet(s)")

            except Exception as e:
                print(f"Error opening Excel file: {e}")