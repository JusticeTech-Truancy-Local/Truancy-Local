from PyQt6.QtWidgets import QMainWindow, QWidget, QPushButton, QVBoxLayout, QFileDialog

from pdf_parser import extract_students_from_pdf

class TruancyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("TruancyRecorder")

        pdf_button = QPushButton("Load PDF")
        pdf_button.clicked.connect(self.print_students)

        center_layout = QVBoxLayout()
        center_layout.addWidget(pdf_button)
        
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
