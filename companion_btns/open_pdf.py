from PyQt6.QtWidgets import QFileDialog
from pdf_parser import extract_students_from_pdf
# import subprocess


def open_pdf(window):

    pdf_path = QFileDialog.getOpenFileName(window, "Open Truancy Report", "/home", "PDF (*.pdf)")[0]
    if not pdf_path:
        return

    # # Open the PDF with system's default viewer
    # subprocess.Popen([pdf_path], shell=True)

    students = extract_students_from_pdf(pdf_path)

    if len(students) == 0:
        print("No students")
    #     QMessageBox.warning(self, "No Data", "No students found in PDF")

    else:
        students[0].printHeaders()
        for s in students:
            s.print()

    window.pdf_opened.emit(students)


