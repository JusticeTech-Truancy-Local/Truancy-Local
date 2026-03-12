from PyQt6.QtWidgets import QFileDialog, QMessageBox
from PyQt6.QtGui import QTextCharFormat, QColor
from PyQt6.QtCore import QSettings
from pdf_parser import extract_students_from_pdf
from constructor import Student
import subprocess
import os


def open_pdf(window):
    saved_pdf_dir = window.settings.value("pdf_dir", "/home")

    pdf_path = QFileDialog.getOpenFileName(window, "Open Truancy Report", saved_pdf_dir, "PDF (*.pdf)")[0]
    if not pdf_path:
        return
    
    # Should be saved in registry
    window.settings.setValue("pdf_dir", os.path.dirname(pdf_path))
    window.settings.sync()
    window.pdf_path_bar.setText(pdf_path)

    # Open the PDF with system's default viewer
    subprocess.Popen([pdf_path], shell=True)

    students = extract_students_from_pdf(pdf_path)

    if len(students) == 0:
        print("No students")
    #     QMessageBox.warning(self, "No Data", "No students found in PDF")

    else:
        students[0].printHeaders()
        for s in students:
            s.print()

    window.pdf_opened.emit(students)


