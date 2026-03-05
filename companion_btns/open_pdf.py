from PyQt6.QtWidgets import QFileDialog, QMessageBox
from PyQt6.QtGui import QTextCharFormat, QColor
from pdf_parser import extract_students_from_pdf
from constructor import Student
import subprocess


def open_pdf(self):

    pdf_path = QFileDialog.getOpenFileName(self, "Open Truancy Report", "/home", "PDF (*.pdf)")[0]
    if not pdf_path:
        return

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

        # Display loaded students and hours in the status box
        cursor = self.status_box.textCursor()
        format = QTextCharFormat()
        for s in students:
            # Highlight students over unexcused threshold in red
            if float(s.unexcused) >= Student.redThreshold:
                format.setBackground(QColor(255, 0, 0, 80))
            cursor.insertText(f"{s.firstName} {s.lastName} - {s.unexcused} hrs\n", format)
            format.clearBackground()
        cursor.setCharFormat(format)

    self.pdf_opened.emit(students)


