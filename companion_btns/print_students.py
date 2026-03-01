from PyQt6.QtWidgets import QFileDialog, QMessageBox
from PyQt6.QtGui import QTextCharFormat, QColor
from pdf_parser import extract_students_from_pdf
from constructor import Student

def print_students(window):
    pdf_path = QFileDialog.getOpenFileName(window, "Open Truancy Report", "/home", "PDF (*.pdf)")[0]
    if not pdf_path:
        return
        
    window.students = extract_students_from_pdf(pdf_path)
    if len(window.students) == 0:
        print("No students")
        QMessageBox.warning(window, "No Data", "No students found in PDF")
        window.status_box.append("No students found in PDF")
    else:
        window.students[0].printHeaders()
        for s in window.students:
            s.print()

        # Display loaded students and hours in the status box
        cursor = window.status_box.textCursor()
        format = QTextCharFormat()
        for s in window.students:
            # Highlight students over unexcused threshold in red
            if s.absenceTotal and float(s.absenceTotal) >= 40:
                format.setBackground(QColor(255, 0, 0, 80))
            cursor.insertText(f"{s.firstName} {s.lastName} - {s.absenceTotal} hrs\n", format)
            format.clearBackground()

        QMessageBox.information(window, "Success", f"Loaded {len(window.students)} students from PDF")
        window.status_box.append(f"Loaded {len(window.students)} students from PDF")
