from PyQt6.QtWidgets import QFileDialog, QMessageBox
from PyQt6.QtGui import QTextCharFormat, QColor
from pdf_parser import extract_students_from_pdf
from constructor import Student
import os

def print_students(self):
    saved_pdf_dir = self.settings.value("pdf_dir", "/home")

    pdf_path = QFileDialog.getOpenFileName(self, "Open Truancy Report", saved_pdf_dir, "PDF (*.pdf)")[0]
    if not pdf_path:
        return
    
    # Should be saved in registry
    self.settings.setValue("pdf_dir", os.path.dirname(pdf_path))
    self.settings.sync()
    self.pdf_path_bar.setText(pdf_path)
        
    self.students = extract_students_from_pdf(pdf_path)
    if len(self.students) == 0:
        print("No students")
        QMessageBox.warning(self, "No Data", "No students found in PDF")
    else:
        self.students[0].printHeaders()
        for s in self.students:
            s.print()

        # Display loaded students and hours in the status box
        cursor = self.status_box.textCursor()
        format = QTextCharFormat()
        for s in self.students:
            # Highlight students over unexcused threshold in red
            if float(s.unexcused) >= Student.redThreshold:
                format.setBackground(QColor(255, 0, 0, 80))
            cursor.insertText(f"{s.firstName} {s.lastName} - {s.unexcused} hrs\n", format)
            format.clearBackground()

        QMessageBox.information(self, "Success", f"Loaded {len(self.students)} students from PDF")
