from PyQt6.QtWidgets import QMainWindow, QWidget, QPushButton, QVBoxLayout, QFileDialog, QMessageBox
import xlwings as xw

from pdf_parser import extract_students_from_pdf

class TruancyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("TruancyRecorder")
        
        # Store loaded students and workbook
        self.students = []
        self.workbook = None

        # Associated with print_students
        pdf_button = QPushButton("Load PDF")
        pdf_button.clicked.connect(self.print_students)

        # Associated with open excel
        excel_button = QPushButton("Open Excel File")
        excel_button.clicked.connect(self.open_excel)
        
        # Button to add total absences to Excel
        add_absences_button = QPushButton("Add Total Absences to Excel")
        add_absences_button.clicked.connect(self.add_total_absences)

        center_layout = QVBoxLayout()
        center_layout.addWidget(pdf_button)
        center_layout.addWidget(excel_button)
        center_layout.addWidget(add_absences_button)
        
        center_widget = QWidget()
        center_widget.setLayout(center_layout)
        self.setCentralWidget(center_widget)


    def print_students(self):
        pdf_path = QFileDialog.getOpenFileName(self, "Open Truancy Report", "/home", "PDF (*.pdf)")[0]
        if not pdf_path:
            return
            
        self.students = extract_students_from_pdf(pdf_path)
        if len(self.students) == 0:
            print("No students")
            QMessageBox.warning(self, "No Data", "No students found in PDF")
        else:
            self.students[0].printHeaders()
            for s in self.students:
                s.print()
            QMessageBox.information(self, "Success", f"Loaded {len(self.students)} students from PDF")

    def open_excel(self):
        excel_path = QFileDialog.getOpenFileName(self, "Open Excel File", "/home", "Excel Files (*.xlsx *.xls)")[0]
        if excel_path:
            try:
                # Open the Excel file with xlwings
                self.workbook = xw.Book(excel_path)
                print(f"Opened Excel file: {excel_path}")
                print(f"Workbook has {len(self.workbook.sheets)} sheet(s)")
                QMessageBox.information(self, "Success", f"Opened Excel file with {len(self.workbook.sheets)} sheet(s)")

            except Exception as e:
                print(f"Error opening Excel file: {e}")
                QMessageBox.critical(self, "Error", f"Error opening Excel file: {e}")

    def add_total_absences(self):
        """Add Total Absences column from PDF data to Excel file"""
        
        # Check if we have both PDF data and Excel file
        if not self.students:
            QMessageBox.warning(self, "No PDF Data", "Please load a PDF file first")
            return
        
        if not self.workbook:
            QMessageBox.warning(self, "No Excel File", "Please open an Excel file first")
            return
        
        try:
            # Get the active sheet
            sheet = self.workbook.sheets.active
            
            # Insert column at position 22 (between column 21 and current 22)
            insert_col = 22
            
            # Insert a new column
            sheet.range(f'{self._col_letter(insert_col)}:{self._col_letter(insert_col)}').api.Insert()
            
            # Add header
            sheet.range(f'{self._col_letter(insert_col)}1').value = "Total Absences (from PDF)"
            
            print(f"\n=== ADDING TOTAL ABSENCES ===")
            
            # Add Total Absences from PDF starting at row 2
            added_count = 0
            for i, student in enumerate(self.students):
                row = i + 2  # Start at row 2 
                
                if student.absenceTotal:
                    try:
                        total_abs = float(student.absenceTotal)
                        sheet.range(f'{self._col_letter(insert_col)}{row}').value = total_abs
                        print(f"Row {row}: Added {total_abs} for {student.firstName} {student.lastName} (ID: {student.id})")
                        added_count += 1
                    except (ValueError, TypeError):
                        print(f"Warning: Could not convert absenceTotal for student {student.id}: {student.absenceTotal}")
            
            print(f"\n=== SUMMARY ===")
            print(f"Total absences added: {added_count}")
            
            QMessageBox.information(self, "Success", 
                f"Added Total Absences for {added_count} students\n"
                f"Values added to column {self._col_letter(insert_col)} starting at row 2")
            
        except Exception as e:
            import traceback
            print(f"Error adding total absences: {e}")
            print(traceback.format_exc())
            QMessageBox.critical(self, "Error", f"Error adding total absences: {e}")
    
    def _col_letter(self, col_num):
        """Convert column number to Excel column letter (1=A, 2=B, )"""
        string = ""
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            string = chr(65 + remainder) + string
        return string
