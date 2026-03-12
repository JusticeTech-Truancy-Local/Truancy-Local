from PyQt6.QtWidgets import QMainWindow, QWidget, QPushButton, QVBoxLayout, QHBoxLayout, QTextEdit, QMessageBox, QDialog, QLabel
from PyQt6.QtCore import Qt
import os
import webbrowser
from companion_btns.print_students import print_students
from companion_btns.open_excel import open_excel
from companion_btns.add_total_absences import add_total_absences

class TruancyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Show Terms of Service popup at startup
        if not self.show_terms_of_service():
            # User declined, exit application
            import sys
            sys.exit()
        
        # User accepted, continue with app
        self.setWindowTitle("TruancyRecorder")
        
        # Store loaded students and workbook
        self.students = []
        self.workbook = None
        
        # PDF loading button
        pdf_button = QPushButton("Load PDF")
        pdf_button.clicked.connect(lambda: print_students(self))
        
        # Excel file button
        excel_button = QPushButton("Open Excel File")
        excel_button.clicked.connect(lambda: open_excel(self))
        
        # Add absences button
        add_absences_button = QPushButton("Add Total Absences to Excel")
        add_absences_button.clicked.connect(lambda: add_total_absences(self))
        
        # Help button
        help_button = QPushButton("?")
        help_button.setMaximumWidth(30)
        help_button.clicked.connect(self.show_instructions)
        help_button.setToolTip("Click for instructions")
        
        # Create ? layout top right corner
        top_layout = QHBoxLayout()
        top_layout.addStretch()
        top_layout.addWidget(help_button)
        
        # Text box to hold status messages for user
        self.status_box = QTextEdit()
        self.status_box.setReadOnly(True)
        
        # Main layout
        center_layout = QVBoxLayout()
        center_layout.addLayout(top_layout)
        center_layout.addWidget(pdf_button)
        center_layout.addWidget(excel_button)
        center_layout.addWidget(add_absences_button)
        center_layout.addWidget(self.status_box)
        
        center_widget = QWidget()
        center_widget.setLayout(center_layout)
        self.setCentralWidget(center_widget)
    
    def open_terms_file(self):
        """Opens the Terms of Service file"""
        terms_file = "TermsOfService.txt"
        
        if not os.path.exists(terms_file):
            QMessageBox.warning(self, "File Not Found", f"Could not find {terms_file}")
            return
        try:
            webbrowser.open(os.path.abspath(terms_file))
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not open file: {e}")
    
    def show_terms_of_service(self):
        """Shows terms popup with View Full Terms button, returns True if user accepts"""
        # Create custom dialog
        dialog = QDialog(self)
        dialog.setWindowTitle("Terms of Service")
        dialog.setModal(True)
        
        # Main text
        terms_summary = QLabel(
            "TRUANCY RECORDER - TERMS OF SERVICE\n\n"
            "This application processes student absence data for educational purposes only.\n\n"
            "By using this application, you agree to our full Terms of Service.\n\n"
            "Do you accept these terms?"
        )
        terms_summary.setWordWrap(True)
        terms_summary.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # View Full Terms button
        view_terms_button = QPushButton("View Full Terms")
        view_terms_button.clicked.connect(self.open_terms_file)
        
        # Accept button
        accept_button = QPushButton("I Accept")
        accept_button.clicked.connect(dialog.accept)
        
        # Decline button
        decline_button = QPushButton("I Don't Accept")
        decline_button.clicked.connect(dialog.reject)
        
        # Button layout
        button_layout = QHBoxLayout()
        button_layout.addWidget(accept_button)
        button_layout.addWidget(decline_button)
        
        # Main layout
        layout = QVBoxLayout()
        layout.addWidget(terms_summary)
        layout.addWidget(view_terms_button)
        layout.addSpacing(20)
        layout.addLayout(button_layout)
        
        dialog.setLayout(layout)
        
        # Show dialog and return result
        result = dialog.exec()
        return result == QDialog.DialogCode.Accepted
    
    def show_instructions(self):
        """Shows help popup with usage instructions"""
        instructions_text = """HOW TO USE TRUANCY RECORDER

Step 1: Load PDF
   Click "Load PDF" and select the truancy report PDF file.

Step 2: Open Excel File  
   Click "Open Excel File" and select your tracking spreadsheet.

Step 3: Add Data to Excel
   Click "Add Total Absences to Excel" and enter the date when prompted.
   Students will be matched by ID/name and data will be added.

Red highlighting = 40+ unexcused hours - requires court intervention

Remember to save your Excel file after adding data!"""
        
        QMessageBox.information(self, "Instructions", instructions_text)
