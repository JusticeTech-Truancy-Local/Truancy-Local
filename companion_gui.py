from PyQt6.QtWidgets import QMainWindow, QWidget, QPushButton, QVBoxLayout, QTextEdit, QHBoxLayout, QLineEdit, QLabel, QMessageBox
from PyQt6.QtCore import QSettings

from companion_btns.print_students import print_students
from companion_btns.open_excel import open_excel
from companion_btns.add_total_absences import add_total_absences

class TruancyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.settings = QSettings("TruancyApp", "TruancyRecorder")
        
        # Show Terms of Service to user first
        if not self.show_terms_of_service():
            # User declined - exit the application
            import sys
            sys.exit()
        
        # User accepted, continue with app
        self.setWindowTitle("TruancyRecorder")
        
        # Store loaded students and workbook
        self.students = []
        self.workbook = None
        
        # Associated with print_students
        pdf_button = QPushButton("Load PDF")
        pdf_button.clicked.connect(lambda: print_students(self))

        # Bar displaying the path
        self.pdf_path_bar = QLineEdit()
        self.pdf_path_bar.setReadOnly(True)
        self.pdf_path_bar.setPlaceholderText("No PDF loaded")

        # Horizontal Ordering
        pdf_row = QHBoxLayout()
        pdf_row.addWidget(self.pdf_path_bar)
        pdf_row.addWidget(pdf_button)

        
        # Associated with open excel
        excel_button = QPushButton("Open Excel File")
        excel_button.clicked.connect(lambda: open_excel(self))

        # Excel Bar displaying paths
        self.excel_path_bar = QLineEdit()
        self.excel_path_bar.setPlaceholderText("No Excel file selected")

        # Horizontal Ordering
        excel_row = QHBoxLayout()
        excel_row.addWidget(self.excel_path_bar)
        excel_row.addWidget(excel_button)
        
        # Button to add total absences to Excel
        add_absences_button = QPushButton("Add Total Absences to Excel")
        add_absences_button.clicked.connect(lambda: add_total_absences(self))
        
        # Help button: Question mark ? 
        help_button = QPushButton("?")
        help_button.setMaximumWidth(30)
        help_button.clicked.connect(self.show_instructions)
        help_button.setToolTip("Click for instructions")
        
        # Create ? layout top right corner
        top_layout = QHBoxLayout()
        top_layout.addStretch()  # Push help button to the right
        top_layout.addWidget(help_button)
        
        # Text box to hold status messages for user
        self.status_box = QTextEdit()
        self.status_box.setReadOnly(True)
        
        # Main layout
        center_layout = QVBoxLayout()
        center_layout.addLayout(pdf_row)
        center_layout.addLayout(excel_row)
        center_layout.addLayout(top_layout)  # Add help button at top
        center_layout.addWidget(pdf_button)
        center_layout.addWidget(excel_button)
        center_layout.addWidget(add_absences_button)
        center_layout.addWidget(self.status_box)
        
        center_widget = QWidget()
        center_widget.setLayout(center_layout)
        self.setCentralWidget(center_widget)

        # Restores the saved paths for the labels
        saved_excel = self.settings.value("excel_path", "")
        if saved_excel:
            self.excel_path_bar.setText(saved_excel)
    
    def show_terms_of_service(self):
        """Show Terms of Service dialog. Returns True if accepted, False if declined."""
        terms_text = """TRUANCY - TERMS OF SERVICE

This application processes student absence data for educational purposes only.

By using this application, you agree to:
- Use this software only for legitimate truancy tracking purposes..
- Maintain confidentiality...
- Comply with FERPA...
- Not share, distribute, or misuse student information...

The developers are not responsible for:
- Data loss or corruption...
- Misuse of student information...
- Decisions made based on this data...

Do you accept these terms?"""
        
        # Create message box with accept/decline buttons
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Terms of Service")
        msg_box.setText(terms_text)
        msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        # Customize button text
        accept_button = msg_box.button(QMessageBox.StandardButton.Yes)
        accept_button.setText("I Accept")
        
        decline_button = msg_box.button(QMessageBox.StandardButton.No)
        decline_button.setText("I Don't Accept")
        
        # Show dialog and get response
        response = msg_box.exec()
        
        # Return true if accepted, false if declined
        return response == QMessageBox.StandardButton.Yes
    
    def show_instructions(self):
        """Show instructions dialog when help button is clicked."""
        instructions_text = """HOW TO USE TRUANCY RECORDER

Step 1: Load PDF
   Click Load PDF and select the truancy report PDF file.

Step 2: Open Excel File  
   Click Open Excel File and select your tracking spreadsheet.

Step 3: Add Data to Excel
   Click Add Total Absences to Excel and enter the date when prompted.
   Students will be matched by ID/name and data will be added.

Red highlighting = 40+ unexcused hours -requires court intervention

Remember to save your Excel file after adding data"""
        
        # Show instructions in a message box
        QMessageBox.information(self, "Instructions", instructions_text)
