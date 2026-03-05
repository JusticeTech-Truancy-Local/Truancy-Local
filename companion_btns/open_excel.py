from PyQt6.QtWidgets import QFileDialog, QMessageBox
import xlwings as xw


def open_excel(self):

    excel_path = QFileDialog.getOpenFileName(self, "Open Excel File", "/home", "Excel Files (*.xlsx *.xls)")[0]
    if excel_path:
        try:
            # Open the Excel file with xlwings
            workbook = xw.Book(excel_path)
            print(f"Opened Excel file: {excel_path}")
            print(f"Workbook has {len(workbook.sheets)} sheet(s)")

            self.excel_opened.emit(workbook)

        except Exception as e:
            print(f"Error opening Excel file: {e}")
            QMessageBox.critical(self, "Error", f"Error opening Excel file: {e}")