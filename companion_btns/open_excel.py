from PyQt6.QtWidgets import QFileDialog, QMessageBox
import xlwings as xw
import os

def open_excel(self):
    saved_excel_dir = self.settings.value("excel_dir", "/home")

    excel_path = QFileDialog.getOpenFileName(self, "Open Excel File", saved_excel_dir, "Excel Files (*.xlsx *.xls)")[0]
    if excel_path:
        try:
            # Open the Excel file with xlwings
            self.workbook = xw.Book(excel_path)
            self.excel_path_bar.setText(excel_path)
            self.settings.setValue("excel_path", excel_path)
            self.settings.setValue("excel_dir", os.path.dirname(excel_path))
            self.settings.sync()

            print(f"Opened Excel file: {excel_path}")
            print(f"Workbook has {len(self.workbook.sheets)} sheet(s)")
            QMessageBox.information(self, "Success", f"Opened Excel file with {len(self.workbook.sheets)} sheet(s)")

        except Exception as e:
            print(f"Error opening Excel file: {e}")
            QMessageBox.critical(self, "Error", f"Error opening Excel file: {e}")