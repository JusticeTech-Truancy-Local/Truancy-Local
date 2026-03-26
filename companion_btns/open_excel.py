from PyQt6.QtWidgets import QFileDialog, QMessageBox
from PyQt6.QtCore import QSettings
import xlwings as xw
import os


def open_excel(window):
    saved_excel_dir = window.settings.value("excel_dir", "/home")

    excel_path = QFileDialog.getOpenFileName(window, "Open Excel File", saved_excel_dir, "Excel Files (*.xlsx *.xls)")[0]
    if excel_path:
        try:
            # Open the Excel file with xlwings
            window.excel_path_bar.setText(excel_path)
            window.settings.setValue("excel_path", excel_path)
            window.settings.setValue("excel_dir", os.path.dirname(excel_path))
            window.settings.sync()

            workbook = xw.Book(excel_path)

            print(f"Opened Excel file: {excel_path}")
            print(f"Workbook has {len(workbook.sheets)} sheet(s)")

            window.excel_opened.emit(workbook)

        except Exception as e:
            print(f"Error opening Excel file: {e}")
            QMessageBox.critical(window, "Error", f"Error opening Excel file: {e}")