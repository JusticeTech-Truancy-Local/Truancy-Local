from PyQt6.QtWidgets import QFileDialog, QMessageBox
import xlwings as xw
import os
import threading

def launch_excel(path, window):
    workbook = xw.Book(path)
    print(f"Opened Excel file: {path}")
    print(f"Workbook has {len(workbook.sheets)} sheet(s)")

    window.excel_opened.emit(path)

def open_excel(window):
    saved_excel_dir = window.settings.value("excel_dir", "/home")

    excel_path = QFileDialog.getOpenFileName(window, "Open Excel File", saved_excel_dir, "Excel Files (*.xlsx *.xls)")[0]
    if excel_path:
        try:
            # Open the Excel file with xlwings
            window.settings.setValue("excel_path", excel_path)
            window.settings.setValue("excel_dir", os.path.dirname(excel_path))
            window.settings.sync()
            window.excel_path_bar.setText(excel_path)

            # Open workbook first in new thread so GUI is not interrupted
            t = threading.Thread(target=launch_excel, args=(excel_path, window))
            t.setDaemon(True)
            t.start()


        except Exception as e:
            print(f"Error opening Excel file: {e}")
            QMessageBox.critical(window, "Error", f"Error opening Excel file: {e}")
