from PyQt6.QtWidgets import QFileDialog, QMessageBox
from PyQt6.QtCore import QSettings
from docxtpl import DocxTemplate
import os


def open_docx(window):
    saved_docx_dir = window.settings.value("docx_dir", "/home")

    docx_path = QFileDialog.getOpenFileName(window, "Open Word Document", saved_docx_dir, "Word Documents (*.docx)")[0]
    if docx_path:
        try:
            # Opens docx with docxtpl
            window.docx_path_bar.setText(docx_path)
            window.settings.setValue("docx_path", docx_path)
            window.settings.setValue("docx_dir", os.path.dirname(docx_path))
            window.settings.sync()

            template = DocxTemplate(docx_path)

            print(f"Opened Word document: {docx_path}")

            window.docx_opened.emit(docx_path, template)

        except Exception as e:
            print(f"Error opening Word document: {e}")
            QMessageBox.critical(window, "Error", f"Error opening Word document: {e}")