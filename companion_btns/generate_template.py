from PyQt6.QtWidgets import QMessageBox
from docxtpl import DocxTemplate
import shutil
import os

# Pulls the selected row's name from Excel, creates a subdirectory, copies and fills the added word template
def generate_template(window):

    # Ensures prerequisites are met
    if not window.workbook:
        QMessageBox.warning(window, "No Workbook", "Please connect to an Excel file first.")
        return

    if not window.docx_path or not window.docx_template:
        QMessageBox.warning(window, "No Template", "Please load a Word document template first.")
        return

    # Pulls name from the selected Excel row
    try:
        sel = window.workbook.selection

        if sel is None:
            QMessageBox.warning(window, "No Selection", "Please select a row in Excel first.")
            return

        sheet = sel.sheet
        row = sel.row

        header_row = sheet.range("1:1").value
        row_values = sheet.range((row, 1), (row, len(header_row))).value

        row_data = {
            str(h).strip().lower(): v
            for h, v in zip(header_row, row_values)
            if h is not None
        }

        # Returns the cell value from the header key that matches
        def find_value(data, *candidates):
            for key in data:
                if any(c.lower() in key for c in candidates):
                    return str(data[key] or "").strip()
            return ""

        first_name  = find_value(row_data, "first name")
        last_name   = find_value(row_data, "last name")
        parent_name = find_value(row_data, "custodian") or "Custodian"

        if not first_name and not last_name:
            QMessageBox.warning(window, "Empty Name", "The selected row has no name data.")
            return

        full_name = f"{first_name} {last_name}".strip()

    except Exception as e:
        QMessageBox.critical(window, "Excel Error", f"Could not read from Excel:\n {e}")
        return

    # Asks user to confirm the selected name before generating
    confirm = QMessageBox.question(
        window, "Generate Letter",
        f"Generate letter for:\n {full_name} \nContinue?",
        QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel
    )
    if confirm == QMessageBox.StandardButton.Cancel:
        return

    # Sanitizes and Creates subdirectory in the same folder as the loaded template
    docx_dir = window.settings.value("docx_dir", os.path.dirname(window.docx_path))
    sanitize = "".join(c for c in full_name if c.isalnum() or c in (" ", "_", "-")).strip()
    subdir = os.path.join(docx_dir, sanitize)

    try:
        os.makedirs(subdir, exist_ok=True)
    except Exception as e:
        QMessageBox.critical(window, "Directory Error", f"Could not create subdirectory:\n {e}")
        return

    # Copies and Renames the template into the subdirectory
    template_ext = os.path.splitext(window.docx_path)[1]
    new_file_name = f"{sanitize}{template_ext}"
    new_file_path = os.path.join(subdir, new_file_name)

    try:
        shutil.copy2(window.docx_path, new_file_path)
    except Exception as e:
        QMessageBox.critical(window, "Copy Error", f"Could not copy template:\n {e}")
        return

    # Renders the {{ youth_name }} and {{ parent_name }} tag in the copied template
    try:
        template = DocxTemplate(new_file_path)
        context = {
            "youth_name":  full_name,
            "parent_name": parent_name,
        }
        template.render(context)
        template.save(new_file_path)
    except Exception as e:
        QMessageBox.critical(window, "Template Error", f"Could not render Word template:\n {e}")
        return

    print(f"Template generated: {new_file_path}")

    # Completion confirmation Message
    QMessageBox.information(
        window,
        "Done",
        f"Template created for {full_name}:\n{new_file_path}"
    )