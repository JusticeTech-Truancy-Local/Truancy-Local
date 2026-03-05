from PyQt6.QtWidgets import QMessageBox, QInputDialog
from datetime import datetime

from constructor import Student

def add_report_to_sheet(window):
    """Add Total Absences column from PDF data to Excel file with ID/name matching and color coding"""
    
    # Check if we have both PDF data and Excel file
    if not window.check_files_ready():
        if not window.students:
            QMessageBox.warning(window, "No PDF Data", "Please load a PDF file first")
            return

        if not window.workbook:
            QMessageBox.warning(window, "No Excel File", "Please open an Excel file first")
            return
    
    try:
        # Get the active sheet
        sheet = window.workbook.sheets.active
        
        # Ask user for column header - today's date
        suggested_date = datetime.now().strftime("%m/%d/%Y")
        label, ok = QInputDialog.getText(
            window,
            "Column Label",
            "Enter date for this week's absences:",
            text=suggested_date
        )

        # Update date in totals columns
        sheet.range(f'I1').value = f"{label} Excused Absences"
        sheet.range(f'J1').value = f"{label} Unexcused Absences"
        sheet.range(f'K1').value = f"{label} Total Absences (minus suspension hours)"

        # Find last column Outcome of Correspondence
        outcome_col = None
        for col in range(1, sheet.used_range.last_cell.column + 1):
            header_value = sheet.range(f'{_col_letter(col)}1').value
            if header_value and "Outcome of Correspondence" in str(header_value):
                outcome_col = col
                break
        
        # Insert data before last column
        if outcome_col:
            insert_col = outcome_col  # Insert pushes existing column to the right
        else:
            insert_col = 22  # Default to column V 
        
        # Insert a new column
        sheet.range(f'{_col_letter(insert_col)}:{_col_letter(insert_col)}').api.Insert()
 
        # Find the last row by counting up to the last non-clear row
        last_row = sheet.used_range.last_cell.row
        while not sheet.range(f'A{last_row}').value:
            last_row -= 1

        # Clear all colors from the new column
        for row in range(2, last_row + 1):
            sheet.range(f'{_col_letter(insert_col)}{row}').color = None 

        # Add header with user's label
        header_cell = sheet.range(f'{_col_letter(insert_col)}1')
        header_cell.value = f"{label} Unexcused Absences"
        
        print(f" ADDING ABSENCES WITH MATCHING ")
        
        # Build lookup dictionaries from PDF students
        pdf_by_id = {}  # {student_id: student_object}
        pdf_by_name = {}  # {(last_name, first_name): student_object}
        unmatched = set()
        
        for student in window.students:
            # Add to ID lookup
            if student.id:
                pdf_by_id[str(student.id).strip()] = student
            
            # Add to name lookup
            if student.lastName and student.firstName:
                last = str(student.lastName).strip().lower()
                first = str(student.firstName).strip().lower()
                pdf_by_name[(last, first)] = student

            # Add to set tracking students that haven't been matched
            unmatched.add(student)
        
        print(f"PDF students indexed: {len(pdf_by_id)} by ID, {len(pdf_by_name)} by name")
        
        # Get last row in Excel
        print(f"Excel has {last_row - 1} data rows (rows 2-{last_row})")
        
        # Track data
        no_match = 0
        over_limit_count = 0
        
        # Loop through Excel rows and match
        for row in range(2, last_row + 1):
            # Get Student ID from Excel (column C)
            excel_student_id = sheet.range(f'C{row}').value
            
            # Get names from Excel
            excel_first_name = sheet.range(f'B{row}').value
            excel_last_name = sheet.range(f'A{row}').value
            
            matched_student = None
            
            # Match by Student ID
            if excel_student_id:
                excel_student_id_str = str(int(excel_student_id)).strip()
                if excel_student_id_str in pdf_by_id:
                    matched_student = pdf_by_id[excel_student_id_str]
                    print(f"Row {row}: Matched by ID {excel_student_id_str}")
            
            # Check if name matches just in case
            if matched_student and excel_last_name and excel_first_name:
                last = str(excel_last_name).strip().lower()
                first = str(excel_first_name).strip().lower()

                if (last, first) in pdf_by_name and pdf_by_name[(last, first)] != matched_student:
                    print(f"!!! Student name mismatch: ID {excel_student_id} matches {matched_student.firstName} "
                          f"{matched_student.lastName} in Excel, {first} {last} in PDF")
            
            # Write Unexcused if matches are found
            if matched_student:
                # Remove from unmatched set
                unmatched.remove(matched_student)

                has_data, over_limit = add_student(sheet, matched_student, insert_col, row)
                if over_limit:
                    over_limit_count += 1

            else:
                # No match found; leave blank no value no color
                print(f"Row {row}: No match found for {excel_first_name} {excel_last_name} (ID: {excel_student_id})")
                no_match += 1

                # Add "no data" to the new entry
                sheet.range(f'{_col_letter(insert_col)}{row}').value = "no data"

        # Add new rows for unmatched students
        extra_row = last_row + 1
        for student in unmatched:
            sheet.range(f'B{extra_row}')
            add_student(sheet, student, insert_col, extra_row)

            sheet.range(f'A{extra_row}').value = student.lastName
            sheet.range(f'B{extra_row}').value = student.firstName
            sheet.range(f'C{extra_row}').value = student.id
            sheet.range(f'D{extra_row}').value = student.age
            sheet.range(f'E{extra_row}').value = student.grade

            extra_row += 1

        # Print summary
        print(f"\n=== SUMMARY ===")
        print(f"No match found: {no_match}")
        print(f"Over limit (Red): {over_limit_count}")
        print(f"Total rows processed: {last_row - 1}")

        
        # # Show results to user
        # QMessageBox.information(
        #     window,
        #     "Success",
        #     f"Added Total Absences column with color coding!\n\n"
        #     f"Matched by Student ID: {matched_by_id}\n"
        #     f"Matched by Name: {matched_by_name}\n"
        #     f"No match found: {no_match}\n\n"
        #     f" High Risk (40+ hrs): {high_risk_count}\n"
        #     f" Medium Risk (21-39 hrs): {medium_risk_count}\n"
        #     f" Low Risk (0-20 hrs): {low_risk_count}\n\n"
        #     f"Total rows: {last_row - 1}"
        # )
        
    except Exception as e:
        import traceback
        print(f"Error adding absences: {e}")
        print(traceback.format_exc())
        QMessageBox.critical(window, "Error", f"Error adding absences: {e}")


def add_student(sheet, student, column, row):
    over_limit = False
    has_data = False

    if student.unexcused:
        try:
            excused = float(student.excused)
            unexcused = float(student.unexcused)
            suspension = float(student.suspension)
            total_no_suspension = float(student.absenceTotal) - suspension
            cell = sheet.range(f'{_col_letter(column)}{row}')
            cell.value = unexcused

            # Color code based on absence hours
            if unexcused >= Student.redThreshold:
                # Red for over limit
                cell.color = (255, 0, 0)  # Red
                over_limit = True

            # Update totals columns
            sheet.range(f'H{row}').value = suspension
            sheet.range(f'I{row}').value = excused
            sheet.range(f'J{row}').value = unexcused
            # Check for mismatch with report's total and calculated total
            if str(excused + unexcused) != student.absenceTotal:
                print(f"!!! Total hours mismatch for {student.firstName} {student.lastName}"
                      f": Excel says {excused + unexcused}, PDF says {student.absenceTotal}")
            sheet.range(f'K{row}').value = total_no_suspension

            has_data = True

        except (ValueError, TypeError):
            print(f"Warning: Could not convert an absence total")
            # Invalid data
            sheet.range(f'{_col_letter(column)}{row}').value = "no data"
    else:
        # Student matched but has no absence data; no color
        sheet.range(f'{_col_letter(column)}{row}').value = "no data"

    return has_data, over_limit


def _col_letter(col_num):
    """Convert column number to Excel column letter (1=A, 2=B, ...)"""
    string = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        string = chr(65 + remainder) + string
    return string
