from PyQt6.QtGui import QTextCharFormat, QColor, QFont
from PyQt6.QtWidgets import QMessageBox, QInputDialog
from datetime import datetime

from constructor import Student

BASE_HEADINGS = ["Last Name", "First Name", "Student #", "Age", "Grade", "Custodian", "Address", "Suspension Hours", "Outcome of Correspondence"]

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
        # Get the sheet matching the index in the dropdown
        sheet_idx = window.sheets_combo.currentIndex() - 1
        if sheet_idx < 0:
            # Make new sheet
            sheet = blank_sheet(window.workbook, window.school_name)
        else:
            sheet = window.workbook.sheets[sheet_idx]

        # Find locations of all columns
        column_locs = {}
        for i, col in enumerate(range(1, sheet.used_range.last_cell.column + 1)):
            header_value = sheet.range(f'{_col_letter(col)}1').value
            if header_value in BASE_HEADINGS:
                column_locs[header_value] = i + 1
        for header in BASE_HEADINGS:
            assert(header in column_locs)

        # Ask user for column header - today's date
        suggested_date = datetime.now().strftime("%m/%d/%Y")
        label, ok = QInputDialog.getText(
            window,
            "Column Label",
            "Enter date for this week's absences:",
            text=suggested_date
        )
        # Don't add data if user pressed cancel
        if not ok:
            return

        sheet.select()

        # Insert data before last column
        if "Outcome of Correspondence" in column_locs:
            insert_col = column_locs["Outcome of Correspondence"]  # Insert pushes existing column to the right
        else:
            insert_col = sheet.used_range.last_cell.column + 1  # Default to final column
        
        # Insert two new columns
        sheet.range(f'{_col_letter(insert_col)}:{_col_letter(insert_col)}').api.Insert()
        sheet.range(f'{_col_letter(insert_col + 1)}:{_col_letter(insert_col + 1)}').api.Insert()
        # Shift known column locations to compensate for new columns
        for heading in column_locs:
            if column_locs[heading] >= insert_col:
                column_locs[heading] += 2
 
        # Find the last row by counting up to the last non-clear row
        last_row = sheet.used_range.last_cell.row
        while not sheet.range(f'A{last_row}').value:
            last_row -= 1


        # Add header with user's label
        header_cell_ex = sheet.range(f'{_col_letter(insert_col)}1')
        header_cell_ex.value = f"{label} Excused Absences"
        header_cell_unex = sheet.range(f'{_col_letter(insert_col + 1)}1')
        header_cell_unex.value = f"{label} Unexcused Absences"

        # Clear all colors from the new column
        for row in range(1, last_row + 1):
            sheet.range(f'{_col_letter(insert_col)}{row}').color = None
            sheet.range(f'{_col_letter(insert_col + 1)}{row}').color = None
        
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
        groups = {"2nd time over limit": [],
                  "1st time over limit": [],
                  "No longer over": [],
                  "New students": [],
                  "All students over limit": []}
        
        # Loop through Excel rows and match
        for row in range(2, last_row + 1):
            # Get Student ID from Excel
            excel_student_id = sheet.range(f'{_col_letter(column_locs["Student #"])}{row}').value
            
            # Get names from Excel
            excel_first_name = sheet.range(f'{_col_letter(column_locs["First Name"])}{row}').value
            excel_last_name = sheet.range(f'{_col_letter(column_locs["Last Name"])}{row}').value
            
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

                history = add_student(sheet, matched_student, insert_col, row, column_locs["Suspension Hours"])
                track_group(matched_student, history, groups)

            else:
                # No match found; leave blank no value
                print(f"Row {row}: No match found for {excel_first_name} {excel_last_name} (ID: {excel_student_id})")
                no_match += 1

                # Add "no data" to the new entry
                sheet.range(f'{_col_letter(insert_col)}{row}').value = "no data"
                sheet.range(f'{_col_letter(insert_col + 1)}{row}').value = "no data"
                sheet.range(f'{_col_letter(insert_col)}{row}').color = (217, 217, 217)
                sheet.range(f'{_col_letter(insert_col + 1)}{row}').color = (217, 217, 217)

        # Suffixes to add to grade number. Matched by index in list
        grade_suffix = ["", "st", "nd", "rd", "th", "th", "th", "th", "th", "th", "th", "th", "th"]

        # Add new rows for unmatched students
        extra_row = last_row + 1
        for student in unmatched:
            history = add_student(sheet, student, insert_col, extra_row, column_locs["Suspension Hours"])
            track_group(student, history, groups, True)

            try:
                gradenum = int(student.grade)
                grade = str(gradenum) + grade_suffix[gradenum]
            except TypeError:
                grade = student.grade

            sheet.range(f'{_col_letter(column_locs["Last Name"])}{extra_row}').value = student.lastName
            sheet.range(f'{_col_letter(column_locs["First Name"])}{extra_row}').value = student.firstName
            sheet.range(f'{_col_letter(column_locs["Student #"])}{extra_row}').value = student.id
            sheet.range(f'{_col_letter(column_locs["Age"])}{extra_row}').value = student.age
            sheet.range(f'{_col_letter(column_locs["Grade"])}{extra_row}').value = grade

            extra_row += 1

        # Print summary
        print(f"\n=== SUMMARY ===")
        print(f"No match found: {no_match}")
        print(f"Total rows processed: {last_row - 1}")

        # Write results to status box
        update_status_box(window.status_box, groups, label)
        
    except Exception as e:
        import traceback
        print(f"Error adding absences: {e}")
        print(traceback.format_exc())
        QMessageBox.critical(window, "Error", f"Error adding absences: {e}")

def blank_sheet(workbook, name):
    # Create new sheet
    sheet = workbook.sheets.add(name=name)

    sheet.range('A2').value = ""
    # Add base headings to sheet
    for i, heading in enumerate(BASE_HEADINGS):
        sheet.range(f'{_col_letter(i + 1)}1').value = heading
        sheet.range(f'{_col_letter(i + 1)}1').font.bold = True

    suspension_idx = BASE_HEADINGS.index("Suspension Hours") + 1
    sheet.range(f'{_col_letter(suspension_idx)}1').color = (255, 255, 0)

    return sheet

def add_student(sheet, student, column, row, suspension_col):
    history = [] # Last three weeks' status. True = over limit, False = under limit, None = no data

    # March thru previous weeks to record history of being over the limit
    for c in range(max(9, column-4), column, 2):
        val_ex = sheet.range(f'{_col_letter(c)}{row}').value
        val_unex = sheet.range(f'{_col_letter(c + 1)}{row}').value
        try:
            val_int = int(val_ex) + int(val_unex)
            history.append(val_int >= Student.redThreshold)
        except TypeError:
            history.append(None)

    if student.unexcused:
        try:
            excused = float(student.excused)
            unexcused = float(student.unexcused)
            suspension = float(student.suspension)
            total_no_suspension = float(student.absenceTotal) - suspension
            cell_ex = sheet.range(f'{_col_letter(column)}{row}')
            cell_ex.value = excused
            cell_unex = sheet.range(f'{_col_letter(column + 1)}{row}')
            cell_unex.value = unexcused

            # Check for mismatch with report's total and calculated total
            if excused + unexcused != total_no_suspension:
                print(f"!!! Total hours mismatch for {student.firstName} {student.lastName}"
                      f": Excel says {excused + unexcused}, PDF says {total_no_suspension}")

            # Update suspensions column
            sheet.range(f'{_col_letter(suspension_col)}{row}').value = suspension
            sheet.range(f'{_col_letter(suspension_col)}{row}').color = (255, 255, 0)

            # Color code based on absence hours
            if unexcused + excused >= Student.redThreshold:
                # Red for over limit
                cell_ex.color = (255, 0, 0)  # Red
                cell_unex.color = (255, 0, 0)
                history.append(True)
            else:
                history.append(False)

        except (ValueError, TypeError):
            print(f"Warning: Could not convert an absence total")
            # Invalid data
            sheet.range(f'{_col_letter(column)}{row}').value = "no data"
            sheet.range(f'{_col_letter(column + 1)}{row}').value = "no data"
            sheet.range(f'{_col_letter(column)}{row}').color = (217, 217, 217)
            sheet.range(f'{_col_letter(column + 1)}{row}').color = (217, 217, 217)
            history.append(None)
    else:
        # Student matched but has no absence data; no color
        sheet.range(f'{_col_letter(column)}{row}').value = "no data"
        sheet.range(f'{_col_letter(column + 1)}{row}').value = "no data"
        sheet.range(f'{_col_letter(column)}{row}').color = (217, 217, 217)
        sheet.range(f'{_col_letter(column + 1)}{row}').color = (217, 217, 217)
        history.append(None)

    return history


def track_group(student, history, groups, is_new=False):
    # Add students to groups based on whether they were over limits the last three weeks
    updated_history = [False * (3 - len(history))] + [bool(x) for x in history]

    if is_new:
        groups["New students"].append(student)

    if updated_history[-1]:
        groups["All students over limit"].append(student)
        if updated_history[-2]:
            if not updated_history[-3]:
                groups["2nd time over limit"].append(student)
        else:
            groups["1st time over limit"].append(student)
    elif updated_history[-2]:
            groups["No longer over"].append(student)


def update_status_box(status_box, groups, label):
    cursor = status_box.textCursor()
    format = QTextCharFormat()

    format.setFontUnderline(True)
    cursor.insertText(label, format)
    format.setFontUnderline(False)

    order = ["New students", "1st time over limit", "2nd time over limit", "No longer over", "All students over limit"]
    highlight = {"2nd time over limit"}

    for group in order:
        format.setFontWeight(QFont.Weight.Bold)
        cursor.insertText("\n"+group+"\n", format)
        format.setFontWeight(QFont.Weight.Normal)

        if group in highlight:
            format.setBackground(QColor(255, 0, 0, 80))

        for student in groups[group]:
            cursor.insertText(f"{student.lastName}, {student.firstName}\n", format)

        format.clearBackground()

    cursor.setCharFormat(format)


def _col_letter(col_num):
    """Convert column number to Excel column letter (1=A, 2=B, ...)"""
    string = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        string = chr(65 + remainder) + string
    return string
