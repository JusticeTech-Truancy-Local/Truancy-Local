from PyQt6.QtWidgets import QMessageBox, QInputDialog
from datetime import datetime

def add_report_to_sheet(window):
    """Add Total Absences column from PDF data to Excel file with ID/name matching and color coding"""
    
    # Check if we have both PDF data and Excel file
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
 
        # Clear all colors from the new column 
        last_row = sheet.used_range.last_cell.row
        for row in range(2, last_row + 1):
            sheet.range(f'{_col_letter(insert_col)}{row}').color = None 

        # Add header with user's label
        header_cell = sheet.range(f'{_col_letter(insert_col)}1')
        header_cell.value = f"Total Absences - {label}"
        header_cell.color = (200, 200, 200) 
        
        print(f" ADDING TOTAL ABSENCES WITH MATCHING ")
        
        # Build lookup dictionaries from PDF students
        pdf_by_id = {}  # {student_id: student_object}
        pdf_by_name = {}  # {(last_name, first_name): student_object}
        
        for student in window.students:
            # Add to ID lookup
            if student.id:
                pdf_by_id[str(student.id).strip()] = student
            
            # Add to name lookup
            if student.lastName and student.firstName:
                last = str(student.lastName).strip().lower()
                first = str(student.firstName).strip().lower()
                pdf_by_name[(last, first)] = student
        
        print(f"PDF students indexed: {len(pdf_by_id)} by ID, {len(pdf_by_name)} by name")
        
        # Get last row in Excel
        last_row = sheet.used_range.last_cell.row
        print(f"Excel has {last_row - 1} data rows (rows 2-{last_row})")
        
        # Track data
        matched_by_id = 0
        matched_by_name = 0
        no_match = 0
        high_risk_count = 0
        medium_risk_count = 0
        low_risk_count = 0
        
        # Loop through Excel rows and match
        for row in range(2, last_row + 1):
            # Get Student ID from Excel (column Z/column 26)
            excel_student_id = sheet.range(f'Z{row}').value
            
            # Get names from Excel
            excel_first_name = sheet.range(f'B{row}').value
            excel_last_name = sheet.range(f'A{row}').value
            
            matched_student = None
            match_type = None
            
            # Match by Student ID first
            if excel_student_id:
                excel_student_id_str = str(excel_student_id).strip()
                if excel_student_id_str in pdf_by_id:
                    matched_student = pdf_by_id[excel_student_id_str]
                    match_type = 'id'
                    print(f"Row {row}: Matched by ID {excel_student_id_str}")
            
            # If no ID matches than by name
            if not matched_student and excel_last_name and excel_first_name:
                last = str(excel_last_name).strip().lower()
                first = str(excel_first_name).strip().lower()
                
                if (last, first) in pdf_by_name:
                    matched_student = pdf_by_name[(last, first)]
                    match_type = 'name'
                    print(f"Row {row}: Matched by name {first} {last}")
            
            # Write Total Absences if matches are found
            if matched_student:
                if matched_student.absenceTotal:
                    try:
                        total_abs = float(matched_student.absenceTotal)
                        cell = sheet.range(f'{_col_letter(insert_col)}{row}')
                        cell.value = total_abs
                        
                        # Color code based on absence hours
                        if total_abs >= 40:
                            # Red for high risk (40+ hours)
                            cell.color = (255, 0, 0)  # Red
                            high_risk_count += 1
                            # Ask Stacey if we should email parents automatically when student hits 40+ hours
                        elif total_abs >= 21:
                            # Yellow for medium risk (21-39 hours)
                            cell.color = (255, 255, 200)  # yellow
                            medium_risk_count += 1
                        else:
                            # Green for low risk (0-20 hours)
                            cell.color = (200, 255, 200)  # green
                            low_risk_count += 1
                        
                        # Track match type
                        if match_type == 'id':
                            matched_by_id += 1
                        else:
                            matched_by_name += 1
                            
                    except (ValueError, TypeError):
                        print(f"Warning: Could not convert absenceTotal: {matched_student.absenceTotal}")
                        # Invalid data
                        sheet.range(f'{_col_letter(insert_col)}{row}').value = "N/A"
                        if match_type == 'id':
                            matched_by_id += 1
                        else:
                            matched_by_name += 1
                else:
                    # Student matched but has no absence data; no color
                    sheet.range(f'{_col_letter(insert_col)}{row}').value = 0
                    if match_type == 'id':
                        matched_by_id += 1
                    else:
                        matched_by_name += 1
            else:
                # No match found; leave blank no value no color
                print(f"Row {row}: No match found for {excel_first_name} {excel_last_name} (ID: {excel_student_id})")
                no_match += 1
        
        # Print summary
        total_matched = matched_by_id + matched_by_name
        print(f"\n=== SUMMARY ===")
        print(f"Matched by Student ID: {matched_by_id}")
        print(f"Matched by Name: {matched_by_name}")
        print(f"No match found: {no_match}")
        print(f"High Risk (Red): {high_risk_count}")
        print(f"Medium Risk (Yellow): {medium_risk_count}")
        print(f"Low Risk (Green): {low_risk_count}")
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
        print(f"Error adding total absences: {e}")
        print(traceback.format_exc())
        QMessageBox.critical(window, "Error", f"Error adding total absences: {e}")


def _col_letter(col_num):
    """Convert column number to Excel column letter (1=A, 2=B, ...)"""
    string = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        string = chr(65 + remainder) + string
    return string
