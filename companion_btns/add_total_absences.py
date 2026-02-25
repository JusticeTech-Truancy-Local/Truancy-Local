from PyQt6.QtWidgets import QMessageBox
import traceback

def add_total_absences(self):
    """Add Total Absences column from PDF data to Excel file"""
    
    # Check if we have both PDF data and Excel file
    if not self.students:
        QMessageBox.warning(self, "No PDF Data", "Please load a PDF file first")
        return
    
    if not self.workbook:
        QMessageBox.warning(self, "No Excel File", "Please open an Excel file first")
        return
    
    try:
        # Get the active sheet
        sheet = self.workbook.sheets.active
        
        # Insert column at position 22 (between column 21 and current 22)
        insert_col = 22
        
        # Insert a new column
        sheet.range(f'{self._col_letter(insert_col)}:{self._col_letter(insert_col)}').api.Insert()
        
        # Add header
        sheet.range(f'{self._col_letter(insert_col)}1').value = "Total Absences (from PDF)"
        
        print(f"\n=== ADDING TOTAL ABSENCES ===")
        
        # Add Total Absences from PDF starting at row 2
        added_count = 0
        for i, student in enumerate(self.students):
            row = i + 2  # Start at row 2 
            
            if student.absenceTotal:
                try:
                    total_abs = float(student.absenceTotal)
                    sheet.range(f'{self._col_letter(insert_col)}{row}').value = total_abs
                    print(f"Row {row}: Added {total_abs} for {student.firstName} {student.lastName} (ID: {student.id})")
                    added_count += 1
                except (ValueError, TypeError):
                    print(f"Warning: Could not convert absenceTotal for student {student.id}: {student.absenceTotal}")
        
        print(f"\n=== SUMMARY ===")
        print(f"Total absences added: {added_count}")
        
        QMessageBox.information(self, "Success", 
            f"Added Total Absences for {added_count} students\n"
            f"Values added to column {self._col_letter(insert_col)} starting at row 2")
        
    except Exception as e:
        import traceback
        print(f"Error adding total absences: {e}")
        print(traceback.format_exc())
        QMessageBox.critical(self, "Error", f"Error adding total absences: {e}")

def _col_letter(self, col_num):
    """Convert column number to Excel column letter (1=A, 2=B, )"""
    string = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        string = chr(65 + remainder) + string
    return string