from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QTextCharFormat, QFont, QTextCursor, QColor
from PyQt6.QtWidgets import QTextEdit, QApplication


class StatusBox(QTextEdit):
    go_to_cell = pyqtSignal(str, str)

    def __init__(self):
        super().__init__()

        self.anchor = None
        self.setReadOnly(True)
        self.setTabStopDistance(8 * self.fontMetrics().horizontalAdvance(' '))

    def mouseReleaseEvent(self, e):
        if self.anchor:
            QApplication.setOverrideCursor(Qt.CursorShape.ArrowCursor)
            anchorparts = self.anchor.split("!")
            self.anchor = None
            self.go_to_cell.emit(anchorparts[0], anchorparts[1])

    def mouseMoveEvent(self, e):
        anchor = self.anchorAt(e.pos())
        if anchor:
            self.anchor = anchor
            QApplication.setOverrideCursor(Qt.CursorShape.PointingHandCursor)
        elif self.anchor:
            QApplication.setOverrideCursor(Qt.CursorShape.ArrowCursor)
            self.anchor = None


    def report_update(self, groups, label, threshold, sheet, insert_cols):
        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.Start)
        format = QTextCharFormat()

        format.setFontUnderline(True)
        cursor.insertText(label + "\n", format)
        format.setFontUnderline(False)

        cursor.insertTable(1 + len(groups[1]) + len(groups[2]) + len(groups[3]), 4)

        format.setFontWeight(QFont.Weight.Bold)
        cursor.insertText("Student", format)
        cursor.movePosition(QTextCursor.MoveOperation.NextCell)
        cursor.insertText(f"Consecutive\nweeks\nover {threshold} hrs", format)
        cursor.movePosition(QTextCursor.MoveOperation.NextCell)
        cursor.insertText(f"Prelim\nLetter", format)
        cursor.movePosition(QTextCursor.MoveOperation.NextCell)
        cursor.insertText(f"Mediation\nLetter", format)
        format.setFontWeight(QFont.Weight.Normal)
        cursor.movePosition(QTextCursor.MoveOperation.NextCell)

        for student in groups[1]:
            self.write_student(student[0], format, cursor, sheet, student[1], insert_cols)
            cursor.movePosition(QTextCursor.MoveOperation.NextCell)
            cursor.insertText(f" 1\t🟥", format)
            cursor.movePosition(QTextCursor.MoveOperation.NextCell)
            cursor.insertText(student[2], format)
            cursor.movePosition(QTextCursor.MoveOperation.NextCell)
            cursor.insertText(student[3], format)
            cursor.movePosition(QTextCursor.MoveOperation.NextCell)
        for student in groups[2]:
            self.write_student(student[0], format, cursor, sheet, student[1], insert_cols)
            cursor.movePosition(QTextCursor.MoveOperation.NextCell)
            cursor.insertText(f" 2\t🟥🟥", format)
            cursor.movePosition(QTextCursor.MoveOperation.NextCell)
            cursor.insertText(student[2], format)
            cursor.movePosition(QTextCursor.MoveOperation.NextCell)
            cursor.insertText(student[3], format)
            cursor.movePosition(QTextCursor.MoveOperation.NextCell)
        for student in groups[3]:
            self.write_student(student[0], format, cursor, sheet, student[1], insert_cols)
            cursor.movePosition(QTextCursor.MoveOperation.NextCell)
            cursor.insertText(f" 3+\t🟥🟥🟥...", format)
            cursor.movePosition(QTextCursor.MoveOperation.NextCell)
            cursor.insertText(student[2], format)
            cursor.movePosition(QTextCursor.MoveOperation.NextCell)
            cursor.insertText(student[3], format)
            cursor.movePosition(QTextCursor.MoveOperation.NextCell)

        cursor.movePosition(QTextCursor.MoveOperation.NextBlock)

        format.setFontWeight(QFont.Weight.Bold)
        cursor.insertText(f"\n\nDropped below {threshold} hrs\n", format)
        format.setFontWeight(QFont.Weight.Normal)
        for student in groups[-1]:
            self.write_student(student[0], format, cursor, sheet, student[1], insert_cols)
            cursor.insertText("\n", format)

        format.setFontWeight(QFont.Weight.Bold)
        cursor.insertText(f"\nNew students\n", format)
        format.setFontWeight(QFont.Weight.Normal)
        for student in groups[0]:
            self.write_student(student[0], format, cursor, sheet, student[1], insert_cols)
            cursor.insertText("\n", format)

        cursor.insertText("\n", format)

        cursor.setCharFormat(format)

    def write_student(self, student, format, cursor, sheet, row, cols):
        format.setFontUnderline(True)
        format.setAnchor(True)
        format.setAnchorHref(sheet + "!" + cols[0] + str(row) + ":" + cols[1] + str(row))
        format.setForeground(QColor('blue'))
        cursor.insertText(f"{student.lastName}, {student.firstName}", format)
        format.setAnchor(False)
        format.setAnchorHref(None)
        format.setForeground(QColor('black'))
        format.setFontUnderline(False)