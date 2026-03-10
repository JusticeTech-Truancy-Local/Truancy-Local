import xlwings as xw
import sys
import json

def compare_sheets(workbook, sha, shb):
    """
    Return dictionary of cells with different data between two sheets in a workbook
    :param workbook: Opened workbook in xlwings
    :param sha: Name of sheet A
    :param shb: Name of sheet B
    :return: Nested dict of cell addresses that have differences between the sheets mapped to dicts of those differences
    """
    sheet_a = workbook.sheets(sha)
    sheet_b = workbook.sheets(shb)
    range_a = sheet_a.used_range
    range_b = sheet_b.used_range
    # Pick a range that encompasses both ranges
    rows = max(range_a.rows.count, range_b.rows.count)
    cols = max(range_a.columns.count, range_b.columns.count)

    diff_values = {}
    diff_colors = {}

    # Iterate all cells to check individually
    for r in range(rows):
        for c in range(cols):
            addr = (c, r)
            cell_a = sheet_a[addr]
            cell_b = sheet_b[addr]
            if cell_a.value != cell_b.value:
                diff_values[cell_a.address] = [cell_a.value, cell_b.value]
            if cell_a.color != cell_b.color:
                diff_colors[cell_a.address] = [cell_a.color, cell_b.color]
    return {"values": diff_values, "colors": diff_colors}


if __name__ == "__main__":
    # Open Excel invisibly
    app = xw.App(visible=False, add_book=False)
    workbook = app.books.open(sys.argv[1])
    print(json.dumps(compare_sheets(workbook, sys.argv[2], sys.argv[3]), indent=2))
    workbook.close()
    app.kill()