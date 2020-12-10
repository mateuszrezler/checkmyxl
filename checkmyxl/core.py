from xlwings import Book


def main():
    book = Book.caller()
    active_sheet = book.sheets.active
    active_range = active_sheet.used_range
    code_sheet = book.sheets[f'{active_sheet.name}.code']
    code_range = code_sheet.used_range
    conf_sheet = book.sheets['checkmyxl.conf']
    imports = conf_sheet['B1'].value
    exec(imports)
    for cell in code_range:
        if cell.value is not None:
            this_sheet = active_sheet
            this_table = active_range
            this_row = active_sheet.range(
                (cell.row, active_range.column),
                (cell.row, active_range.last_cell.column)
            )
            this_column = active_sheet.range(
                (active_range.row+1, cell.column),
                (active_range.last_cell.row, cell.column)
            )
            this_cell = active_sheet[cell.get_address(False, False)]
            exec(cell.value)

