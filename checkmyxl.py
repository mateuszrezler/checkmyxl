from tasks import TASKS
from xlwings import Book


class ColumnChecker(object):

    def __init__(self, sheet, selection):
        self.sheet = sheet
        self.selection = selection

    def check(self):
        for column in self.selection.columns:
            col_num = column.column-1
            if col_num not in TASKS:
                continue
            for row in self.selection.rows:
                row_num = row.row-1
                cell = self.sheet[(row_num, col_num)]
                if callable(TASKS[col_num]):
                    function = TASKS[col_num]
                    function(row, column, cell)
                elif isinstance(TASKS[col_num], tuple):
                    function = TASKS[col_num][0]
                    args = TASKS[col_num][1:]
                    function(row, column, cell, *args)


def main(selection=None, header=True):
    book = Book.caller()
    sheet = book.sheets.active
    if selection:
        header = False
        selection = sheet[selection]
    else:
        selection = sheet.used_range
    if header:
        selection = skip_header(sheet, selection)
    cc = ColumnChecker(sheet, selection)
    cc.check()


def skip_header(sheet, selection):
    return sheet.range(
        (selection.row+1, selection.column),
        (selection.last_cell.row, selection.last_cell.column))


if __name__ == '__main__':
    Book('checkmyxl.xlsm').set_mock_caller()
    main()

