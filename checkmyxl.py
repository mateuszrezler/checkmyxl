from tasks import TASKS
from xlwings import Book


class ColumnChecker(object):

    def __init__(self, sheet, selection):
        self.sheet = sheet
        self.selection = selection

    def check(self):
        for column in self.selection.columns:
            for row in self.selection.rows:
                col_num = column.column-1
                row_num = row.row-1
                cell = self.sheet[(row_num, col_num)]
                if col_num in TASKS:
                    function = TASKS[col_num]
                    function(row, column, cell)


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
