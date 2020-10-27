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
                    function(column, cell)


def main():
    book = Book.caller()
    sheet = book.sheets.active
    selection = sheet.used_range
    cc = ColumnChecker(sheet, selection)
    cc.check()


if __name__ == '__main__':
    Book('checkmyxl.xlsm').set_mock_caller()
    main()
