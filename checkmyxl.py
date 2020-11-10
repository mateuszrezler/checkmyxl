from pandas import read_csv
from sys import argv
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
    sheet = load_sheet()
    if selection:
        header = False
        selection = sheet[selection]
    else:
        selection = sheet.used_range
    if header:
        selection = skip_header(sheet, selection)
    cc = ColumnChecker(sheet, selection)
    cc.check()


def make_sample():
    sheet = load_sheet()
    sample = read_csv('sample.csv', header=None)
    sheet['A1'].value = sample.values
    sheet.autofit('columns')


def skip_header(sheet, selection):
    return sheet.range(
        (selection.row+1, selection.column),
        (selection.last_cell.row, selection.last_cell.column))


def load_sheet():
    book = Book.caller()
    sheet = book.sheets.active
    return sheet


if __name__ == '__main__':
    Book('checkmyxl.xlsm').set_mock_caller()
    if len(argv) == 2 and argv[1] == 'make_sample':
        make_sample()
    else:
        main()
