from json import load as load_json
from os.path import dirname, join as join_path, realpath
from pandas import read_csv
from sys import argv
from src.tasks import TASKS
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


def _load_config():
    config_path = _get_abs_path('config.json')
    with open(config_path) as cf:
        config = load_json(cf)
    return config


def _load_sample():
    config = _load_config()
    sample_path = _get_abs_path(config['sample_file'])
    sample = read_csv(sample_path, header=None).values
    return sample


def _load_sheet():
    config = _load_config()
    excel_path = _get_abs_path(config['excel_file'])
    Book(excel_path).set_mock_caller()
    book = Book.caller()
    sheet = book.sheets.active
    return book, sheet


def _get_abs_path(rel_path):
    parent_dir = dirname(realpath(__file__))
    return join_path(parent_dir, *rel_path.split('/'))


def _skip_header(sheet, selection):
    return sheet.range(
        (selection.row+1, selection.column),
        (selection.last_cell.row, selection.last_cell.column))


def main(selection=None):
    config = _load_config()
    book, sheet = _load_sheet()
    if selection:
        header = False
        selection = sheet[selection]
    else:
        selection = sheet.used_range
    if config['header']:
        selection = _skip_header(sheet, selection)
    if config['reset_colors']:
        selection.color = None
    cc = ColumnChecker(sheet, selection)
    cc.check()


def make_sample():
    book, sheet = _load_sheet()
    sheet['A1'].value = _load_sample()
    sheet.autofit('columns')


if __name__ == '__main__':
    if len(argv) == 2 and argv[1] == 'make_sample':
        make_sample()
    else:
        main()

