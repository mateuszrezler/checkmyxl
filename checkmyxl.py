from argparse import ArgumentParser
from src.tasks import TASKS
from src.utils import get_abs_path, load_config, load_sample, load_sheet, \
    skip_header
from sys import argv
from xlwings import App, apps, Book


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
                cell.c, cell.r = column, row
                task = TASKS[col_num]
                if callable(task):
                    function = task
                    kwargs = {}
                elif isinstance(task, tuple):
                    function = task[0]
                    kwargs = task[1]
                else:
                    error_msg = 'Task should be a function or a two-element' \
                        + ' tuple containing in turn: a function and' \
                        + ' a dictionary of its keyword arguments.'
                    raise TypeError(error_msg)
                function(cell, **kwargs)


def main(selection=None):
    config = load_config()
    book, sheet = load_sheet()
    if selection:
        header = False
        selection = sheet[selection]
    else:
        selection = sheet.used_range
        if config['header']:
            selection = skip_header(sheet, selection)
    if config['reset_colors']:
        selection.color = None
    cc = ColumnChecker(sheet, selection)
    cc.check()


def make_sample():
    book, sheet = load_sheet()
    sheet['A1'].value = load_sample()
    sheet.autofit('columns')


def parse_args(args):
    ap = ArgumentParser()
    ap.add_argument('-ms', '--make-sample', action='store_true')
    return ap.parse_args(args)


def start(args=None):
    if not args:
        args = parse_args(argv[1:])
    config = load_config()
    excel_path = get_abs_path(config['excel_file'])
    if apps.count == 0:
        App()
    Book(excel_path).set_mock_caller()
    if args.make_sample:
        make_sample()
    main()


def undo():
    config = load_config()
    book, sheet = load_sheet()
    book.close()
    excel_path = get_abs_path(config['excel_file'])
    Book(excel_path)


if __name__ == '__main__':
    start()

