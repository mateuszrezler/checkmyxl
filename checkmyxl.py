from sys import argv
from src.tasks import TASKS
from src.utils import get_abs_path, load_config, load_sample, load_sheet, \
    skip_header


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
                task = TASKS[col_num]
                if callable(task):
                    function = task
                    function(row, column, cell)
                elif isinstance(task, tuple):
                    function = task[0]
                    args = task[1:]
                    function(row, column, cell, *args)


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


if __name__ == '__main__':
    if len(argv) == 2 and argv[1] == 'make_sample':
        make_sample()
    else:
        main()

