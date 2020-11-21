from .tasks import TASKS
from .utils import get_abs_path, load_config, load_sheet, load_sample
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


def make_sample():
    book, sheet = load_sheet()
    sheet['A1'].value = load_sample()
    sheet.autofit('columns')


def undo():
    config = load_config()
    book, sheet = load_sheet()
    book.close()
    excel_path = get_abs_path(config['excel_file'])
    Book(excel_path)

