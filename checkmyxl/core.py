from checkmyxl.utils import skip_header
from config.tasks import TASKS


class ColumnChecker(object):

    def __init__(self, sheet, selection, header, reset_colors):
        self.sheet = sheet
        self.selection = selection
        self.header = header
        self.reset_colors = reset_colors
        self.selection = self.set_selection()

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
                    raise TypeError('Task should be a function or a tuple.')
                function(cell, **kwargs)

    def set_selection(self):
        if self.selection:
            self.header = False
            self.selection = self.sheet[self.selection]
        else:
            self.selection = self.sheet.used_range
            if self.header:
                self.selection = skip_header(self.sheet, self.selection)
        if self.reset_colors:
            self.selection.color = None
        return self.selection

