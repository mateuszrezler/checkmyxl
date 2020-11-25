"""
Core module of checkmyxl package.

This module provides: `ColumnChecker` class.

Classes
-------
ColumnChecker    engine of the checking process.

"""
from checkmyxl.utils import skip_header
from config.tasks import TASKS


class ColumnChecker(object):
    """
    Engine of the checking process.

    Checking columns as specified in `TASKS` dictionary.

    Parameters
    ----------
    sheet : xlwings.main.Sheet
        Active sheet.
    selection : xlwings.main.Range
        Selected region in the active sheet.
    header : bool
        When `True`, the first row is skipped by `ColumnChecker`.
    reset_colors : bool
        When `True`, checking area background color is removed.

    """
    def __init__(self, sheet, selection, header, reset_colors):
        self.sheet = sheet
        self.selection = selection
        self.header = header
        self.reset_colors = reset_colors
        self.selection = self.set_selection()

    def check(self):
        """
        For each column do a specific task, row by row.

        Skip columns with indices not found in `TASKS` dictionary.
        If task is a function then call it with `cell` as argument.
        When task is a tuple of funtion and dictionary then call this function
        with `cell` as positional argument and unpacked dictionary as keyword
        arguments.

        """
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
        """
        Preprocess selection before run.

        Set selection as is, without header if selection is passed as argument.
        Otherwise set selection as used range and skip header if needed.
        Reset colors of set selection if `reset_colors` is `True`.

        """
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

