"""
Tasks for checkmyxl package.

This module provides: `TASKS` dictionary for `ColumnChecker` object.

Constants
---------
TASKS    dictionary containing functions from `predefined` or `custom` module
         and their parameters if needed.

"""
from checkmyxl.functions.predefined import is_instance, is_greatest_in_row, \
    is_unique, matches_regex, show_groups


TASKS = {
    0: (is_instance, {'instance': bool}),
    1: is_greatest_in_row,
    2: is_unique,
    3: (matches_regex, {'regex': r'\d{3}'}),
    4: show_groups
}

