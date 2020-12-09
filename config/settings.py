"""
Startup settings for checkmyxl package.

This module provides: constants for `utils.load_config` function.

Constants
---------
EXCEL_FILE      Name of Excel file.
HEADER          When `True`, checking area background color is removed.
RESET_COLORS    When `True`, the first row is skipped by `ColumnChecker`.
SAMPLE_FILE     Name of sample `csv` file.

"""
EXCEL_FILE = 'book.xlsm'
HEADER = True
PALETTE = {
    'blue': (187, 221, 255),
    'green': (187, 255, 187),
    'orange': (255, 221, 187),
    'red': (255, 187, 187),
    'violet': (187, 187, 221),
    'yellow': (255, 255, 187)
}
RESET_COLORS = True
SAMPLE_FILE = 'data/sample.csv'

