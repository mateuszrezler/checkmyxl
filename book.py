"""
Startup script for checkmyxl package.

For proper startup, the file system must be organized in such a way:
```
    config - configuration directory
        custom.py - user-defined functions
        settings.py - startup settings
        tasks.py - pipeline of tasks
    <name>.xlsm - Excel file to be checked
    <name>.py - this file (name must be the same as the name of the Excel file)
```

This module provides: functions called directly from the associated Excel file
containing `xlwings` ribbon and / or vba macros.

Functions
---------
main    initialize checkmyxl, open Excel file and do the check.
undo    reopen the active book without saving it.

"""
from checkmyxl import ColumnChecker
from checkmyxl.utils import load_config, load_sheet, run_from_script
from os.path import dirname, realpath
from sys import argv, path as sys_path
from xlwings import Book


def main(argv=[], selection=None):
    """
    Initialize checkmyxl, open Excel file and do the check.

    argv : list, optional
        A list of command-line arguments passed to the program.
        This list is empty when `main` function is called directly from Excel.
    selection : str, optional
        Selected range in `<letters><digits>(:<letters><digits)` format.

    """
    if argv:
        run_from_script(argv, excel_path, sample_path)
    sheet = load_sheet()
    cc = ColumnChecker(sheet, selection, header, reset_colors)
    cc.check()


def undo():
    """
    Reopen the active book without saving it.

    A workaround solution for inability to undo changes made by scripts.

    """
    book, sheet = load_sheet(book_also=True)
    book.close()
    Book(excel_path)


excel_dir = dirname(realpath(__file__))
excel_path, header, reset_colors, sample_path = load_config(excel_dir)
if __name__ == '__main__':
    main(argv)

