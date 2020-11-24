"""
Utility functions for checkmyxl package.

This module provides: routines for loading configuration, managing sheets
and parsing command-line options.

Functions
---------
load_config        load initial configuration.
load_sheet         return active book and sheet.
make_sample        add the content of the sample file to the active sheet.
parse_args         parse command-line options.
run_from_script    routine for command-line startup.
skip_header        reduce the selection by the first row.

"""
from argparse import ArgumentParser
from os.path import join as join_path
from pandas import read_csv
from sys import path as sys_path
from xlwings import App, apps, Book


def load_config(excel_dir):
    """
    Load initial configuration.
      1. Add `config` directory to `sys.path`.
      2. Import configuration constants from `settings` module.
      3. Join paths for `EXCEL_FILE` and `SAMPLE_FILE`

    Parameters
    ----------
    excel_dir : str
        Absolute path to the directory containing Excel file.

    Returns
    -------
    excel_path : str
        Absolute path to Excel file.
    HEADER : bool
        When `True`, the first row is skipped by `ColumnChecker`.
    RESET_COLORS : bool
        When `True`, checking area background color is removed.
    sample_path : str
        Absolute path of sample `csv` file.
    """
    config_path = join_path(excel_dir, 'config')
    sys_path.append(config_path)
    from config.settings import EXCEL_FILE, HEADER, RESET_COLORS, SAMPLE_FILE
    excel_path = join_path(excel_dir, EXCEL_FILE)
    sample_path = join_path(excel_dir, *SAMPLE_FILE.split('/'))
    return excel_path, HEADER, RESET_COLORS, sample_path


def load_sheet(book_also=False):
    """
    Return active book and sheet.

    Parameters
    ----------
    book_also : bool, optional
        When `True`, a tuple of book and sheet is returned.
        When `False`, sheet is returned only.

    Returns
    -------
    book, sheet : tuple
        A tuple of active book and sheet.
    sheet : xlwings.main.Sheet
        Active sheet.

    """
    book = Book.caller()
    sheet = book.sheets.active
    return (book, sheet) if book_also else sheet


def make_sample(sample_path):
    """
    Add the content of the sample file to the active sheet.

    Parameters
    ----------
    sample_path : str
        Absolute path of sample `csv` file.

    """
    sheet = load_sheet()
    sample = read_csv(sample_path, header=None).values
    sheet['A1'].value = sample
    sheet.autofit('columns')


def parse_args(args):
    """
    Parse command-line options.

    Parameters
    ----------
    args : list
        A list of command-line options.

    Returns
    -------
    ap : argparse.ArgumentParser
        `ArgumentParser` object with parsed arguments as attributes.

    """
    ap = ArgumentParser()
    ap.add_argument('-ms', '--make-sample', action='store_true')
    return ap.parse_args(args)


def run_from_script(argv, excel_path, sample_path, excel_dir):
    """
    Routine for command-line startup.

    Parameters
    ----------
    argv : list
        A list of command-line arguments passed to the program.
    excel_path : str
        Absolute path to Excel file.
    sample_path : str
        Absolute path of sample `csv` file.
    excel_dir : str
        Absolute path to the directory containing Excel file.

    """
    if apps.count == 0:
        App()
    Book(excel_path).set_mock_caller()
    if len(argv) > 1:
        args = parse_args(argv[1:])
        if args.make_sample:
            make_sample(sample_path)


def skip_header(sheet, selection):
    """
    Reduce the selection by the first row.

    Parameters
    ----------
    sheet : xlwings.main.Sheet
        Active sheet containing a selection.
    selection : xlwings.main.Range
        Selected region in the active sheet.

    Returns
    -------
    reduced_selection : xlwings.main.Range
        Selected region minus the first line.

    """
    coordinates = (selection.row+1, selection.column), \
        (selection.last_cell.row, selection.last_cell.column)
    reduced_selection = sheet.range(*coordinates)
    return reduced_selection

