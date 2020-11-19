"""
Utility functions for checkmyxl package.

This module provides: routines for loading files and managing sheets.

Functions
---------
get_abs_path    return absolute path of a given relative path.
load_config     read `config.json` file and return configuration dictionary.
load_sample     read sample `csv` file and return its contents as data frame.
load_sheet      return active book and sheet.
skip_header     reduce the selection by the first row.

"""
from argparse import ArgumentParser
from json import load as load_json
from os.path import dirname, join as join_path, realpath
from pandas import read_csv
from xlwings import Book


def get_abs_path(rel_path):
    """
    Return absolute path of a given relative path.

    Parameters
    ----------
    rel_path : str
        Path relative to `checkmyxl.py` file.

    Returns
    -------
    abs_path : str
        Absolute path.

    """
    parent_dir = dirname(realpath(__file__))
    abs_path = join_path(parent_dir, '..', *rel_path.split('/'))
    return abs_path


def load_config():
    """
    Read `config.json` file and return configuration dictionary.

    Returns
    -------
    config : dict
        A dictionary with initial settings.

    """
    config_path = get_abs_path('config.json')
    with open(config_path) as cf:
        config = load_json(cf)
    return config


def load_sample():
    """
    Read sample `csv` file and return its contents as data frame.

    Returns
    -------
    sample : pandas.core.frame.DataFrame
        Sample data frame.

    """
    config = load_config()
    sample_path = get_abs_path(config['sample_file'])
    sample = read_csv(sample_path, header=None).values
    return sample


def load_sheet():
    """
    Return active book and sheet.

    Returns
    -------
    book, sheet : tuple
        A tuple of active book and sheet.

    """
    book = Book.caller()
    sheet = book.sheets.active
    return book, sheet


def parse_args(args):
    ap = ArgumentParser()
    ap.add_argument('-ms', '--make-sample', action='store_true')
    return ap.parse_args(args)


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

