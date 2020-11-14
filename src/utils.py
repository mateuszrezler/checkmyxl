from json import load as load_json
from os.path import dirname, join as join_path, realpath
from pandas import read_csv
from xlwings import Book


def get_abs_path(rel_path):
    parent_dir = dirname(realpath(__file__))
    return join_path(parent_dir, '..', *rel_path.split('/'))


def load_config():
    config_path = get_abs_path('config.json')
    with open(config_path) as cf:
        config = load_json(cf)
    return config


def load_sample():
    config = load_config()
    sample_path = get_abs_path(config['sample_file'])
    sample = read_csv(sample_path, header=None).values
    return sample


def load_sheet():
    book = Book.caller()
    sheet = book.sheets.active
    return book, sheet


def skip_header(sheet, selection):
    return sheet.range(
        (selection.row+1, selection.column),
        (selection.last_cell.row, selection.last_cell.column))

