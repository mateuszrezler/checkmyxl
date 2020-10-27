from collections import Counter
from format import highlight
from re import search


def contains_digit(selection, cell):
    logic_test = search(r'\d', str(cell.value))
    highlight(cell, logic_test)


def has_no_duplicates(selection, cell):
    duplicates = [item for item, count
                  in Counter(selection.value).items() if count > 1]
    logic_test = cell.value not in duplicates
    highlight(cell, logic_test)
