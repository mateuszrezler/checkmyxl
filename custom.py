from collections import Counter
from format import highlight
from re import search


def contains_digit(row, column, cell):
    logic_test = search(r'\d', str(cell.value))
    highlight(cell, logic_test)


def is_bool(row, column, cell):
    logic_test = isinstance(cell.value, bool)
    highlight(cell, logic_test)


def is_greatest_in_row(row, column, cell):
    logic_test = cell.value == max(row.value)
    cell.value = max(row.value)
    highlight(cell, logic_test, autocorrect=True)


def is_unique(row, column, cell):
    duplicates = [item for item, count
                  in Counter(column.value).items() if count > 1]
    logic_test = cell.value not in duplicates
    highlight(cell, logic_test)


def show_groups(row, column, cell):
    colors = [(128, 128, 128), None]
    logic_test = cell.value == cell.offset(-1, 0).value
    if cell.offset(-1, 0).color == colors[0]:
        cell.color = colors[not logic_test]
    else:
        cell.color = colors[logic_test]
