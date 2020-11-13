from .format import check, group
from collections import Counter
from re import search


def are_in(row, column, cell, iterable, sep=';', col_offset=0):
    elements = str(cell.offset(0, col_offset).value).split(sep)
    logic_test = [element for element in elements if element in iterable]
    check(cell, logic_test)


def is_in(row, column, cell, iterable, col_offset=0):
    logic_test = cell.offset(0, col_offset).value in iterable
    check(cell, logic_test)


def is_instance(row, column, cell, instance, col_offset=0):
    logic_test = isinstance(cell.offset(0, col_offset).value, instance)
    check(cell, logic_test)


def is_greatest_in_row(row, column, cell):
    values = [x if type(x) in (int, float) else 0 for x in row.value]
    logic_test = cell.value == max(values)
    cell.value = max(values)
    check(cell, logic_test, autocorrect=True)


def is_unique(row, column, cell):
    duplicates = [item for item, count
                  in Counter(column.value).items() if count > 1]
    logic_test = cell.value not in duplicates
    check(cell, logic_test)


def matches_regex(row, column, cell, regex, col_offset=0):
    logic_test = search(regex, str(cell.offset(0, col_offset).value))
    check(cell, logic_test)


def show_groups(row, column, cell):
    group(cell)

