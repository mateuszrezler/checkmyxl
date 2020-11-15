from .format import check, group
from collections import Counter
from re import search


def are_in(row, column, cell, iterable, sep=';', col_offset=0):
    eval_cell = cell.offset(0, col_offset)
    elements = str(eval_cell.value).split(sep)
    not_found = [element for element in elements if element not in iterable]
    logic_test = not not_found
    check(cell, logic_test)


def is_in(row, column, cell, iterable, col_offset=0):
    eval_cell = cell.offset(0, col_offset)
    logic_test = eval_cell.value in iterable
    check(cell, logic_test)


def is_instance(row, column, cell, instance, col_offset=0):
    eval_cell = cell.offset(0, col_offset)
    logic_test = isinstance(eval_cell.value, instance)
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
    eval_cell = cell.offset(0, col_offset)
    logic_test = search(regex, str(eval_cell.value))
    check(cell, logic_test)


def show_groups(row, column, cell):
    group(cell)

