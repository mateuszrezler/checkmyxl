from .format import highlight
from collections import Counter
from re import search


group = 0


def is_in(row, column, cell, iterable, col_offset=0):
    logic_test = cell.offset(0, col_offset).value in iterable
    highlight(cell, logic_test)


def is_instance(row, column, cell, instance, col_offset=0):
    logic_test = isinstance(cell.offset(0, col_offset).value, instance)
    highlight(cell, logic_test)


def is_greatest_in_row(row, column, cell):
    values = [x if type(x) in (int, float) else 0 for x in row.value]
    logic_test = cell.value == max(values)
    cell.value = max(values)
    highlight(cell, logic_test, autocorrect=True)


def is_unique(row, column, cell):
    duplicates = [item for item, count
                  in Counter(column.value).items() if count > 1]
    logic_test = cell.value not in duplicates
    highlight(cell, logic_test)


def matches_regex(row, column, cell, regex, col_offset=0):
    logic_test = search(regex, str(cell.offset(0, col_offset).value))
    highlight(cell, logic_test)


def show_groups(row, column, cell):
    global group
    logic_test = cell.value == cell.offset(-1, 0).value
    group = highlight(cell, logic_test, groups=True, group=group)

