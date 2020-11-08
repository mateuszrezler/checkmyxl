from collections import Counter
from format import highlight
from re import search


def contains_digit(row, column, cell):
    logic_test = search(r'\d', str(cell.value))
    highlight(cell, logic_test)


def is_instance(row, column, cell, instance):
    logic_test = isinstance(cell.value, instance)
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


def matches_regex(row, column, cell, regex):
    logic_test = search(regex, str(cell.value))
    highlight(cell, logic_test)


def show_groups(row, column, cell):
    group = 0
    logic_test = cell.value == cell.offset(-1, 0).value
    group = highlight(cell, logic_test, groups=True, group=group)

