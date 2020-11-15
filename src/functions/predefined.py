from .format import check, group
from collections import Counter
from re import search


def are_in(row, column, cell, iterable, sep=';', show_not_found=False,
           col_offset=0):
    evaluated_cell = cell.offset(0, col_offset)
    elements = str(evaluated_cell.value).split(sep)
    not_found = [element for element in elements if element not in iterable]
    if not_found and show_not_found:
        cell.value = f'{sep.join(not_found)}|{evaluated_cell.value}'
    logic_test = not not_found
    check(cell, logic_test)


def is_in(row, column, cell, iterable, col_offset=0):
    evaluated_cell = cell.offset(0, col_offset)
    logic_test = evaluated_cell.value in iterable
    check(cell, logic_test)


def is_instance(row, column, cell, instance, col_offset=0):
    evaluated_cell = cell.offset(0, col_offset)
    logic_test = isinstance(evaluated_cell.value, instance)
    check(cell, logic_test)


def is_greatest_in_row(row, column, cell):
    values = [x if type(x) in (int, float) else 0 for x in row.value]
    cell.value = max(values)
    logic_test = cell.value == max(values)
    check(cell, logic_test, autocorrect=True)


def is_unique(row, column, cell):
    duplicates = [item for item, count
                  in Counter(column.value).items() if count > 1]
    logic_test = cell.value not in duplicates
    check(cell, logic_test)


def matches_regex(row, column, cell, regex, col_offset=0):
    evaluated_cell = cell.offset(0, col_offset)
    logic_test = search(regex, str(evaluated_cell.value))
    check(cell, logic_test)


def show_groups(row, column, cell):
    group(cell)


def translate(row, column, cell, dictionary, sep=';', show_not_found=False,
              col_offset=0):
    evaluated_cell = cell.offset(0, col_offset)
    elements = str(evaluated_cell.value).split(sep)
    translated = [dictionary[element] for element in elements
                  if element in dictionary]
    not_found = [element for element in elements if element not in dictionary]
    if not_found and show_not_found:
        cell.value = f'{sep.join(not_found)}|{sep.join(translated)}'
    else:
        cell.value = str(sep.join(translated))
    logic_test = not not_found
    check(cell, logic_test)

