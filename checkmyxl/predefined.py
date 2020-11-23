from .format import check, group
from collections import Counter
from re import search, sub


def are_in(cell, iterable, sep=';', show_not_found=False, col_offset=0):
    evaluated_cell = cell.offset(0, col_offset)
    elements = str(evaluated_cell.value).split(sep)
    not_found = [element for element in elements if element not in iterable]
    if not_found and show_not_found:
        cell.value = f'{sep.join(not_found)}|{evaluated_cell.value}'
    logic_test = not not_found
    check(cell, logic_test)


def is_in(cell, iterable, col_offset=0):
    evaluated_cell = cell.offset(0, col_offset)
    logic_test = evaluated_cell.value in iterable
    check(cell, logic_test)


def is_instance(cell, instance, col_offset=0):
    evaluated_cell = cell.offset(0, col_offset)
    logic_test = isinstance(evaluated_cell.value, instance)
    check(cell, logic_test)


def is_greatest_in_row(cell, autocorrect=False, col_offset=0):
    evaluated_cell = cell.offset(0, col_offset)
    values = [x if type(x) in (int, float) else 0 for x in cell.r.value]
    logic_test = evaluated_cell.value == max(values)
    check(cell, logic_test, autocorrect=max(values)*autocorrect)


def is_unique(cell):
    duplicates = [item for item, count
                  in Counter(cell.c.value).items() if count > 1]
    logic_test = cell.value not in duplicates
    check(cell, logic_test)


def make_link(cell, col_offset=0):
    evaluated_cell = cell.offset(0, col_offset)
    logic_test = evaluated_cell.value is not None
    if logic_test:
        prefix = ''
        if not search(r'^https?://', evaluated_cell.value):
            prefix = 'https://'
        link = prefix + evaluated_cell.value
        cell.add_hyperlink(link, evaluated_cell.value)
    check(cell, logic_test, correct_color=None)


def matches_regex(cell, regex, col_offset=0):
    evaluated_cell = cell.offset(0, col_offset)
    logic_test = search(regex, str(evaluated_cell.value))
    check(cell, logic_test)


def show_groups(cell):
    group(cell)


def sub_and_group(cell, regex, replacement):
    if cell.value is None:
        cell.value = ''
    else:
        cell.value = sub(regex, replacement, str(cell.value))
    group(cell)


def translate(cell, dictionary, sep=';', show_not_found=False, col_offset=0):
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

