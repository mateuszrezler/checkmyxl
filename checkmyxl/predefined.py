"""
Predefined functions for checkmyxl package.

This module provides: predefined functions for `ColumnChecker` object
that are ordered in `tasks` module.

"""
from .format import check, group
from collections import Counter
from re import search, sub


def are_in(cell, iterable, sep=';', show_not_found=False,
           not_found_prefix='Not found: ', col_offset=0):
    """
    Check if elements of splitted cell value are in the specified iterable.

    Show not found elements with prefix and / or set column offset if needed.

    Parameters
    ----------
    cell : xlwings.main.Range
        Evaluated cell containing text to be splitted.
    iterable : iterable
        Any iterable object.
    sep : str, optional
        Split separator.
    show_not_found : bool, optional
        When `True`, show not found elements in the cell.
    not_found_prefix : str, optional
        Prefix before the list of not found elements.
    col_offset : int, optional
        Relative location of the evaluated cell.

    """
    evaluated_cell = cell.offset(0, col_offset)
    elements = str(evaluated_cell.value).split(sep)
    not_found = [element for element in elements if element not in iterable]
    if not_found and show_not_found:
        cell.value = not_found_prefix + sep.join(not_found)
    logic_test = not not_found
    check(cell, logic_test)


def is_in(cell, iterable, col_offset=0):
    """
    Check if cell value is in the specified iterable.

    Set column offset if needed.
    Simplified version of `are_in` function.

    Parameters
    ----------
    cell : xlwings.main.Range
        Evaluated cell.
    iterable : iterable
        Any iterable object.
    col_offset : int, optional
        Relative location of the evaluated cell.

    """
    evaluated_cell = cell.offset(0, col_offset)
    logic_test = evaluated_cell.value in iterable
    check(cell, logic_test)


def is_instance(cell, instance, col_offset=0):
    """
    Check if the type of cell value is as specified.

    Set column offset if needed.

    Parameters
    ----------
    cell : xlwings.main.Range
        Evaluated cell.
    instance : instance
        Specified instance.
    col_offset : int
        Relative location of the evaluated cell.

    """
    evaluated_cell = cell.offset(0, col_offset)
    logic_test = isinstance(evaluated_cell.value, instance)
    check(cell, logic_test)


def is_greatest_in_row(cell, autocorrect=False, col_offset=0):
    """
    Check if cell value is the greatest in the row.

    Ignore non numeric types.
    Do autocorrection and / or set column offset if needed.

    Parameters
    ----------
    cell : xlwings.main.Range
        Evaluated cell containing a number.
    autocorrect : bool, optional
        When `True`, do autocorrection.
    col_offset : int, optional
        Relative location of the evaluated cell.

    """
    evaluated_cell = cell.offset(0, col_offset)
    values = [x if type(x) in (int, float) else 0 for x in cell.r.value]
    logic_test = evaluated_cell.value == max(values)
    check(cell, logic_test, autocorrect=max(values)*autocorrect)


def is_unique(cell):
    """
    Check if cell value is unique in the column.

    Parameters
    ----------
    cell : xlwings.main.Range
        Evaluated cell.

    """
    duplicates = [item for item, count
                  in Counter(cell.c.value).items() if count > 1]
    logic_test = cell.value not in duplicates
    check(cell, logic_test)


def make_link(cell, col_offset=0):
    """
    Add hyperlink to the cell.

    Highlight the cell in `incorrect_color` if is empty.
    Start with `https://` when such prefix is not found.
    Set column offset if needed.

    Parameters
    ----------
    cell : xlwings.main.Range
        Evaluated cell containing an url.
    col_offset : int, optional
        Relative location of the evaluated cell.

    """
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
    """
    Check if cell value matches a specified regular expression.

    Set column offset if needed.

    Parameters
    ----------
    cell : xlwings.main.Range
        Evaluated cell containing a number.
    regex : str
        Regular expression.
    col_offset : int, optional
        Relative location of the evaluated cell.

    """
    evaluated_cell = cell.offset(0, col_offset)
    logic_test = search(regex, str(evaluated_cell.value))
    check(cell, logic_test)


def show_groups(cell):
    """
    Apply the same background color to cell if the previous one contains
    the same content.

    Simplified version of `are_in` function.

    Parameters
    ----------
    cell : xlwings.main.Range
        Evaluated cell.

    """
    group(cell)


def sub_and_group(cell, regex, replacement):
    """
    Replace the regular expression matches in the cell value with
    the specified phrase. Then apply the same background color to cell
    if the previous one contains the same content.


    Parameters
    ----------
    cell : xlwings.main.Range
        Evaluated cell.
    regex : str
        Regular expression.
    replacement : str
        Replacement phrase for regular expression matches.

    """
    if cell.value is None:
        cell.value = ''
    else:
        cell.value = sub(regex, replacement, str(cell.value))
    group(cell)


def translate(cell, dictionary, sep=';', not_found_tag='<not found>',
              col_offset=0):
    """
    Translate elements of splitted cell value using specified dictionary.

    If an element is not in the dictionary, translate to `not_found_tag`.
    Set column offset if needed.

    Parameters
    ----------
    cell : xlwings.main.Range
        Evaluated cell containing text to be splitted.
    dictionary : dict
        Dictionary for translation.
    sep : str, optional
        Split separator.
    not_found_tag : str, optional
        Tag for translation of elements not found in the dictionary.
    col_offset : int, optional
        Relative location of the evaluated cell.

    """
    evaluated_cell = cell.offset(0, col_offset)
    elements = str(evaluated_cell.value).split(sep)
    translated = [str(dictionary[element]) if element in dictionary
                  else not_found_tag for element in elements]
    not_found = [element for element in elements if element not in dictionary]
    cell.value = sep.join(translated)
    logic_test = not not_found
    check(cell, logic_test)

