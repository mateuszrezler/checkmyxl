"""
Formatting functions for checkmyxl package.

This module provides: functions changing appearance of a sheet.

Functions
---------
check    do a logic test and change the background color of a cell.
group    apply the same background color to cell if the previous one contains
         the same content.

"""
from config.settings import PALETTE


group_switch = True


def check(cell, logic_test, autocorrect=False, correct_color=PALETTE['green'],
          incorrect_color=PALETTE['red'], corrected_color=PALETTE['yellow']):
    """
    Do a logic test and change the background color of a cell.

    Parameters
    ----------
    cell : xlwings.main.Range
        Evaluated cell.
    logic_test : object
        Logic test.
    autocorrect : object, optional
        When considered `True`, cell background color is changed to
        `corrected_color` and its value is changed to `autocorrect`.
        When `False`, autocorrection is not applied.
    correct_color : tuple, optional
        A tuple of RGB values for cell color when `logic_test is True`.
        Green by default.
    incorrect_color : tuple, optional
        A tuple of RGB values for cell color when `logic_test is False`
        and `autocorrect is False`.
        Red by default.
    corrected_color : tuple, optional
        A tuple of RGB values for cell color when `logic_test is False`
        and `autocorrect` is considered `True`.
        Yellow by default.

    """
    if logic_test:
        cell.color = correct_color
    else:
        if autocorrect:
            cell.color = corrected_color
            cell.value = autocorrect
        else:
            cell.color = incorrect_color


def group(cell, colors=[PALETTE['blue'], None]):
    """
    Apply the same background color to cell if the previous one contains
    the same content.

    Parameters
    ----------
    cell : xlwings.main.Range
        Evaluated cell.
    colors : list, optional
        A pair of colors to apply.
        Light blue and no background by default.

    """
    global group_switch
    if cell.value == cell.offset(-1, 0).value:
        cell.color = colors[group_switch]
    else:
        group_switch = not group_switch
        cell.color = colors[group_switch]

