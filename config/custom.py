"""
User defined functions for checkmyxl package.

This module provides: user defined functions for `ColumnChecker` object
that are ordered in `tasks` module.

Functions
---------
circle_area    sample function.

"""
from checkmyxl.format import check
from math import pi


def is_area_above_threshold(cell, threshold=100.0):
    """
    Sample function.

    Check if the area of ​​the circle with a given radius exceeds the threshold.

    Parameters
    ----------
    cell : xlwings.main.Range
        Evaluated cell containing radius value.
    threshold : float, optional
        Threshold, set to `100.0` by default.

    """
    radius = cell.value
    if type(radius) in (type(None), str):
        radius = 0.0
    area = pi*radius**2
    logic_test = area > threshold
    check(cell, logic_test)

