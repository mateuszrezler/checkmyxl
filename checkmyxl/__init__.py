"""
checkmyxl
=========

Automate validation of your data in Microsoft Excel sheet.

This package provides
  1. Checking user-defined correctness of cells, including autocorrection.
  2. Showing groups of cells with identical content.
  3. Any custom operation available in `xlwings` package.

How to use the documentation
----------------------------
Documentation is available in the form of docstrings provided with the code.
Use the built-in `help` function to view a function's docstring.

"""
from .core import ColumnChecker


__name__ = 'checkmyxl'
__author__ = 'Mateusz Rezler'
__copyright__ = 'Mateusz Rezler'
__license__ = 'MIT'

