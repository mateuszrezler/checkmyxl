"""
Tasks for checkmyxl package.

This module provides: `TASKS` dictionary for `ColumnChecker` object.

Constants
---------
TASKS    dictionary containing functions from `predefined` or `custom` module
         and their parameters if needed.

"""
from checkmyxl.predefined import are_in, is_in, is_instance, \
    is_greatest_in_row, is_unique, make_link, matches_regex, show_groups, \
    sub_and_group, translate
from config.custom import is_area_above_threshold
from config.settings import PALETTE


TASKS = {
    0: (are_in, {'iterable': ['ant', 'bat']}),
    1: (are_in, {'iterable': ['cat', 'dog'],
                 'sep': ';',
                 'show_not_found': True,
                 'not_found_prefix': 'Not found: ',
                 'col_offset': -1}),
    2: (is_in, {'iterable': [True, 1]}),
    3: (is_in, {'iterable': (None, False, 0, 'NA'),
                'col_offset': -1}),
    4: (is_instance, {'instance': float}),
    5: (is_instance, {'instance': bool,
                      'col_offset': -3}),
    6: is_greatest_in_row,
    7: (is_greatest_in_row, {'autocorrect': True,
                             'col_offset': -1}),
    8: is_unique,
    9: (make_link, {'col_offset': 1}),
    10: make_link,
    11: (matches_regex, {'regex': r'A\d$'}),
    12: (matches_regex, {'regex': r'\d{3}',
                         'col_offset': -8}),
    13: show_groups,
    14: (show_groups, {'colors': [PALETTE['orange'], PALETTE['yellow']]}),
    15: (sub_and_group, {'regex': r'0*\.0',
                         'replacement': '000'}),
    16: (sub_and_group, {'regex': r'[A-Z]',
                         'replacement': lambda match: match.group(0).lower(),
                         'colors': [PALETTE['orange'], PALETTE['yellow']]}),
    17: (translate, {'dictionary': {'ant': 0, 'bat': 1},
                     'sep': ', '}),
    18: (translate, {'dictionary': {'cat': True, 'dog': False},
                    'not_found_tag': '<?>',
                    'col_offset': -18}),
    19: is_area_above_threshold,
    20: (is_area_above_threshold, {'threshold': 10})
}

