from custom import contains_digit, is_bool, is_greatest_in_row, is_unique, \
    matches_regex, show_groups

TASKS = {
    0: contains_digit,
    1: is_bool,
    2: is_greatest_in_row,
    3: is_unique,
    4: (matches_regex, r'\d{3}'),
    5: show_groups
}
