GREEN = (0, 255, 0)
RED = (255, 0, 0)


def iterate(function):
    def wrapper(sheet_range, *args, **kwargs):
        for cell in sheet_range:
            function(cell, *args, **kwargs)
    return wrapper


def highlight(function):
    def wrapper(cell, *args, correct_color=GREEN, incorrect_color=RED,
                **kwargs):
        result = function(cell, *args, **kwargs)
        cell.color = (incorrect_color, correct_color)[int(result)]
        return result
    return wrapper


@iterate
@highlight
def is_instance(cell, instance):
    return isinstance(cell.value, instance)

