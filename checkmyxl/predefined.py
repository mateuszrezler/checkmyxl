GREEN = (0, 255, 0)
RED = (255, 0, 0)


def iterate(func):
    def wrapper(sheet_range, *args):
        for cell in sheet_range:
            func(cell, *args)
    return wrapper


def highlight(func):
    def wrapper(cell, *args):
        if func(cell, *args):
            cell.color = GREEN
        else:
            cell.color = RED
    return wrapper


@iterate
@highlight
def is_instance(cell, instance):
    return isinstance(cell.value, instance)

