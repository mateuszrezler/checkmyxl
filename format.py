def highlight(cell, logic_test):
    if logic_test:
        cell.color = (0, 255, 0)
    else:
        cell.color = (255, 0, 0)
