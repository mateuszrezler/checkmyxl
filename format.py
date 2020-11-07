def highlight(cell, logic_test, autocorrect=False):
    if logic_test:
        cell.color = (0, 255, 0)
    else:
        if autocorrect:
            cell.color = (255, 255, 0)
        else:
            cell.color = (255, 0, 0)

