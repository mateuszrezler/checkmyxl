group_switch = True


def check(cell, logic_test, autocorrect=False, correct_color=(0, 255, 0),
          incorrect_color=(255, 0, 0), corrected_color=(255, 255, 0)):
    if logic_test:
        cell.color = correct_color
    else:
        if autocorrect:
            cell.color = corrected_color
        else:
            cell.color = incorrect_color


def group(cell, colors=[(153, 204, 255), None]):
    global group_switch
    if cell.value == cell.offset(-1, 0).value:
        cell.color = colors[group_switch]
    else:
        group_switch = not group_switch
        cell.color = colors[group_switch]

