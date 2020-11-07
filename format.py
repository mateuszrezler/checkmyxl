def highlight(cell,
              logic_test,
              autocorrect=False,
              groups=False,
              correct_color=(0, 255, 0),
              incorrect_color=(255, 0, 0),
              corrected_color=(255, 255, 0),
              group_colors=[(153, 204, 255), None],
              group=0):
    if logic_test:
        if groups:
            cell.color = group_colors[group]
            return group
        else:
            cell.color = correct_color
    else:
        if autocorrect:
            cell.color = corrected_color
        else:
            cell.color = incorrect_color
        if groups:
            cell.color = group_colors[0**group]
            return group

