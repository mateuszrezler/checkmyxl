from sys import path as sys_path

sys_path.append('..')

from book import main, undo
from checkmyxl.utils import load_sheet
from xlwings import apps


GREEN = (187, 255, 187)


def quit_app():
    for app in apps:
        app.quit()


def test_make_sample():
    quit_app()
    main(['book.py', '--make-sample'])
    sheet = load_sheet()
    assert sheet['A1'].value == "are_in(\n    iterable=['ant', 'bat']\n)"


def test_selection():
    main(['book.py'], selection='A2:B13')
    sheet = load_sheet()
    assert sheet['A2'].color == GREEN


def test_undo():
    undo()
    quit_app()

