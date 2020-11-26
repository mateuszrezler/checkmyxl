from sys import path as sys_path

sys_path.append('..')

from book import main, undo
from xlwings import apps


def quit_app():
    for app in apps:
        app.quit()


def test_sample():
    quit_app()
    main(['book.py', '--make-sample'])
    main(['book.py'], selection='A2:B13')


def test_undo():
    undo()
    quit_app()

