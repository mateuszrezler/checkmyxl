from checkmyxl import ColumnChecker
from checkmyxl.utils import load_config, load_sheet, run_from_script
from os.path import dirname, realpath
from sys import argv, path as sys_path
from xlwings import Book


def main(argv=[], selection=None):
    if argv:
        run_from_script(argv, excel_path, sample_path)
    sheet = load_sheet()
    cc = ColumnChecker(sheet, selection, header, reset_colors)
    cc.check()


def undo():
    book, sheet = load_sheet(book_also=True)
    book.close()
    Book(excel_path)


excel_dir = dirname(realpath(__file__))
excel_path, header, reset_colors, sample_path = load_config(excel_dir)
if __name__ == '__main__':
    main(argv)

