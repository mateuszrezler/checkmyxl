from checkmyxl import ColumnChecker, make_sample
from checkmyxl.utils import get_abs_path, load_config, load_sheet, \
    parse_args, skip_header
from sys import argv
from xlwings import App, apps, Book


def main(selection=None):
    config = load_config()
    book, sheet = load_sheet()
    if selection:
        header = False
        selection = sheet[selection]
    else:
        selection = sheet.used_range
        if config['header']:
            selection = skip_header(sheet, selection)
    if config['reset_colors']:
        selection.color = None
    cc = ColumnChecker(sheet, selection)
    cc.check()


def start(argv):
    config = load_config()
    excel_path = get_abs_path(config['excel_file'])
    if apps.count == 0:
        App()
    Book(excel_path).set_mock_caller()
    if len(argv) > 1:
        args = parse_args(argv[1:])
        if args.make_sample:
            make_sample()
    main()


if __name__ == '__main__':
    start(argv)

