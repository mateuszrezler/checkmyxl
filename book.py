from checkmyxl import main
from xlwings import App, apps, Book

if __name__ == '__main__':
    if apps.count == 0:
        App()
    Book('book.xlsm').set_mock_caller()
    main()

