from checkmyxl import main
from xlwings import Book

if __name__ == '__main__':
    Book('book.xlsm').set_mock_caller()
    main()

