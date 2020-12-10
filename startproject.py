from subprocess import CalledProcessError, run
from sys import exit
from xlwings import App, apps, Book


def main():
    try:
        run(['xlwings', 'quickstart', 'book'], check=True)
    except FileNotFoundError:
        print ('Command not found: `xlwings`.',
               'xlwings package is not installed.',
               'Online installation guide:',
               'https://docs.xlwings.org/en/stable/installation.html.',
               sep='\n')
        exit(1)
    try:
        run(['mv', 'book/book.xlsm', 'book.xlsm'], check=True)
    except CalledProcessError:
        print('The project `book` has not been created.')
        exit(1)
    try:
        run(['rm', '-r', 'book'], check=True)
    except CalledProcessError:
        print('Temporary directory `book` cannot be removed.')
        exit(1)
    if apps.count == 0:
        App()
    Book('book.xlsm').set_mock_caller()
    book = Book.caller()
    sheets = book.sheets
    sheets.add('Sheet1.code', after='Sheet1')
    sheets.add('checkmyxl.conf', after='Sheet1.code')
    book.save()
    print('Successfully generated `book.xlsm` file.')


if __name__ == '__main__':
    main()

