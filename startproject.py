from argparse import ArgumentParser
from pandas import read_csv
from subprocess import CalledProcessError, run
from sys import exit
from xlwings import App, apps, Book


def main():
    print('checkmyxl project generator')
    parser = ArgumentParser(add_help=False)
    parser.add_argument('-s', '--sample', action='store_true')
    args = parser.parse_args()
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
    if args.sample:
        sheet1, sheet1_code, checkmyxl_conf = \
            sheets['Sheet1'], sheets['Sheet1.code'], sheets['checkmyxl.conf']
        sheet1['A1'].options(index=False).value = \
            read_csv('sample/sample.csv')
        sheet1_code['A1'].options(index=False).value = \
            read_csv('sample/sample_code.csv')
        checkmyxl_conf['A1'].options(index=False).value = \
            read_csv('sample/sample_conf.csv')
        print('Sample added.')


if __name__ == '__main__':
    main()

