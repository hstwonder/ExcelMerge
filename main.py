# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

import openpyxl
from openpyxl import Workbook


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press ⌘F8 to toggle the breakpoint.


# See PyCharm help at https://www.jetbrains.com/help/pycharm/


def load_excel_file(filename):
    book = openpyxl.load_workbook(filename)
    sheet = book.active
    sheet.guess_types = True
    print(sheet['A1'].value)

    for row in sheet.iter_rows(min_row=1, min_col=1):
        for cell in row:
            print(cell.value, end=" ")
        print()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')
    load_excel_file('1.xlsx')

