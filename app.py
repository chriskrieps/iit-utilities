from ctypes import wstring_at
from openpyxl import Workbook, load_workbook
import datetime

def main():
    wb = load_workbook(filename = '.\Data\Meter Data Summary.xlsx')

    for ws in wb:
        print(ws.title)
        
    ws = wb['TN Elec']

    index = 0


    """ for i, cell in enumerate(ws['B']):
        print(cell.number_format)
        print(cell.value)
        index = i

    print('Number of rows: '+str(i)) """

if __name__ == "__main__":
    main()