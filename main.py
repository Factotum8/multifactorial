# -*- encoding: utf-8 -*-
import sys

from openpyxl import load_workbook

def main():
    wb = load_workbook('./data_matrix.xlsx')
    pm = wb.get_sheet_by_name(wb.get_sheet_names()[0]) # planning matrix
    print (pm.title)
    for i in range(3, 11):
        print(pm.cell(row=i, column=1).value, pm.cell(row=i, column=3).value)

if __name__ == '__main__':
    main()