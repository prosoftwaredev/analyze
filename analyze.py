# import numpy
import xlrd


def readData():
    workbook = xlrd.open_workbook('./data/Trade History.xlsx')
    worksheet = workbook.sheet_by_index(0)
    rows = []
    for i, row in enumerate(range(worksheet.nrows)):
        if i <= 1:
            continue
        r = []
        for j, col in enumerate(range(worksheet.ncols)):
            r.append(worksheet.cell_value(i, j))
        rows.append(r)

    return rows

    
def main():
    data = readData()
    print data

if __name__ == "__main__":
    main()
