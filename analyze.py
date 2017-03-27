import numpy
import xlrd
import math


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

def correlation(x, y):
    sumX = 0
    sumY = 0
    sumSquareX = 0
    sumSquareY = 0
    sumMulti = 0
    n = len(x)
    for idx in range(n):
        sumX += x[idx]
        sumY += y[idx]
        sumSquareX += math.pow(x[idx], 2)
        sumSquareY += math.pow(y[idx], 2)
        sumMulti += x[idx] * y[idx]
    return (n * sumMulti - sumX * sumY) / math.sqrt((n * sumSquareX - math.pow(sumX, 2) ) * (n * sumSquareY - math.pow(sumY, 2) ))

def findMiniCor(data):
    minIdx = []
    minIdx.append(0)
    corArray = []
    for index in range(len(data) - 1):
        # corArray.append(numpy.corrcoef(data[index], data[index + 1])[0, 1])
        corArray.append(correlation(data[index], data[index + 1]))
        if corArray[minIdx[0]] > corArray[index]:
            del minIdx[:]
            minIdx.append(index)
        elif corArray[minIdx[0]] == corArray[index]:
            minIdx.append(index)
    return minIdx, corArray
    
def main():
    data = readData()
    minIdx, corArray = findMiniCor(data)
    print minIdx
    print corArray[160]

if __name__ == "__main__":
    main()
