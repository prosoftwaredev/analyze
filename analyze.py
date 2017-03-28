import numpy
import xlrd
import xlsxwriter
import xlwt
import math
import random


def readData():
    workbook = xlrd.open_workbook('./data/Trade History.xlsx')
    worksheet = workbook.sheet_by_index(0)
    rows = []
    for i, col in enumerate(range(worksheet.ncols)):
        r = []
        for j, row in enumerate(range(worksheet.nrows)):
            if j <= 2:
                continue
            r.append(worksheet.cell_value(j, i))
        rows.append(r)

    return rows

def drawChart(data, length):
    workbook = xlsxwriter.Workbook('output.xlsx')
    worksheet = workbook.add_worksheet('Chart')
    worksheet.write_column('A1', data)
    chart = workbook.add_chart({'type': 'scatter'})
    chart.add_series({
        'name': 'Correlations',
        'categories': 'correlation',
        'values': '=Chart!$A$1:$A$%d' % length,
    })
    worksheet.insert_chart('D2', chart, {'x_offset': 25, 'y_offset': 10})
    workbook.close()
    return

def random_color():
    return chr(random.randrange(0, 255))

def writeData(data):
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('result')
    for i in range(len(data[0])):
        sheet.write(0, i + 1, 'SYS %d' % (i+1))
        sheet.write(i + 1, 0, 'SYS %d' % (i+1))
    for i in range(len(data)):
        for j in range(len(data[i])):
            sheet.write(j + i + 1, i + 1 ,  data[i][j])

    workbook.save('output.xlsx')

def getCor(data):
    corArray = []
    length = len(data)
    for i in range(length):
        corArray.append([])
        for j in range(i, length):
            cor = numpy.corrcoef(data[i], data[j])[0, 1]
            corArray[i].append(cor)

    return corArray
    
def main():
    data = readData()
    corArray = getCor(data)
    writeData(corArray)

if __name__ == "__main__":
    main()
